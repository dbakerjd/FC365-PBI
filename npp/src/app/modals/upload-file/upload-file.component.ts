import { Inject, Component, OnInit } from '@angular/core';
import { UploadFileConfig } from 'src/app/shared/forms/upload-file.config';
import { MatDialog, MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { FormControl, FormGroup } from '@angular/forms';
import { NPPFolder, SharepointService } from 'src/app/services/sharepoint.service';
import { Observable } from 'rxjs';
import { take } from 'rxjs/operators';
import { ConfirmDialogComponent } from '../confirm-dialog/confirm-dialog.component';

@Component({
  selector: 'app-upload-file',
  templateUrl: './upload-file.component.html',
  styleUrls: ['./upload-file.component.scss']
})
export class UploadFileComponent implements OnInit {
  formConfig: UploadFileConfig = new UploadFileConfig();
  fields: FormlyFieldConfig[] = [];
  form: FormGroup = new FormGroup({});
  model: any = { };
  folders: NPPFolder[] = [];
  uploading = false; // spinner control

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<UploadFileComponent>,
    public matDialog: MatDialog,
    private readonly sharepoint: SharepointService,
  ) { 
    
  }

  ngOnInit(): void {
    this.formConfig = new UploadFileConfig();
    this.fields = this.formConfig.fields(
      this.data.opportunityId, 
      this.data.masterStageId, 
      this.data.folderList,
      this.data.selectedFolder,
      this.data.geographies,
      this.data.scenarios);
    this.form = new FormGroup({});
  }

  async onSubmit() {
    if (this.form.invalid) {
      this.validateAllFormFields(this.form);
      return;
    }
    let fileData = {
      StageNameId: this.model.StageNameId,
      OpportunityNameId: this.model.OpportunityNameId,
    };

    this.uploading = this.dialogRef.disableClose = true;

    let fileFolder = '/' + this.model.OpportunityNameId + '/' + this.model.StageNameId + '/' + this.model.category;
    if (this.data.folderList.find((f: NPPFolder) => f.ID == this.model.category).containsModels) {
      // add geography to folder route
      fileFolder += '/' + this.model.geography;

      // read opp geography to get master ID of country / geography
      const oppGeographies = await this.sharepoint.getOpportunityGeographies(this.model.OpportunityNameId);
      const geography = oppGeographies.find(el => el.Id == this.model.geography);

      Object.assign(fileData, {
        OpportunityGeographyId: geography.Id ? geography.Id : null,
        ModelScenarioId: this.model.scenario,
        ModelApprovalComments: this.model.description,
        ApprovalStatusId: await this.sharepoint.getApprovalStatusId("In Progress"),
      });
      let scenarioFileName = this.model.file[0].name.replace(/[~#%&*{}:<>?+|"/\\]/g, "");

      if (await this.sharepoint.existsFile(scenarioFileName, this.sharepoint.getBaseFilesFolder() + fileFolder)) {
        const dialogRef = this.matDialog.open(ConfirmDialogComponent, {
          maxWidth: "400px",
          height: "200px",
          data: {
            message: `A model with this name (${scenarioFileName}) already exists in this location. Do you want to overwrite it?`,
            confirmButtonText: 'Yes, overwrite',
            cancelButtonText: 'No, keep the original'
          }
        });
      
        dialogRef.afterClosed()
          .pipe(take(1))
          .subscribe(async uploadConfirmed => {
            if (uploadConfirmed) {
              this.uploadFileToFolder(fileData, scenarioFileName, this.sharepoint.getBaseFilesFolder() + fileFolder);
            } else {
              // do nothing and close
              this.dialogRef.close({
                success: true, 
                uploaded: false
              });
            }
          });
      } else {
        this.uploadFileToFolder(fileData, scenarioFileName, this.sharepoint.getBaseFilesFolder() + fileFolder);
      }
      
    } else {
      // regular file
      Object.assign(fileData, {
        ModelApprovalComments: this.model.description
      });
      const cleanFilename = this.model.file[0].name.replace(/[~#%&*{}:<>?+|"/\\]/g, "");
      if (await this.sharepoint.existsFile(cleanFilename, this.sharepoint.getBaseFilesFolder() + fileFolder)) {
        const dialogRef = this.matDialog.open(ConfirmDialogComponent, {
          maxWidth: "400px",
          height: "200px",
          data: {
            message: `A file with this name (${cleanFilename}) already exists inside the folder. Do you want to overwrite it?`,
            confirmButtonText: 'Yes, overwrite',
            cancelButtonText: 'No, keep the original'
          }
        });
    
        dialogRef.afterClosed()
          .pipe(take(1))
          .subscribe(async uploadConfirmed => {
            if (uploadConfirmed) {
              this.uploadFileToFolder(fileData, cleanFilename, this.sharepoint.getBaseFilesFolder() + fileFolder);
            } else {
              this.dialogRef.close({
                success: true, 
                uploaded: false
              });
            }
          });
      } else {
        this.uploadFileToFolder(fileData, cleanFilename, this.sharepoint.getBaseFilesFolder() + fileFolder);
      }
      
    }
  }

  private async uploadFileToFolder(fileData: any, fileName: string, folder: string) {
    this.readFileDataAsText(this.model.file[0]).subscribe(
      data => {
        this.sharepoint.uploadFile(data, folder, fileName, fileData).then(
          r => { 
            if (Object.keys(r).length > 0) {
              this.uploading = this.dialogRef.disableClose = false; // finished

              this.dialogRef.close({
                success: true, 
                uploaded: true,
                name: fileName
              });
            }
            else {
              this.dialogRef.close({
                success: false,
              });
            }
          }
        );
      }
    );
  }

  private readFileDataAsText(file: any): Observable<string> {
    return new Observable(obs => {
      const reader = new FileReader();
      reader.onloadend = () => {
        obs.next(reader.result as string);
        obs.complete();
      }
      reader.readAsArrayBuffer(file);
    });
  }

  private validateAllFormFields(formGroup: FormGroup) {
    Object.keys(formGroup.controls).forEach(field => {
      const control = formGroup.get(field);
      if (control instanceof FormControl) {
        control.markAsTouched({ onlySelf: true });
        control.markAsDirty({ onlySelf: true });
      } else if (control instanceof FormGroup) {
        this.validateAllFormFields(control);
      }
    });
  }
}
