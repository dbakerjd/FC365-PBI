import { Inject, Component, OnInit } from '@angular/core';
import { UploadFileConfig } from '@shared/forms/upload-file.config';
import { MatDialog, MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { FormControl, FormGroup } from '@angular/forms';
import { SharepointService } from 'src/app/services/sharepoint.service';
import { Observable } from 'rxjs';
import { take } from 'rxjs/operators';
import { ConfirmDialogComponent } from '../confirm-dialog/confirm-dialog.component';
import { NPPFolder } from '@shared/models/file-system';
import { Indication } from '@shared/models/entity';

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
  businessUnitId: number = 0;
  forecastCycleId: number = 0;
  geoId: number = 0;

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<UploadFileComponent>,
    public matDialog: MatDialog,
    private readonly sharepoint: SharepointService,
  ) { 
    this.businessUnitId = this.data.entity.BusinessUnitId ? this.data.entity.BusinessUnitId : 0;
    this.forecastCycleId = this.data.entity.ForecastCycleId ? this.data.entity.ForecastCycleId : 0;
  }

  ngOnInit(): void {
    this.formConfig = new UploadFileConfig();
    this.fields = this.formConfig.fields(
      this.data.entity.Id, 
      this.data.masterStageId, 
      this.data.folderList,
      this.data.selectedFolder,
      this.data.geographies,
      this.data.scenarios,
      this.data.entity.Indication?.map((el: Indication) => {
        return { label: el.Title, value: el.ID }
      }));
    this.form = new FormGroup({});
  }

  async onSubmit() {
    if (this.form.invalid) {
      this.validateAllFormFields(this.form);
      return;
    }

    const user = await this.sharepoint.getCurrentUserInfo();
    let userName = user.Title && user.Title.indexOf("@") == -1 ? user.Title : user.Email;

    let fileData = {
      StageNameId: this.model.StageNameId,
      EntityNameId: this.model.EntityNameId,
      Comments: this.model.description ? '[{"text":"'+this.model.description.replace(/'/g, "{COMMA}")+'","email":"'+user.Email+'","name": "'+ userName +'","userId":'+user.Id+',"createdAt":"'+new Date().toISOString()+'"}]' : '[]',
    };

    this.uploading = this.dialogRef.disableClose = true;

    let fileFolder = '/' + this.businessUnitId + '/' + this.model.EntityNameId + '/' + this.model.StageNameId + '/' + this.model.category;
    
    if (this.data.folderList.find((f: NPPFolder) => f.DepartmentID == this.model.category).containsModels) {
      // add geography to folder route
      fileFolder += '/' + this.model.geography + '/0';

      // read opp geography to get master ID of country / geography
      const oppGeographies = await this.sharepoint.getOpportunityGeographies(this.model.EntityNameId);
      const geography = oppGeographies.find(el => el.Id == this.model.geography);
      const user = await this.sharepoint.getCurrentUserInfo();
      let userName = user.Title && user.Title.indexOf("@") == -1 ? user.Title : user.Email;

      Object.assign(fileData, {
        EntityGeographyId: geography.Id ? geography.Id : null,
        ModelScenarioId: this.model.scenario,
        Comments: this.model.description ? '[{"text":"'+this.model.description.replace(/'/g, "{COMMA}")+'","email":"'+user.Email+'","name": "'+ userName +'","userId":'+user.Id+',"createdAt":"'+new Date().toISOString()+'"}]' : '[]',
        ApprovalStatusId: await this.sharepoint.getApprovalStatusId("In Progress"),
        IndicationId: this.model.IndicationId
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
      fileFolder = fileFolder + '/0/0';
      // regular file
      Object.assign(fileData, {
        Comments: this.model.description
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
