import { Inject, Component, OnInit } from '@angular/core';
import { UploadFileConfig } from 'src/app/shared/forms/upload-file.config';
import { MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { FormControl, FormGroup } from '@angular/forms';
import { NPPFolder, SharepointService } from 'src/app/services/sharepoint.service';
import { Observable } from 'rxjs';

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

      for (const scen of this.model.scenario) {
        // forecast model file
        let newFileData = {...fileData};
        Object.assign(newFileData, {
          // CountryId: this.model.country,
          GeographyId: this.model.geography,
          ModelScenarioId: [scen],
          ModelApprovalComments: this.model.description,
          ApprovalStatusId: this.sharepoint.getApprovalStatusId("In Progress"),
        });
        let scenarioFileName = this.model.file[0].name.replace(/[~#%&*{}:<>?+|"/\\]/g, "");

        if (this.model.scenario.length > 1) {
          // add sufix to every copy
          scenarioFileName = await this.sharepoint.addScenarioSufixToFilename(scenarioFileName, scen);
        }
        this.uploadFileToFolder(newFileData, scenarioFileName, this.sharepoint.getBaseFilesFolder() + fileFolder);
      }
    } else {
      // regular file
      Object.assign(fileData, {
        ModelApprovalComments: this.model.description
      });
      this.uploadFileToFolder(fileData, this.model.file[0].name.replace(/[~#%&*{}:<>?+|"/\\]/g, ""), this.sharepoint.getBaseFilesFolder() + fileFolder);
    }
  }

  private uploadFileToFolder(fileData: any, fileName: string, folder: string) {
    this.readFileDataAsText(this.model.file[0]).subscribe(
      data => {
        this.sharepoint.uploadFile(data, folder, fileName, fileData).then(
          r => { 
            if (Object.keys(r).length > 0) {
              this.uploading = this.dialogRef.disableClose = false; // finished

              this.dialogRef.close({
                success: true, 
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
