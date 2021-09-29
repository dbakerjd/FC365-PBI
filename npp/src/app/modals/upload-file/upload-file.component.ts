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

  onSubmit() {
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
      // forecast model file
      Object.assign(fileData, {
        CountryId: this.model.country,
        GeographyId: this.model.geography,
        ModelScenarioId: this.model.scenario,
        ModelApprovalComments: this.model.description,
        ApprovalStatusId: this.sharepoint.getApprovalStatusId("In Progress"),
      });
      fileFolder += '/' + this.model.geography;
    } else {
      // regular file
      Object.assign(fileData, {
        ModelApprovalComments: this.model.description
      });
    }

    // upload file to correspondent folder
    const folder = this.sharepoint.getBaseFilesFolder() + fileFolder;
    this.readFileDataAsText(this.model.file[0]).subscribe(
      data => {
        this.sharepoint.uploadFile(data, folder, this.model.file[0].name, fileData).then(
          r => { 
            if (Object.keys(r).length > 0) {
              this.uploading = this.dialogRef.disableClose = false; // finished

              this.dialogRef.close({
                success: true, 
                name: this.model.file[0].name
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
