import { Inject, Component, OnInit } from '@angular/core';
import { UploadFileConfig } from 'src/app/shared/forms/upload-file.config';
import { MatDialog, MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { FormControl, FormGroup } from '@angular/forms';
import { FOLDER_DOCUMENTS, FOLDER_WIP, FORECAST_MODELS_FOLDER_NAME, Indication, NPPFolder, SharepointService } from 'src/app/services/sharepoint.service';
import { Observable } from 'rxjs';
import { ConfirmDialogComponent } from '../confirm-dialog/confirm-dialog.component';
import { take } from 'rxjs/operators';
import { ExternalUploadFileConfig } from 'src/app/shared/forms/external-upload-file.config';
import { InlineNppDisambiguationService } from 'src/app/services/inline-npp-disambiguation.service';

@Component({
  selector: 'app-external-upload-file',
  templateUrl: './external-upload-file.component.html',
  styleUrls: ['./external-upload-file.component.scss']
})
export class ExternalUploadFileComponent implements OnInit {
  formConfig: ExternalUploadFileConfig = new ExternalUploadFileConfig();
  fields: FormlyFieldConfig[] = [];
  form: FormGroup = new FormGroup({});
  model: any = { };
  folders: NPPFolder[] = [];
  uploading = false; // spinner control

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<ExternalUploadFileComponent>,
    private readonly sharepoint: SharepointService,
    private matDialog: MatDialog,
    private readonly disambiguator: InlineNppDisambiguationService
  ) { 
    
  }

  ngOnInit(): void {
    this.formConfig = new ExternalUploadFileConfig();
    this.fields = this.formConfig.fields(
      this.data.entity.ID, 
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
    
    this.uploading = this.dialogRef.disableClose = true;

    let fileData: any = {
      EntityNameId: this.model.entityId
    };

    let fileFolder = FOLDER_WIP+'/'+this.data.entity.BusinessUnitId+'/'+this.data.entity.ID+'/0/0';
    let containsModels = true;
    if(this.model.category !== 0) {
      fileFolder = FOLDER_DOCUMENTS+'/'+this.data.entity.BusinessUnitId+'/'+this.data.entity.ID+'/0/'+this.model.category;
      containsModels = false;
      fileData = {
        Comments: this.model.description
      }
    }

    if (containsModels) {
      // forecast model file

      const oppGeographies = await this.disambiguator.getEntityGeographies(this.data.entity.ID);
      const geography = oppGeographies.find(el => el.Id == this.model.geography);
      const user = await this.sharepoint.getCurrentUserInfo();

      Object.assign(fileData, {
        CountryId: this.model.country,
        EntityGeographyId: geography.Id ? geography.Id : null,
        ModelScenarioId: this.model.scenario,
        Comments: this.model.description ? '[{"text":"'+this.model.description.replace(/'/g, "{COMMA}")+'","email":"'+user.Email+'","createdAt":"'+new Date().toISOString()+'"}]' : '[]',
        ApprovalStatusId: await this.sharepoint.getApprovalStatusId("In Progress"),
        IndicationId: this.model.IndicationId
      });

      fileFolder += '/' + geography.Id + '/0';
    } else {
      fileFolder += '/0/0';
    }

    let scenarioFileName = this.model.file[0].name.replace(/[~#%&*{}:<>?+|"/\\]/g, "");
    let scenarioExists = await this.disambiguator.getFileByScenarios(fileFolder, this.model.scenario);
    let fileExists = await this.sharepoint.existsFile(scenarioFileName, fileFolder);
    if (fileExists || scenarioExists) {
      let message = '';
      if(fileExists) {
        message = `A model with this name (${scenarioFileName}) already exists in this location.`
        if(scenarioExists) {
          message += " Also, a model with the same scenario exists. Do you want to overwrite everything?"
        } else {
          message += " Do you want to overwrite it?"
        }
      } else if(scenarioExists) {
        message = "A model with the same scenario already exists. Do you want to overwrite it?"
      }

      const dialogRef = this.matDialog.open(ConfirmDialogComponent, {
        maxWidth: "400px",
        height: "200px",
        data: {
          message,
          confirmButtonText: 'Yes, overwrite',
          cancelButtonText: 'No, keep the original'
        }
      });
    
      dialogRef.afterClosed()
        .pipe(take(1))
        .subscribe(async uploadConfirmed => {
          if (uploadConfirmed) {
            this.uploadFileToFolder(fileData, scenarioFileName, fileFolder);
          } else {
            // do nothing and close
            this.dialogRef.close({
              success: true, 
              uploaded: false
            });
          }
        });
      } else {
        this.uploadFileToFolder(fileData, scenarioFileName, fileFolder);
      }
  }

  private async uploadFileToFolder(fileData: any, fileName: string, folder: string) {
    this.readFileDataAsText(this.model.file[0]).subscribe(
      data => {
        this.disambiguator.uploadFile(data, folder, fileName, fileData).then(
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
