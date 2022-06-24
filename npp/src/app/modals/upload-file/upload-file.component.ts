import { Inject, Component, OnInit } from '@angular/core';
import { MatDialog, MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { FormControl, FormGroup } from '@angular/forms';
import { Observable } from 'rxjs';
import { ConfirmDialogComponent } from '../confirm-dialog/confirm-dialog.component';
import { take } from 'rxjs/operators';
import { NPPFolder } from '@shared/models/file-system';
import { Indication } from '@shared/models/entity';
import { FilesService } from 'src/app/services/files.service';
import { UploadFileConfig } from '@shared/forms/upload-file.config';

@Component({
  selector: 'app-upload-file',
  templateUrl: './upload-file.component.html',
  styleUrls: ['./upload-file.component.scss']
})
export class UploadFileComponent implements OnInit {
  formConfig: UploadFileConfig;
  fields: FormlyFieldConfig[] = [];
  form: FormGroup = new FormGroup({});
  model: any = { };
  folders: NPPFolder[] = [];
  uploading = false; // spinner control

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<UploadFileComponent>,
    private matDialog: MatDialog,
    private readonly files: FilesService
  ) { 
    this.formConfig = new UploadFileConfig();
  }

  ngOnInit(): void {
    this.fields = this.formConfig.fields(
      this.data.entity.ID, 
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
    
    this.uploading = this.dialogRef.disableClose = true;

    let folderToUpload = await this.files.constructUploadFolder(this.data.entity, this.data.masterStageId, this.model.category, this.model.geography);
    const uploadingModel = !!this.model.geography;

    let fileData = {};
    if (uploadingModel) {
      fileData = await this.files.prepareUploadModelData(this.model);
    } else {
      fileData = await this.files.prepareUploadFileData(this.model);
    }

    let cleanFileName = this.model.file[0].name.replace(/[~#%&*{}:<>?+|"/\\]/g, ""); // clean filename
    let sameTagsExists = null;
    if (uploadingModel) sameTagsExists = await this.files.getFileWithSameTags(folderToUpload, this.model.scenario, this.model.IndicationId);
    let fileExists = await this.files.fileExistsInFolder(cleanFileName, folderToUpload);
    if (fileExists || sameTagsExists) {
      let message = '';
      if(fileExists) {
        message = `A model with this name (${cleanFileName}) already exists in this location.`
        if(sameTagsExists) {
          message += " Also, a model with the same scenarios and indications exists. Do you want to overwrite everything?"
        } else {
          message += " Do you want to overwrite it?"
        }
      } else if(sameTagsExists) {
        message = "A model with the same scenarios and indications already exists. Do you want to overwrite it?"
      }

      const dialogRef = this.matDialog.open(ConfirmDialogComponent, {
        maxWidth: "400px",
        height: "250px",
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
            this.uploadFileToFolder(fileData, cleanFileName, folderToUpload);
          } else {
            // do nothing and close
            this.dialogRef.close({
              success: true, 
              uploaded: false
            });
          }
        });
      } else {
        this.uploadFileToFolder(fileData, cleanFileName, folderToUpload);
      }
  }

  private async uploadFileToFolder(fileData: any, fileName: string, folder: string) {
    this.readFileDataAsText(this.model.file[0]).subscribe(
      data => {
        this.files.uploadFileToFolder(data, folder, fileName, fileData).then(
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
