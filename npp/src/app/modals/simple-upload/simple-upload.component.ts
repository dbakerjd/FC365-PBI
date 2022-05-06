import { Component, Inject, OnInit } from '@angular/core';
import { FormControl, FormGroup } from '@angular/forms';
import { MatDialog, MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { FilesService } from '@services/files.service';
import { GLOBAL_DOCUMENTS_FOLDER } from '@shared/sharepoint/folders';
import { Observable } from 'rxjs';
import { take } from 'rxjs/operators';
import { ConfirmDialogComponent } from '../confirm-dialog/confirm-dialog.component';

@Component({
  selector: 'app-simple-upload',
  templateUrl: './simple-upload.component.html',
  styleUrls: ['./simple-upload.component.scss']
})
export class SimpleUploadComponent implements OnInit {


  fields: FormlyFieldConfig[] = [];
  form: FormGroup = new FormGroup({});
  model: any = {};
  uploading = false; // spinner control

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<SimpleUploadComponent>,
    private matDialog: MatDialog,
    private readonly files: FilesService
  ) { }

  ngOnInit(): void {
    this.fields = [
      {
        fieldGroup: [
          {
            key: 'file',
            type: 'file-input',
            templateOptions: {
                label: 'File',
                placeholder: 'File',
                required: true
            },
          },
          {
            key: 'description',
            type: 'textarea',
            templateOptions: {
              label: 'Description:',
              placeholder: 'Description',
              rows: 3
            }
          },
        ]
      }
    ];
  }

  async onSubmit() {
    if (this.form.invalid) {
      this.validateAllFormFields(this.form);
      return;
    }

    this.uploading = this.dialogRef.disableClose = true;
    let folderToUpload = GLOBAL_DOCUMENTS_FOLDER + '/' + this.data.folder.Name;
    let cleanFileName = this.model.file[0].name.replace(/[~#%&*{}:<>?+|"/\\]/g, ""); // clean filename
    let fileExists = await this.files.fileExistsInFolder(cleanFileName, folderToUpload);

    let fileData = {
      Comments: this.model.description
    };

    if (fileExists) {
      this.askOverwriteAndUpload(fileData, cleanFileName, folderToUpload)
    } else {
      this.uploadFileToFolder(fileData, cleanFileName, folderToUpload);
    }
  }

  private askOverwriteAndUpload(fileData: any, cleanFileName: string, folderToUpload: string) {
    const message = `A file with this name (${cleanFileName}) already exists in this location. Do you want to overwrite it?`;
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
  }

  private async uploadFileToFolder(fileData: any, fileName: string, folder: string) {
    this.readFileDataAsText(this.model.file[0]).subscribe(
      (data: any) => {
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
