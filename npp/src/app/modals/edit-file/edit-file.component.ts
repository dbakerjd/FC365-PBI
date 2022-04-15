import { Component, Inject, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { NPPFile } from '@shared/models/file-system';
import { AppDataService } from '@services/app/app-data.service';

@Component({
  selector: 'app-edit-file',
  templateUrl: './edit-file.component.html',
  styleUrls: ['./edit-file.component.scss']
})
export class EditFileComponent implements OnInit {

  fileInfo: NPPFile | null = null;
  extension = '';

  fields: FormlyFieldConfig[] = [{
    fieldGroup: [{
      key: 'filename',
      type: 'input',
      templateOptions: {
        label: 'File name:',
        required: true
      },
    }]
  }];

  form: FormGroup = new FormGroup({});
  model: any = {}; 

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<EditFileComponent>,
    private readonly appData: AppDataService
  ) { }

  ngOnInit(): void {
    this.fileInfo = this.data.fileInfo;
    if (this.fileInfo) {
      const arrFile = this.fileInfo.Name.split(".");
      this.extension = arrFile[arrFile.length - 1];
      const filename = this.fileInfo.Name.substr(0, this.fileInfo.Name.length - (this.extension.length + 1))

      this.model.filename = filename;
    }
  }

  async onSubmit() {
    if (this.fileInfo) {
      const newFilename = this.model.filename.replace(/[~#%&*{}:<>?+|"/\\]/g, "");
      let result = false;
      if (newFilename.length > 0) {
        result = await this.appData.renameFile(this.fileInfo.ServerRelativeUrl, newFilename);
      }
      this.dialogRef.close({
        success: result,
        filename: newFilename + '.' + this.extension
      });
    }
  }

}
