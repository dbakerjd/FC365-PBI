import { Component, Inject, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { Indication } from '@shared/models/entity';
import { NPPFile } from '@shared/models/file-system';
import { AppDataService } from '@services/app/app-data.service';

@Component({
  selector: 'app-entity-edit-file',
  templateUrl: './entity-edit-file.component.html',
  styleUrls: ['./entity-edit-file.component.scss']
})
export class EntityEditFileComponent implements OnInit {

  fileInfo: NPPFile | null = null;
  extension = '';
  oldName: string = '';

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
    public dialogRef: MatDialogRef<EntityEditFileComponent>,
    private readonly appData: AppDataService
  ) { }

  ngOnInit(): void {
    this.fileInfo = this.data.fileInfo;
    if (this.fileInfo) {
      const arrFile = this.fileInfo.Name.split(".");
      this.extension = arrFile[arrFile.length - 1];
      const filename = this.fileInfo.Name.substr(0, this.fileInfo.Name.length - (this.extension.length + 1))
      this.oldName = filename;
      this.model.filename = filename;
    }
    if(this.fields && this.fields.length && this.fields[0].fieldGroup && this.fields[0].fieldGroup.length && this.fileInfo && this.fileInfo.ListItemAllFields) {
      this.fields[0].fieldGroup.push({
        key: 'IndicationId',
        type: 'ngsearchable',
        templateOptions: {
          label: 'Indication Name:',
          options: this.data.entity.Indication.map((el: Indication) => { return { label: el.Title, value: el.ID}}),
          multiple: true,
          required: true
        },
        defaultValue: this.fileInfo.ListItemAllFields.IndicationId
      });
    }
    
  }
  compareArr(arr1: number[], arr2: number[]){
    arr1.sort()
    arr2.sort()
    return arr1 + "" == arr2 + ""
  }

  async onSubmit() {
    if (this.fileInfo) {
      const newFilename = this.model.filename.replace(/[~#%&*{}:<>?+|"/\\]/g, "");
      let result = false;
      let result2 = false;
      let needsRename = (newFilename != this.oldName);
      let needsIndicationsUpdate = this.fileInfo?.ListItemAllFields?.IndicationId ? !this.compareArr(this.fileInfo.ListItemAllFields.IndicationId, this.model.IndicationId) : this.model.IndicationId;

      if (newFilename.length > 0 && needsRename) {
        result = await this.appData.renameFile(this.fileInfo.ServerRelativeUrl, newFilename);
      }

      if(needsIndicationsUpdate && this.fileInfo.ListItemAllFields) {
        let arrFolder = this.fileInfo.ServerRelativeUrl.split("/");
        let rootFolder = arrFolder[3];  
      
        result2 = await this.appData.updateFilePropertiesById(this.fileInfo.ListItemAllFields.ID, rootFolder, {
          IndicationId: this.model.IndicationId
        });
      }

      if(this)
      this.dialogRef.close({
        success: {
          needsRename,
          needsIndicationsUpdate,
          renameWorked: result,
          indicationsUpdateWorked: result2
        },
        filename: newFilename + '.' + this.extension,
        IndicationId: this.model.IndicationId
      });
    }
  }
}
