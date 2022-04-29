import { Component, Inject, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { Opportunity } from '@shared/models/entity';
import { NPPFile } from '@shared/models/file-system';
import { FilesService } from 'src/app/services/files.service';

@Component({
  selector: 'app-reject-model',
  templateUrl: './reject-model.component.html',
  styleUrls: ['./reject-model.component.scss']
})
export class RejectModelComponent implements OnInit {

  fileId: number | null = null;
  file: NPPFile | null = null;
  rootFolder: string = '';
  entity: Opportunity | null = null;

  fields: FormlyFieldConfig[] = [{
    fieldGroup: [{
      key: 'comments',
      type: 'textarea',
      templateOptions: {
          label: 'Comments:',
          placeholder: 'Please enter comments for rejecting',
          rows: 3
      }
    }]
  }];

  form: FormGroup = new FormGroup({});
  model: any = {}; 

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<RejectModelComponent>,
    private readonly files: FilesService
  ) { }

  ngOnInit(): void {
    this.fileId = this.data.fileId ? this.data.fileId : this.data.file?.ListItemAllFields?.ID;
    this.file = this.data.file;
    this.rootFolder = this.data.rootFolder;
    this.entity = this.data.entity;
  }

  async onSubmit() {
    if (this.file) {
      let commentsStr = '';
      if(this.model.comments) {
        commentsStr = await this.files.addFileComment(this.file, this.model.comments);
      }
      const  result = await this.files.setFileApprovalStatus(this.rootFolder, this.file, this.entity, "In Progress", commentsStr);
      this.dialogRef.close({
        success: result,
        comments: this.model.comments
      });
    }
  }

}
