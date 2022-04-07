import { Component, Inject, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { InlineNppDisambiguationService } from 'src/app/services/inline-npp-disambiguation.service';
import { SharepointService } from 'src/app/services/sharepoint.service';
import { Opportunity } from '@shared/models/entity';
import { NPPFile } from '@shared/models/file-system';
import { AppDataService } from 'src/app/services/app-data.service';
import { FilesService } from 'src/app/services/files.service';

@Component({
  selector: 'app-send-for-approval',
  templateUrl: './send-for-approval.component.html',
  styleUrls: ['./send-for-approval.component.scss']
})
export class SendForApprovalComponent implements OnInit {

  fileId: number | null = null;
  folder: string | null = null;
  file: NPPFile | null = null;
  entity: Opportunity | null = null;
  fields: FormlyFieldConfig[] = [{
    fieldGroup: [{
      key: 'comments',
      type: 'textarea',
      templateOptions: {
          label: 'Comments:',
          placeholder: 'Please enter comments',
          rows: 3
      }
    }]
  }];

  form: FormGroup = new FormGroup({});
  model: any = {}; 

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<SendForApprovalComponent>,
    private readonly sharepoint: SharepointService,
    private readonly disambiguator: InlineNppDisambiguationService,
    private readonly appData: AppDataService,
    private readonly files: FilesService
  ) { }

  ngOnInit(): void {
    this.fileId = this.data.fileId ? this.data.fileId : this.data.file?.ListItemAllFields?.ID;
    this.folder = this.data.rootFolder ? this.data.rootFolder : null;
    this.file = this.data.file;
    this.entity = this.data.entity;
  }

  async onSubmit() {
    if (this.fileId && this.file) {
      let commentsStr = '';
      if(this.model.comments) {
        commentsStr = await this.files.addFileComment(this.file, this.model.comments);
      }
      const result = await this.files.setFileApprovalStatus(this.data.rootFolder, this.file, this.entity, "Submitted", commentsStr);
      this.dialogRef.close({
        success: result,
        comments: this.model.comments
      });
    }
  }

}
