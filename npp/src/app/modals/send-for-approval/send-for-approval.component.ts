import { Component, Inject, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { SharepointService } from 'src/app/services/sharepoint.service';

@Component({
  selector: 'app-send-for-approval',
  templateUrl: './send-for-approval.component.html',
  styleUrls: ['./send-for-approval.component.scss']
})
export class SendForApprovalComponent implements OnInit {

  fileId: number | null = null;
  folder: string | null = null;

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
  ) { }

  ngOnInit(): void {
    this.fileId = this.data.fileId ? this.data.fileId : this.data.file?.ListItemAllFields?.ID;
    this.folder = this.data.rootFolder ? this.data.rootFolder : null;
  }

  async onSubmit() {
    if (this.fileId) {
      const  result = await this.sharepoint.setApprovalStatus(this.fileId, "Submitted", this.model.comments, this.data.rootFolder);
      this.dialogRef.close({
        success: result,
        comments: this.model.comments
      });
    }
  }

}
