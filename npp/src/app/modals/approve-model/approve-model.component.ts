import { Component, Inject, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { SharepointService } from 'src/app/services/sharepoint.service';
import { Opportunity } from '@shared/models/entity';
import { NPPFile } from '@shared/models/file-system';

@Component({
  selector: 'app-approve-model',
  templateUrl: './approve-model.component.html',
  styleUrls: ['./approve-model.component.scss']
})
export class ApproveModelComponent implements OnInit {

  file: NPPFile | null = null;
  brand: Opportunity | null = null; 
  rootFolder: string = "";
  approving = false;

  fields: FormlyFieldConfig[] = [{
    fieldGroup: [{
      key: 'comments',
      type: 'textarea',
      templateOptions: {
          label: 'Comments:',
          placeholder: 'Please enter comment.',
          rows: 3
      }
    }]
  }];

  form: FormGroup = new FormGroup({});
  model: any = {}; 

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<ApproveModelComponent>,
    private readonly sharepoint: SharepointService,
  ) { }

  ngOnInit(): void {
    this.file = this.data.file;
    this.brand = this.data.brand;
    this.rootFolder = this.data.rootFolder;
  }

  async onSubmit() {
    try {
      if (this.file) {
        let commentsStr = '';
        this.approving = true;
        if(this.model.comments) {
          commentsStr = await this.sharepoint.addComment(this.file, this.model.comments);
        }
        const result = await this.sharepoint.setBrandApprovalStatus(this.rootFolder, this.file, this.brand, "Approved", commentsStr);
        this.approving = false;
        this.dialogRef.close({
          success: result,
          comments: commentsStr
        });
      }
    } catch(e) {
      this.approving = false;
    }
    
  }

}
