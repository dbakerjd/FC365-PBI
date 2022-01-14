import { Component, Inject, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { InlineNppDisambiguationService } from 'src/app/services/inline-npp-disambiguation.service';
import { Brand, NPPFile, Opportunity, SharepointService } from 'src/app/services/sharepoint.service';

@Component({
  selector: 'app-external-approve-model',
  templateUrl: './external-approve-model.component.html',
  styleUrls: ['./external-approve-model.component.scss']
})
export class ExternalApproveModelComponent implements OnInit {

  file: NPPFile | null = null;
  entity: Brand | Opportunity | null = null; 
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
    public dialogRef: MatDialogRef<ExternalApproveModelComponent>,
    private readonly disambiguator: InlineNppDisambiguationService,
    private readonly sharepoint: SharepointService,
  ) { }

  ngOnInit(): void {
    this.file = this.data.file;
    this.entity = this.data.entity;
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
        const result = await this.disambiguator.setEntityApprovalStatus(this.rootFolder, this.file, this.entity, "Approved", commentsStr);
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
