import { Inject, Component, OnInit } from '@angular/core';
import { UploadFileConfig } from 'src/app/shared/forms/upload-file.config';
import { MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { FormGroup } from '@angular/forms';
import { SharepointService } from 'src/app/services/sharepoint.service';

@Component({
  selector: 'app-upload-file',
  templateUrl: './upload-file.component.html',
  styleUrls: ['./upload-file.component.scss']
})
export class UploadFileComponent implements OnInit {
  formConfig: UploadFileConfig = new UploadFileConfig();
  fields: FormlyFieldConfig[] = [];
  form: FormGroup = new FormGroup({});
  model: any = { };

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    private readonly sharepoint: SharepointService
  ) { 
    
  }

  ngOnInit(): void {
    this.formConfig = new UploadFileConfig();
    this.fields = this.formConfig.fields(
      this.data.opportunityId, 
      this.data.masterStageId, 
      this.data.folderList,
      this.data.countries,
      this.data.scenarios);
    this.form = new FormGroup({});
  }

  onSubmit() {
    console.log('model', this.model);

    let fileData = {
      StageNameId: this.model.StageNameId,
      OpportunityNameId: this.model.OpportunityNameId,
    };
    console.log(fileData);

    if (this.model.category == 6) { // TODO
      // forecast model file

    } else {
      // regular file
      Object.assign(fileData, {
        ModelApprovalComments: this.model.description
      });

      console.log(fileData);
    }

    this.sharepoint.uploadFile(this.model.file[0], 'Current Opportunity Library',fileData).then(
      r => { console.log('upload response', r); }
    )
  }
}
