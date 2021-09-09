import { Component, Inject, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { SelectInputList, SharepointService, Stage } from 'src/app/services/sharepoint.service';

@Component({
  selector: 'app-stage-settings',
  templateUrl: './stage-settings.component.html',
  styleUrls: ['./stage-settings.component.scss']
})
export class StageSettingsComponent implements OnInit {
  
  form = new FormGroup({});
  model: any = { };
  fields: FormlyFieldConfig[] = [];
  isEdit: boolean = false;

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<StageSettingsComponent>,
    private sharepoint: SharepointService, 
  ) { }

  async ngOnInit() {
    let defaultUsersList: SelectInputList[] = [];
    if (this.data?.stage) {
      defaultUsersList = await this.sharepoint.getUsersList(this.data?.stage.StageUsersId);
    }

    this.fields = [{
      fieldGroup: [{
        key: 'stageType',
        type: 'input',
        hideExpression: true,
      },{
        key: 'opportunityId',
        type: 'input',
        hideExpression: true,
      },{
        key: 'nextMasterStageId',
        type: 'input',
        hideExpression: true,
      },{
        key: 'ID',
        type: 'input',
        hideExpression: true,
      },{
        key: 'Title',
        type: 'input',
        templateOptions: {
          label: 'Stage Name:',
          placeholder: 'Set the next stage name',
          required: true
        },
        expressionProperties: {
          'templateOptions.label': function($viewValue, $modelValue, scope) {
            if (scope?.model.stageType) return `${scope?.model.stageType} Name`;
            else return '';
          },
        },
        hideExpression: 'model.ID'
      },{
        key: 'StageUsersId',
        type: 'ngsearchable',
        templateOptions: {
            label: 'Stage Users',
            placeholder: 'Stage Users',
            filterLocally: false,
            query: 'siteusers',
            multiple: true,
            options: defaultUsersList,
            required: true
        }
      },{
        key: 'StageReview',
        type: 'datepicker',
        templateOptions: {
            label: 'Stage Review',
            required: true
        }
      }]
    }];
    if (this.data?.stage) { // edit stage
      this.isEdit = true;
      this.model.ID = this.data.stage.ID;
      this.model.Title = this.data.stage.Title;
      this.model.StageUsersId = this.data.stage.StageUsersId;
      this.model.StageReview = new Date(this.data.stage.StageReview);
      this.model.stageType = this.data.stage.StageType;
    }
    if (this.data?.next) {
      this.model.stageType = this.data.next.Title;
      this.model.nextMasterStageId = this.data.next.ID;
      this.model.opportunityId = this.data.opportunityId;
    }
  }

  async onSubmit() {
    let success;
    if (this.model.ID) {
      success = await this.sharepoint.updateStage(this.model.ID, {
        StageReview: this.model.StageReview,
        StageUsersId: this.model.StageUsersId
      });
      this.dialogRef.close({
        success, 
        data: this.model
      });
      
    } else {
      const newStage = await this.sharepoint.createStage({
        Title: this.model.Title,
        StageReview: this.model.StageReview,
        StageUsersId: this.model.StageUsersId,
        OpportunityNameId: this.model.opportunityId,
        StageNameId: this.model.nextMasterStageId
      });
      this.dialogRef.close({
        success: newStage ? true : false, 
        data: newStage
      });
    }
  }
}
