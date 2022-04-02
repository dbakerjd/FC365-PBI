import { Component, Inject, OnInit } from '@angular/core';
import { FormControl, FormGroup } from '@angular/forms';
import { MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { AppDataService } from 'src/app/services/app-data.service';
import { SelectInputList, SharepointService } from 'src/app/services/sharepoint.service';

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
  canSetUsers: boolean = false;

  // spinner control
  loading = true;
  updating = false;

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<StageSettingsComponent>,
    private sharepoint: SharepointService, 
    private readonly appData: AppDataService
  ) { }

  async ngOnInit() {
    let defaultUsersList: SelectInputList[] = [];
    if (this.data?.stage) {
      defaultUsersList = await this.appData.getUsersList(this.data?.stage.StageUsersId);
    }
    this.canSetUsers = this.data?.canSetUsers ? this.data.canSetUsers : false;

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
        },
        validation: {
          messages: {
            required: (error) => 'You must enter one or more users',
          },
        },
        hideExpression: !this.canSetUsers
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
    this.loading = false;
  }

  async onSubmit() {
    if (this.form.invalid) {
      this.validateAllFormFields(this.form);
      return;
    }
    let success;
    if (this.model.ID) { // update
      this.updating = this.dialogRef.disableClose = true;
      success = await this.appData.updateStage(this.model.ID, {
        StageReview: this.model.StageReview,
        StageUsersId: this.model.StageUsersId ? this.model.StageUsersId : this.data.stage.StageUsersId
      });
      this.updating = this.dialogRef.disableClose = false;

      this.dialogRef.close({
        success,
        data: this.model
      });

    } else {
      const newStage = await this.appData.createStage({
        StageReview: this.model.StageReview,
        StageUsersId: this.model.StageUsersId,
        EntityNameId: this.model.opportunityId,
        StageNameId: this.model.nextMasterStageId
      });
      this.dialogRef.close({
        success: newStage ? true : false, 
        data: newStage
      });
    }
  }

  validateAllFormFields(formGroup: FormGroup) {
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
