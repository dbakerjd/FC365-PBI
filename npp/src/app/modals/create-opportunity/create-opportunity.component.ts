import { Component, Inject, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { SelectInputList, SharepointService } from 'src/app/services/sharepoint.service';
import { take, takeUntil, tap } from 'rxjs/operators';
import { Subject } from 'rxjs';
import { MatDialog, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { ToastrService } from 'ngx-toastr';

@Component({
  selector: 'app-create-opportunity',
  templateUrl: './create-opportunity.component.html',
  styleUrls: ['./create-opportunity.component.scss']
})
export class CreateOpportunityComponent implements OnInit {

  private readonly _destroying$ = new Subject<void>();
  
  form = new FormGroup({});
  model: any = { };
  fields: FormlyFieldConfig[] = [];
  indications: any[] = [];
  dialogInstance: any;
  firstStepCompleted: boolean = false;
  isEdit = false;

  constructor(
    private sharepoint: SharepointService, 
    private toastr: ToastrService,
    public matDialog: MatDialog,
    @Inject(MAT_DIALOG_DATA) public data: any,
    ) { }

  async ngOnInit() {

    let therapies = await this.sharepoint.getTherapiesList();
    let oppTypes = await this.sharepoint.getOpportunityTypesList();
    let indicationsList: any[] = [];
    let defaultUsersList: SelectInputList[] = [];
    this.firstStepCompleted = false;

    if (this.data?.opportunity) {
      this.isEdit = true;
      indicationsList = await this.sharepoint.getIndicationsList(this.data.opportunity.Indication.TherapyArea);
      defaultUsersList = [{ 
        label: this.data.opportunity.OpportunityOwner.FirstName + ' ' + this.data.opportunity.OpportunityOwner.LastName,
        value: this.data.opportunity.OpportunityOwnerId
      }];
    }

    this.fields = [
      {
        fieldGroup: [{
          key: 'Opportunity.Title',
          type: 'input',
          templateOptions: {
            label: 'Opportunity Name:',
            placeholder: 'Opportunity Name',
            required: true,
          },
          defaultValue: this.data?.opportunity.Title
        }, {
          key: 'Opportunity.MoleculeName',
          type: 'input',
          templateOptions: {
            label: 'Molecule Name:',
            placeholder: 'Molecule Name',
            required: true,
          },
          defaultValue: this.data?.opportunity.MoleculeName
        }, {
          key: 'Opportunity.OpportunityOwnerId',
          type: 'ngsearchable',
          templateOptions: {
            label: 'Opportunity Owner:',
            placeholder: 'Opportunity Owner',
            required: true,
            filterLocally: false,
            query: 'siteusers',
            options: defaultUsersList
          },
          defaultValue: this.data?.opportunity.OpportunityOwnerId
        }, {
          key: 'therapy',
          type: 'select',
          templateOptions: {
            label: 'Therapy Area:',
            options: therapies,
            required: true,
          },
          defaultValue: this.data?.opportunity.Indication.TherapyArea,
        }, {
          key: 'Opportunity.IndicationId',
          type: 'select',
          templateOptions: {
            label: 'Indication Name:',
            options: indicationsList,
            required: true,
          },
          defaultValue: this.data?.opportunity.IndicationId,
          hooks: {
            onInit: (field) => {
              if (!field?.parent?.fieldGroup) return;
              const therapySelect = field.parent.fieldGroup.find(f => f.key === 'therapy');
              if (!therapySelect?.formControl) return;
              therapySelect.formControl.valueChanges.pipe(
                takeUntil(this._destroying$),
                tap(th => {
                  this.sharepoint.getIndicationsList(th).then(r => {
                    field.formControl?.setValue('');
                    if (field.templateOptions) field.templateOptions.options = r;
                  });
                }),
              ).subscribe();
            }
          }
        }, {
          key: 'Opportunity.OpportunityTypeId',
          type: 'select',
          templateOptions: {
            label: 'Opportunity Type:',
            options: oppTypes,
            required: true,
            change: (field) => {
              field.formControl?.valueChanges
              .pipe(take(1), takeUntil(this._destroying$))
              .subscribe(
                (selectedValue) => {
                  this.sharepoint.getStageType(selectedValue).then(r => {
                    if (r) this.model.stageType = r;
                  });
                }
            );
            }
          },
          defaultValue: this.data?.opportunity.OpportunityTypeId
        }, {
          key: 'Opportunity.ProjectStartDate',
          type: 'datepicker',
          templateOptions: {
            label: 'Project Start Date:',
            required: true,
          },
          defaultValue: this.data?.opportunity.ProjectStartDate ? new Date(this.data?.opportunity.ProjectStartDate) : null
        }, {
          key: 'Opportunity.ProjectEndDate',
          type: 'datepicker',
          templateOptions: {
            label: 'Project End Date:',
            required: true,
          },
          defaultValue: this.data?.opportunity.ProjectEndDate ? new Date(this.data?.opportunity.ProjectEndDate) : null
        }],
        hideExpression: this.firstStepCompleted
      },
      {
        template: '<div class="form-header">Complete First Stage Info</div><hr />',
        hideExpression: !this.firstStepCompleted,
        expressionProperties: {
          'template': function($viewValue, $modelValue, scope) {
            return `<div class="form-header">The Opportunity Stage Type is ${scope?.model.stageType}</div><hr />`
          },
        },
      },
      {
        fieldGroup: [{
          key: 'stageType',
          type: 'input',
          hideExpression: true,
        }, {
          key: 'Stage.Title',
          type: 'input',
          templateOptions: {
            label: 'Stage Name:',
            placeholder: 'First Stage Name'
          },
          expressionProperties: {
            'templateOptions.label': function($viewValue, $modelValue, scope) {
              return `${scope?.model.stageType} Name`;
            },
            'templateOptions.placeholder': function($viewValue, $modelValue, scope) {
              return `First ${scope?.model.stageType} Name`;
            },
          },
        },{
          key: 'Stage.StageUsersId',
          type: 'ngsearchable',
          templateOptions: {
              label: 'Stage Users:',
              placeholder: 'Stage Users',
              filterLocally: false,
              query: 'siteusers',
              multiple: true
          }
        },{
          key: 'Stage.StageReview',
          type: 'datepicker',
          templateOptions: {
              label: 'Stage Review:'
          }
        },],
        hideExpression: !this.firstStepCompleted
      }
    ];
  }

  onNext() {
    this.firstStepCompleted = true;
    this.fields[0].hideExpression = true;
    this.fields[1].hideExpression = this.fields[2].hideExpression = false;
  }

  async onSubmit() {
    if (this.isEdit) {
      const success = await this.sharepoint.updateOpportunity(this.data.opportunity.ID, this.model.Opportunity);
      if (success) this.toastr.success("A opportunity was updated successfully", this.model.Title);
      else this.toastr.error("The opportunity couldn't be created", "Try again");
    } else {
      const success = await this.sharepoint.createOpportunity(this.model.Opportunity, this.model.Stage);
      if (success) this.toastr.success("A new opportunity was created successfully", this.model.Title);
      else this.toastr.error("The opportunity couldn't be created", "Try again");
    }
  }
 
  ngOnDestroy(): void {
    this._destroying$.next();
    this._destroying$.complete();
  }
}


