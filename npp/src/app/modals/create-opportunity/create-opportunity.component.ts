import { Component, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { OpportunityInput, SharepointService } from 'src/app/services/sharepoint.service';
import { take, takeUntil, tap } from 'rxjs/operators';
import { Subject } from 'rxjs';
import { MatDialog } from '@angular/material/dialog';
import { StageSettingsComponent } from '../stage-settings/stage-settings.component';

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


  constructor(private sharepoint: SharepointService, public matDialog: MatDialog) { }

  async ngOnInit() {

    let therapies = await this.sharepoint.getTherapiesList();
    let oppTypes = await this.sharepoint.getOpportunityTypesList();
    this.firstStepCompleted = false;

    this.fields = [
      {
        fieldGroup: [{
          key: 'Opportunity.Title',
          type: 'input',
          templateOptions: {
            label: 'Opportunity Name:',
            placeholder: 'Opportunity Name',
            required: true,
          }
        }, {
          key: 'Opportunity.MoleculeName',
          type: 'input',
          templateOptions: {
            label: 'Molecule Name:',
            placeholder: 'Molecule Name',
            required: true,
          }
        }, {
          key: 'Opportunity.OpportunityOwnerId',
          type: 'input',
          templateOptions: {
            label: 'Opportunity Owner:',
            placeholder: 'Opportunity Owner',
            required: true,
          }
        }, {
          key: 'therapy',
          type: 'select',
          templateOptions: {
            label: 'Therapy Area:',
            options: therapies,
            required: true,
          },
        }, {
          key: 'Opportunity.IndicationId',
          type: 'select',
          templateOptions: {
            label: 'Indication Name:',
            options: [],
            required: true,
          },
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
        }, {
          key: 'Opportunity.ProjectStartDate',
          type: 'datepicker',
          templateOptions: {
            label: 'Project Start Date:',
            required: true,
          }
        }, {
          key: 'Opportunity.ProjectEndDate',
          type: 'datepicker',
          templateOptions: {
            label: 'Project End Date:',
            required: true,
          }
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
        }/*,{
          key: 'Stage.StageUsersId',
          type: 'input',
          templateOptions: {
              label: 'Stage Users:',
              placeholder: 'Stage Users'
          }
        }*/,{
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

  async onNext() {
    this.firstStepCompleted = true;
    this.fields[0].hideExpression = true;
    this.fields[1].hideExpression = this.fields[2].hideExpression = false;
  }

  onSubmit() {
    console.log(this.model);
    // delete this.model.therapy;
    // let newOpportunity: OpportunityInput = { ...this.model.Opportunity };
    // this.dialogInstance = this.matDialog.open(StageSettingsComponent, {
    //   height: '700px',
    //   width: '405px'
    // });
    this.sharepoint.createOpportunity(this.model.Opportunity, this.model.Stage);
  }

  /*
  async setIndications(value: any, a?: any) {
    console.log(value);
    console.log(value.model);
    let indications = await this.sharepoint.getIndicationsList(value.model.threapy);
  }
  */

  /*
  getField(key: string, fields: FormlyFieldConfig[]): FormlyFieldConfig | null {
    for (let i = 0, len = fields.length; i < len; i++) {
      const f = fields[i];
      if (f.key === key) {
        return f;
      }
      
      if (f.fieldGroup && !f.key) {
        const cf = this.getField(key, f.fieldGroup);
        if (cf) {
          return cf;
        }
      }
    }
    return null;
  }
  */
  
  ngOnDestroy(): void {
    this._destroying$.next();
    this._destroying$.complete();
  }
}


