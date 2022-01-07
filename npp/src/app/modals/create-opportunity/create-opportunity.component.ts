import { Component, Inject, OnInit } from '@angular/core';
import { FormControl, FormGroup } from '@angular/forms';
import { FormlyFieldConfig, FormlyFormOptions } from '@ngx-formly/core';
import { Opportunity, EntityGeography, SelectInputList, SharepointService, Stage, OpportunityType } from 'src/app/services/sharepoint.service';
import { take, takeUntil, tap } from 'rxjs/operators';
import { Subject } from 'rxjs';
import { MatDialog, MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { THIS_EXPR } from '@angular/compiler/src/output/output_ast';

@Component({
  selector: 'app-create-opportunity',
  templateUrl: './create-opportunity.component.html',
  styleUrls: ['./create-opportunity.component.scss']
})
export class CreateOpportunityComponent implements OnInit {

  private readonly _destroying$ = new Subject<void>();
  
  form = new FormGroup({});
  model: any = { };
  options: FormlyFormOptions = {
    formState: {
      hideStageNumbers: true,
    },
  };
  fields: FormlyFieldConfig[] = [];
  indications: any[] = [];
  dialogInstance: any;
  firstStepCompleted: boolean = false;
  isEdit = false;
  opportunity: Opportunity | null = null;
  stage: Stage | null = null;
  loading = true;
  updating = false;
  geographies: EntityGeography[] = [];
  oppTypes: any[] = [];

  constructor(
    private sharepoint: SharepointService, 
    public matDialog: MatDialog,
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<CreateOpportunityComponent>
    ) { 
      dialogRef.disableClose = true;
    }

  async ngOnInit() {
    this.opportunity = this.data?.opportunity;
    this.isEdit = this.data?.opportunity && !this.data?.createFrom;

    const therapies = await this.sharepoint.getTherapiesList();
    this.oppTypes = await this.sharepoint.getOpportunityTypesList();
    const geo = (await this.sharepoint.getGeographiesList()).map(el => { return { label: el.label, value: 'G-' + el.value } });
    const countries = (await this.sharepoint.getCountriesList()).map(el => { return { label: el.label, value: 'C-' + el.value } });;
    const locationsList = geo.concat(countries);
    let indicationsList: SelectInputList[] = [];
    let businessUnits = await this.sharepoint.getBusinessUnitsList();
    let stageNumbersList: SelectInputList[] = [];
    let defaultUsersList: SelectInputList[] = await this.sharepoint.getSiteOwnersList();
    let defaultStageUsersList: SelectInputList[] = [];
    this.firstStepCompleted = false;
    const trialPhases = await this.sharepoint.getClinicalTrialPhases();
    
    if (this.opportunity) {
      this.geographies = await this.sharepoint.getOpportunityGeographies(this.opportunity?.ID);
      this.model.geographies = this.geographies.map(el => el.CountryId ? 'C-'+el.CountryId : 'G-' + el.GeographyId);
    
      if (this.data?.forceType) { // force Phase opportunity (complete opportunity option)
        this.oppTypes = await this.sharepoint.getOpportunityTypesList('Phase');
        this.opportunity.OpportunityTypeId = -1;
        if (this.oppTypes.length > 0) {
          this.opportunity.OpportunityTypeId = this.oppTypes[0].value;
          this.model.stageType = 'Phase';
          stageNumbersList = await this.sharepoint.getMasterStageNumbers('Phase');
        }
      }

      // default indications for the therapy selected
      if (this.opportunity && this.opportunity.Indication && this.opportunity.Indication.length) {
        indicationsList = await this.sharepoint.getIndicationsList(this.opportunity.Indication[0].TherapyArea);
      }
      // if we are cloning opportunity, get first stage info
      if (this.data?.createFrom && !this.data?.forceType) {
        this.stage = await this.sharepoint.getFirstStage(this.opportunity);
        if (this.stage) {
          defaultStageUsersList = await this.sharepoint.getUsersList(this.stage.StageUsersId);
        }
      }

      /** ALERT: Needed when we retrieve all users. For now, only owners (admin set permissions limitation)   */
      /*
      defaultUsersList = [{ 
        label: this.opportunity.EntityOwner.FirstName + ' ' + this.opportunity.EntityOwner.LastName,
        value: this.opportunity.EntityOwnerId
      }];
      */
    }

    this.loading = false;

    this.fields = [
      {
        validators: {
          validation: [
            { name: 'afterDate', options: { errorPath: 'Opportunity.ProjectEndDate' } },
          ],
        },
        fieldGroup: [{
          key: 'Opportunity.Title',
          type: 'input',
          templateOptions: {
            label: 'Opportunity Name:',
            placeholder: 'Opportunity Name',
            required: true,
          },
          defaultValue: this.opportunity?.Title
        }, {
          key: 'Opportunity.MoleculeName',
          type: 'input',
          templateOptions: {
            label: 'Molecule Name:',
            placeholder: 'Molecule Name',
            required: true,
          },
          defaultValue: this.opportunity?.MoleculeName
        }, {
          key: 'Opportunity.EntityOwnerId',
          type: 'ngsearchable',
          templateOptions: {
            label: 'Opportunity Owner:',
            placeholder: 'Opportunity Owner',
            required: true,
            options: defaultUsersList
            /** ALERT: Needed when we retrieve all users. For now, only owners (admin set permissions limitation)   */
            /*
            filterLocally: false,
            query: 'siteusers'
            */
          },
          defaultValue: this.opportunity?.EntityOwnerId
        }, {
          key: 'therapy',
          type: 'select',
          templateOptions: {
            label: 'Therapy Area:',
            options: therapies,
            required: true,
          },
          defaultValue: this.opportunity && this.opportunity.Indication && this.opportunity.Indication.length ? this.opportunity.Indication[0].TherapyArea : null
        },{
          key: 'Opportunity.IndicationId',
          type: 'ngsearchable',
          templateOptions: {
            label: 'Indication Name:',
            options: indicationsList,
            multiple: true,
            required: true,
          },
          defaultValue: this.opportunity?.IndicationId,
          hooks: {
            onInit: (field) => {
              if (!field?.parent?.fieldGroup) return;
              const therapySelect = field.parent.fieldGroup.find(f => f.key === 'therapy');
              if (!therapySelect?.formControl) return;
              therapySelect.formControl.valueChanges.pipe(
                takeUntil(this._destroying$),
                tap(th => {
                  this.sharepoint.getIndicationsList(th).then(r => {
                    if (r.length > 0) field.formControl?.setValue(r[0].value);
                    if (field.templateOptions) field.templateOptions.options = r;
                  });
                }),
              ).subscribe();
            }
          }
        }, {
          key: 'Opportunity.BusinessUnitId',
          type: 'select',
          templateOptions: {
            label: 'Business Unit:',
            options: businessUnits,
            required: true,
          },
          defaultValue: this.opportunity?.BusinessUnitId
        }, {
          key: 'Opportunity.ClinicalTrialPhaseId',
          type: 'select',
          templateOptions: {
            label: 'Clinical Trial Phase:',
            options: trialPhases,
            required: true,
          },
          defaultValue: this.opportunity?.ClinicalTrialPhaseId
        }, {
          key: 'Opportunity.OpportunityTypeId',
          type: 'select',
          templateOptions: {
            label: 'Opportunity Type:',
            options: this.oppTypes,
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
            },
          },
          defaultValue: this.opportunity?.OpportunityTypeId !== -1 ? this.opportunity?.OpportunityTypeId : null,
          hideExpression: this.isEdit
        }, {
          key: 'Opportunity.ProjectStartDate',
          type: 'datepicker',
          className: 'date-input',
          templateOptions: {
            label: 'Project Start Date:',
            required: true,
          },
          defaultValue: this.opportunity?.ProjectStartDate ? new Date(this.opportunity?.ProjectStartDate) : null,
          hideExpression: this.isEdit
        }, {
          key: 'Opportunity.ProjectEndDate',
          type: 'datepicker',
          className: 'date-input',
          templateOptions: {
            label: 'Project End Date:',
            required: true,
          },
          defaultValue: this.opportunity?.ProjectEndDate ? new Date(this.opportunity?.ProjectEndDate) : null
        },
        {
          key: 'geographies',
          type: 'ngsearchable',
          templateOptions: {
            label: 'Geographies:',
            placeholder: 'Related geographies and countries',
            required: true,
            multiple: true,
            options: locationsList
          }
        }],
        hideExpression: this.firstStepCompleted
      },
      {
        template: '<div class="form-header">Complete First Stage Info</div><hr />',
        hideExpression: !this.firstStepCompleted,
        expressionProperties: {
          'template': function ($viewValue, $modelValue, scope) {
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
          key: 'StageNumber',
          type: 'select',
          templateOptions: {
            label: 'Start Stage Number:',
            options: stageNumbersList,
            required: true,
          },
          hideExpression: (m, fs) => fs.hideStageNumbers,
        }, {
          key: 'Stage.StageUsersId',
          type: 'ngsearchable',
          templateOptions: {
            label: 'Stage Users:',
            placeholder: 'Stage Users',
            filterLocally: false,
            query: 'siteusers',
            multiple: true,
            required: true,
            options: defaultStageUsersList,
          },
          validation: {
            messages: {
              required: (error) => 'You must enter one or more users',
            },
          },
          defaultValue: this.stage?.StageUsersId
        }, {
          key: 'Stage.StageReview',
          type: 'datepicker',
          templateOptions: {
            label: 'Stage Review',
            required: true
          },
          defaultValue: this.stage?.StageReview ? new Date(this.stage.StageReview) : null
        },],
        hideExpression: !this.firstStepCompleted,
      }
    ];
  }

  onNext() {
    if (this.form.invalid) {
      this.validateAllFormFields(this.form);
      return;
    }
    let optype = this.oppTypes.find(el => el.extra.ID == this.model.Opportunity.OpportunityTypeId);
    if(optype && optype.extra && optype.extra.isInternal) {
      this.onSubmit();
      return;
    }

    if(this.model.Opportunity.OpportunityTypeId)
    this.firstStepCompleted = true;
    this.fields[0].hideExpression = true;
    this.fields[1].hideExpression = this.fields[2].hideExpression = false;

    if (this.data?.forceType) {
      this.options.formState.hideStageNumbers = !this.data.forceType;
    }
  }

  async onSubmit() {
    if (this.form.invalid) {
      this.validateAllFormFields(this.form);
      return;
    }

    if (this.isEdit) {

      this.updating = this.dialogRef.disableClose = true;
      await this.sharepoint.updateOpportunityGeographies(this.data.opportunity, this.model.geographies);
      const success = await this.sharepoint.updateOpportunity(this.data.opportunity.ID, this.model.Opportunity);
      this.updating = this.dialogRef.disableClose = false;
      this.dialogRef.close({
        success: success,
        data: this.model.Opportunity
      });
    } else {
      const newOpp = await this.sharepoint.createOpportunity(this.model.Opportunity, this.model.Stage, this.model.StageNumber);
      if (newOpp) {
        await this.sharepoint.createGeographies(
          newOpp.opportunity.ID,
          this.model.geographies.filter((el: string) => el.startsWith('G-')).map((el: string) => +el.substring(2)),
          this.model.geographies.filter((el: string) => el.startsWith('C-')).map((el: string) => +el.substring(2))
        );
      }
      
      this.dialogRef.close({
        success: newOpp ? true : false,
        data: newOpp
      });
    }
  }
 
  ngOnDestroy(): void {
    this._destroying$.next();
    this._destroying$.complete();
  }

  private validateAllFormFields(formGroup: FormGroup) {
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


