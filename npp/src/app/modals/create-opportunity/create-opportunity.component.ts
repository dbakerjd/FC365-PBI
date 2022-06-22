import { Component, Inject, OnInit } from '@angular/core';
import { FormControl, FormGroup } from '@angular/forms';
import { FormlyFieldConfig, FormlyFormOptions } from '@ngx-formly/core';
import { take, takeUntil, tap } from 'rxjs/operators';
import { Subject } from 'rxjs';
import { MatDialog, MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { EntityGeography, Opportunity, Stage } from '@shared/models/entity';
import { AppDataService } from '@services/app/app-data.service';
import { PermissionsService } from 'src/app/services/permissions.service';
import { EntitiesService } from 'src/app/services/entities.service';
import { SelectInputList } from '@shared/models/app-config';
import { SelectListsService } from '@services/select-lists.service';
import { StringMapperService } from '@services/string-mapper.service';

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
      // hideStageNumbers: true,
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
  isInternal: boolean = false;

  constructor(
    private permissions: PermissionsService, 
    private readonly entities: EntitiesService,
    private readonly appData: AppDataService,
    private readonly selectLists: SelectListsService,
    private readonly stringMapper: StringMapperService,
    public matDialog: MatDialog,
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<CreateOpportunityComponent>
    ) { 
      dialogRef.disableClose = true;
    }

  async ngOnInit() {
    this.opportunity = this.data?.opportunity;
    this.isEdit = this.data?.opportunity && !this.data?.createFrom;

    const therapies = await this.selectLists.getTherapiesList();
    let forecastCycles = await this.selectLists.getForecastCyclesList();
    this.oppTypes = await this.selectLists.getOpportunityTypesList();
    const geo = (await this.selectLists.getGeographiesList()).map(el => { return { label: el.label, value: 'G-' + el.value } });
    const countries = (await this.selectLists.getCountriesList()).map(el => { return { label: el.label, value: 'C-' + el.value } });;
    const locationsList = geo.concat(countries);
    let indicationsList: SelectInputList[] = [];
    let businessUnits = await this.selectLists.getBusinessUnitsList();
    // let stageNumbersList: SelectInputList[] = [];
    let defaultUsersList: SelectInputList[] = await this.selectLists.getSiteOwnersList();
    let defaultStageUsersList: SelectInputList[] = [];
    this.firstStepCompleted = false;
    const trialPhases = await this.selectLists.getClinicalTrialPhases();
    const currentYear = new Date().getFullYear();
    let year = currentYear;
    let elegibleYears = [currentYear];
    for(let i=1; i<6; i++) {
      elegibleYears.push(++year);
    }
    
    if (this.opportunity) {
      let type = this.oppTypes.find(el => el.value == this.opportunity?.OpportunityTypeId);
      this.isInternal = type.extra?.IsInternal;
      this.geographies = await this.appData.getEntityGeographies(this.opportunity?.ID);
      this.model.geographies = this.geographies.map(el => el.CountryId ? 'C-'+el.CountryId : 'G-' + el.GeographyId);
    
      if (this.data?.forceInternal) { // force Phase opportunity (complete opportunity option)
        this.isInternal = true;
        this.oppTypes = this.oppTypes.filter(el => el.extra.IsInternal);
        this.opportunity.OpportunityTypeId = -1;
        if (this.oppTypes.length > 0) {
          this.opportunity.OpportunityTypeId = this.oppTypes[0].value;
          // this.model.stageType = 'Phase';
          // stageNumbersList = await this.appData.getMasterStageNumbers('Phase');
        }
      }

      // default indications for the therapy selected
      if (this.opportunity && this.opportunity.Indication && this.opportunity.Indication.length) {
        indicationsList = await this.selectLists.getIndicationsList(this.opportunity.Indication[0].TherapyArea);
      }
      // if we are cloning opportunity, get first stage info
      if (this.data?.createFrom && !this.data?.forceInternal) {
        this.stage = await this.appData.getFirstStage(this.opportunity);
        if (this.stage) {
          defaultStageUsersList = await this.selectLists.getUsersList(this.stage.StageUsersId);
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
            required: true,
          },
          defaultValue: this.opportunity?.Title
        }, {
          key: 'Opportunity.MoleculeName',
          type: 'input',
          templateOptions: {
            label: 'Molecule Name:',
            required: true,
          },
          defaultValue: this.opportunity?.MoleculeName
        }, {
          key: 'Opportunity.EntityOwnerId',
          type: 'ngsearchable',
          templateOptions: {
            label: 'Opportunity Owner:',
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
            label: this.stringMapper.getString('Therapy Area') + ':',
            options: therapies,
            required: true,
          },
          defaultValue: this.opportunity && this.opportunity.Indication && this.opportunity.Indication.length ? this.opportunity.Indication[0].TherapyArea : null
        },{
          key: 'Opportunity.IndicationId',
          type: 'ngsearchable',
          templateOptions: {
            label: this.stringMapper.getString('Indications') + ':',
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
                  this.selectLists.getIndicationsList(th).then(r => {
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
            label: this.stringMapper.getString('Business Unit') + ':',
            options: businessUnits,
            required: true,
          },
          defaultValue: this.opportunity?.BusinessUnitId,
          hideExpression: this.isEdit
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
                    let t = this.oppTypes.find(el => el.value == selectedValue);
                    this.isInternal = t ? t.extra.IsInternal : false;
                    this.appData.getStageType(selectedValue).then(r => {
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
          className: 'date-input firstHalf',
          templateOptions: {
            label: 'Project Start Date:',
            required: true,
          },
          defaultValue: this.opportunity?.ProjectStartDate ? new Date(this.opportunity?.ProjectStartDate) : null,
          hideExpression: () => this.isEdit || this.isInternal
        }, {
          key: 'Opportunity.ProjectEndDate',
          type: 'datepicker',
          className: 'date-input secondHalf',
          templateOptions: {
            label: 'Project End Date:',
            required: true,
          },
          defaultValue: this.opportunity?.ProjectEndDate ? new Date(this.opportunity?.ProjectEndDate) : null,
          hideExpression: () => this.isInternal
        }, {
          key: 'Opportunity.ForecastCycleId',
          type: 'select',
          templateOptions: {
            label: 'Forecast Cycle:',
            options: forecastCycles,
            required: true,
          },
          defaultValue: this.opportunity?.ForecastCycleId,
          hideExpression: () => !this.isInternal
        }, {
          key: 'Opportunity.Year',
          type: 'select',
          templateOptions: {
            label: 'Year:',
            options: elegibleYears.map(el => {
              return {
                label: el,
                value: el
              }
            }),
            required: true,
          },
          defaultValue: this.opportunity?.Year || currentYear,
          hideExpression: () => !this.isInternal
        }, {
          key: 'Opportunity.ForecastCycleDescriptor',
          type: 'input',
          templateOptions: {
            label: 'Forecast Cycle Descriptor:',
            required: false,
          },
          defaultValue: this.opportunity?.ForecastCycleDescriptor,
          hideExpression: () => !this.isInternal
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
        }, /*{
          key: 'StageNumber',
          type: 'select',
          templateOptions: {
            label: 'Start Stage Number:',
            options: stageNumbersList,
            required: true,
          },
          hideExpression: (m, fs) => fs.hideStageNumbers,
        }, */{
          key: 'StageUsersMails',
          type: 'userssearchable',
          templateOptions: {
            label: 'Stage Users:',
            placeholder: 'Stage Users',
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
    if(optype && optype.extra && optype.extra.IsInternal) {
      this.onSubmit();
      return;
    }

    if(this.model.Opportunity.OpportunityTypeId)
    this.firstStepCompleted = true;
    this.fields[0].hideExpression = true;
    this.fields[1].hideExpression = this.fields[2].hideExpression = false;

    // if (this.data?.forceInternal) {
    //   this.options.formState.hideStageNumbers = !this.data.forceInternal;
    // }
  }

  async onSubmit() {
    if (this.form.invalid) {
      this.validateAllFormFields(this.form);
      return;
    }

    if (this.isEdit) {

      this.updating = this.dialogRef.disableClose = true;
      await this.permissions.updateEntityGeographies(this.data.opportunity, this.model.geographies);
      const success = await this.entities.updateEntity(this.data.opportunity.ID, this.model.Opportunity);
      this.updating = this.dialogRef.disableClose = false;
      this.dialogRef.close({
        success: success,
        data: this.model.Opportunity
      });
    } else {
      const newOpp = await this.entities.createOpportunity(this.model.Opportunity, this.model.Stage);
      if (newOpp) {
        await this.permissions.createGeographies(
          newOpp.opportunity.ID,
          this.model.geographies.filter((el: string) => el.startsWith('G-')).map((el: string) => +el.substring(2)),
          this.model.geographies.filter((el: string) => el.startsWith('C-')).map((el: string) => +el.substring(2))
        );
      }
      
      this.dialogRef.close({
        success: newOpp ? true : false,
        data: newOpp,
        users: this.model.StageUsersMails
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


