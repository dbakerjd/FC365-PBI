import { Component, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { OpportunityInput, SharepointService } from 'src/app/services/sharepoint.service';
import { takeUntil, tap } from 'rxjs/operators';
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


  constructor(private sharepoint: SharepointService, public matDialog: MatDialog) { }

  async ngOnInit() {

    let therapies = await this.sharepoint.getTherapiesList();
    let oppTypes = await this.sharepoint.getOpportunityTypesList();

    this.fields = [{
      fieldGroup: [{
        key: 'Title',
        type: 'input',
        templateOptions: {
          label: 'Opportunity Name:',
          placeholder: 'Opportunity Name'
        }
      }, {
        key: 'MoleculeName',
        type: 'input',
        templateOptions: {
          label: 'Molecule Name:',
          placeholder: 'Molecule Name'
        }
      }, {
        key: 'OpportunityOwnerId',
        type: 'input',
        templateOptions: {
          label: 'Opportunity Owner:',
          placeholder: 'Opportunity Owner'
        }
      }, {
        key: 'therapy',
        type: 'select',
        templateOptions: {
          label: 'Therapy Area:',
          options: therapies,
          /*
          change: (field) =>{ 
            console.log('field', field);
            console.log('field0', field.model);
            console.log('field1', field.form?.controls.therapy);
            console.log('field2', field.formControl?.value);
            // console.log($event);
            const tabletField = field.formControl?.parent?.get('indication');
            console.log('ind field', tabletField);
            // getField('nome', this.fields).templateOptions.options = this.newOptions;
            this.sharepoint.getIndicationsList(field.model.therapy).then((r) => {
              console.log('gagsaga');
              console.log(r);
              console.log(field.form);
              console.log(this.fields);
              let f = this.getField('indication', this.fields);
              if (f && f.templateOptions) {
                f.templateOptions.options = r;
              }
            });
            console.log('indic ch', this.indications);
            // if (field.form)
            // this.setIndications(field.form.controls.editor.value);
          }
          */
          
        //  change: this.setIndications
        },
      }, {
        key: 'IndicationId',
        type: 'select',
        templateOptions: {
          label: 'Indication Name:',
          options: [],
        },
        lifecycle: {
          onInit: (form, field) => {
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
        key: 'OpportunityTypeId',
        type: 'select',
        templateOptions: {
          label: 'Opportunity Type:',
          options: oppTypes,
        }
      }, {
        key: 'ProjectStartDate',
        type: 'datepicker',
        templateOptions: {
          label: 'Project Start Date:'
        }
      }, {
        key: 'ProjectEndDate',
        type: 'datepicker',
        templateOptions: {
          label: 'Project End Date:'
        }
      }]
    }];
  }

  onSubmit() {
    console.log(this.model);
    delete this.model.therapy;
    let newOpportunity: OpportunityInput = { ...this.model };
    this.dialogInstance = this.matDialog.open(StageSettingsComponent, {
      height: '700px',
      width: '405px'
    });
    // this.sharepoint.createOpportunity(newOpportunity);
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


