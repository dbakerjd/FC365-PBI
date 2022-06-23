import { Component, Inject, OnInit } from '@angular/core';
import { FormControl, FormGroup } from '@angular/forms';
import { MatDialog, MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { ErrorService } from '@services/app/error.service';
import { WorkInProgressService } from '@services/app/work-in-progress.service';
import { EntityForecastCycle, Opportunity } from '@shared/models/entity';
import { EntitiesService } from 'src/app/services/entities.service';
import { SelectInputList } from '@shared/models/app-config';
import { SelectListsService } from '@services/select-lists.service';
import { takeUntil, tap } from 'rxjs/operators';
import { Subject } from 'rxjs';
import { StringMapperService } from '@services/string-mapper.service';

@Component({
  selector: 'app-create-forecast-cycle',
  templateUrl: './create-forecast-cycle.component.html',
  styleUrls: ['./create-forecast-cycle.component.scss']
})
export class CreateForecastCycleComponent implements OnInit {

  fields: FormlyFieldConfig[] = [];
  form: FormGroup = new FormGroup({});
  model: any = {};
  entity: Opportunity | undefined;
  cyclesList: SelectInputList[] = [];
  yearsOptions: any[] = [];
  private readonly _destroying$ = new Subject<void>();

  // flow control
  updating = false;

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<CreateForecastCycleComponent>,
    public matDialog: MatDialog,
    private error: ErrorService,
    public jobs: WorkInProgressService,
    private readonly entities: EntitiesService,
    private readonly selectLists: SelectListsService,
    public readonly stringMapper: StringMapperService
  ) { }

  async ngOnInit(): Promise<void> {

    this.entity = this.data.entity;
    const usedCycles: EntityForecastCycle[] = this.data.cycles;
    this.cyclesList = await this.selectLists.getForecastCyclesList();
    const currentYear = new Date().getFullYear();
    
    let year = currentYear;
    let elegibleYears = [currentYear];
    for(let i=1; i<6; i++) {
      elegibleYears.push(++year);
    }

    // remove years already used in every forecast cycle type
    for (const c of this.cyclesList) {
      const yearsUsed: number[] = usedCycles.filter(uc => uc.ForecastCycleTypeId === c.value).map(e => +e.Year);
      this.yearsOptions[c.value] = elegibleYears.filter(y => !yearsUsed.includes(y));
    }

    this.fields = [{
      fieldGroup: [{
        key: 'ForecastCycle',
        type: 'select',
        templateOptions: {
            label: this.stringMapper.getString('Forecast Cycle') + ' Type',
            options: this.cyclesList,
            required: true,
            multiple: false
        }
      },{
        key: 'Year',
        type: 'select',
        templateOptions: {
          label: this.stringMapper.getString('FC Year') + ':',
          options: [],
          required: true,
        },
        defaultValue: this.entity?.Year || currentYear,
        hooks: {
          onInit: (field) => {
            if (!field?.parent?.fieldGroup) return;
            const cycleSelect = field.parent.fieldGroup.find(f => f.key === 'ForecastCycle');
            if (!cycleSelect?.formControl) return;

            // initial value
            if (cycleSelect.formControl.value) {
              field.templateOptions!.options = this.yearsOptions[cycleSelect.formControl.value].map((el: number) => {
                return {
                  label: el,
                  value: el
                }
              });
            }
            
            // subscription to value changes
            cycleSelect.formControl.valueChanges.pipe(
              takeUntil(this._destroying$),
              tap(cycleId => {
                  if (field.templateOptions) field.templateOptions.options = this.yearsOptions[cycleId].map((el: number) => {
                    return {
                      label: el,
                      value: el
                    }
                  });
              }),
            ).subscribe();
          }
        }
      },{
        key: 'ForecastCycleDescriptor',
        type: 'input',
        templateOptions: {
          label: 'Descriptor',
          required: false
        }
      }]
    }];

  }

  async onSubmit() {
    let job = this.jobs.startJob(
      "Creating " + this.stringMapper.getString('Forecast Cycle')
      );
    try {
      if (this.form.invalid || !this.entity) {
        this.validateAllFormFields(this.form);
        this.jobs.finishJob(job.id);
        return;
      }
      
      this.updating = this.dialogRef.disableClose = true;
      let success = await this.entities.createEntityForecastCycle(this.entity, this.form.value);
      this.jobs.finishJob(job.id);
      this.updating = this.dialogRef.disableClose = false;
      this.dialogRef.close(success);
    } catch(e) {
      this.jobs.finishJob(job.id);
      this.error.handleError(e);
      this.updating  = this.dialogRef.disableClose = false;
    }
    
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

  ngOnDestroy(): void {
    this._destroying$.next();
    this._destroying$.complete();
  }
}
