import { Component, Inject, OnInit } from '@angular/core';
import { FormControl, FormGroup } from '@angular/forms';
import { MatDialog, MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { ErrorService } from 'src/app/services/error.service';
import { WorkInProgressService } from 'src/app/services/work-in-progress.service';
import { Opportunity } from '@shared/models/entity';
import { AppDataService } from 'src/app/services/app-data.service';
import { EntitiesService } from 'src/app/services/entities.service';
import { SelectInputList } from '@shared/models/app-config';


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
  cycles: SelectInputList[] = [];

  // flow control
  updating = false;

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<CreateForecastCycleComponent>,
    public matDialog: MatDialog,
    private error: ErrorService,
    public jobs: WorkInProgressService,
    private readonly appData: AppDataService,
    private readonly entities: EntitiesService
  ) { }

  async ngOnInit(): Promise<void> {

    this.entity = this.data.entity;
    this.cycles = await this.appData.getForecastCycles();
    const currentYear = new Date().getFullYear();
    
    let year = currentYear;
    let elegibleYears = [currentYear];
    for(let i=1; i<6; i++) {
      elegibleYears.push(++year);
    }

    this.fields = [{
      fieldGroup: [{
        key: 'ForecastCycle',
        type: 'select',
        templateOptions: {
            label: 'Forecast Cycle Type',
            options: this.cycles,
            required: true,
            multiple: false
        }
      },{
        key: 'Year',
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
        defaultValue: this.entity?.Year || currentYear
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
      "Creating Forecast Cycle"
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
}
