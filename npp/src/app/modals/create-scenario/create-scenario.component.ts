import { Component, Inject, OnInit } from '@angular/core';
import { FormControl, FormGroup } from '@angular/forms';
import { MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { NPPFile, SelectInputList, SharepointService } from 'src/app/services/sharepoint.service';

@Component({
  selector: 'app-create-scenario',
  templateUrl: './create-scenario.component.html',
  styleUrls: ['./create-scenario.component.scss']
})
export class CreateScenarioComponent implements OnInit {

  fields: FormlyFieldConfig[] = [];
  form: FormGroup = new FormGroup({});
  model: any = {};

  scenarios: SelectInputList[] = [];
  file: NPPFile | null = null;

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<CreateScenarioComponent>,
    private readonly sharepoint: SharepointService
  ) { }

  async ngOnInit(): Promise<void> {

    this.file = this.data.file;
    this.scenarios = await this.sharepoint.getScenariosList();

    this.fields = [{
      fieldGroup: [{
        key: 'scenario',
        type: 'select',
        templateOptions: {
            label: 'New Scenario',
            options: this.scenarios,
            required: true
        }
      },
      {
        key: 'comments',
        type: 'textarea',
        templateOptions: {
            label: 'Comments:',
            placeholder: 'Please enter comments',
            rows: 3
        }
      }]
    }];

  }

  async onSubmit() {
    if (this.form.invalid || !this.file) {
      this.validateAllFormFields(this.form);
      return;
    }

    const success = await this.sharepoint.cloneForecastModel(this.file, this.model.scenario, this.model.comments);
    this.dialogRef.close(success);
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
