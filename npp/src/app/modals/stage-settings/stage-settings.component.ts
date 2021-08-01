import { Component, Inject, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { Gate } from 'src/app/services/sharepoint.service';

@Component({
  selector: 'app-stage-settings',
  templateUrl: './stage-settings.component.html',
  styleUrls: ['./stage-settings.component.scss']
})
export class StageSettingsComponent implements OnInit {
  fields: FormlyFieldConfig[] = [{
    fieldGroup: [{
      key: 'users',
      type: 'input',
      templateOptions: {
          label: 'Stage Users:',
          placeholder: 'Stage Users'
      }
    },{
      key: 'review',
      type: 'datepicker',
      templateOptions: {
          label: 'Stage Review:'
      }
    }]
  }];

  form: FormGroup = new FormGroup({});
  model: any;   
  gate: Gate | undefined;

  constructor(@Inject(MAT_DIALOG_DATA) public data: any) { }

  ngOnInit(): void {
    this.gate = this.data.gate;
  }
}
