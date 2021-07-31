import { Component, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { FormlyFieldConfig } from '@ngx-formly/core';

@Component({
  selector: 'app-create-scenario',
  templateUrl: './create-scenario.component.html',
  styleUrls: ['./create-scenario.component.scss']
})
export class CreateScenarioComponent implements OnInit {
  fields: FormlyFieldConfig[] = [{
    fieldGroup: [{
      key: 'comments',
      type: 'textarea',
      templateOptions: {
          label: 'Comments:',
          placeholder: 'Please enter comments',
          rows: 3
      }
    },{
      key: 'scenario',
      type: 'select',
      templateOptions: {
          label: 'Scenarios:',
          options: [{
              name: 'Base Case',
              value: '1'
          },{
              name: 'Upside',
              value: '2'
          },{
              name: 'Downside',
              value: '3'
          }],
          valueProp: 'value',
          labelProp: 'name'
      }
    }]
  }];

  form: FormGroup = new FormGroup({});
  model: any; 
  constructor() { }

  ngOnInit(): void {
  }

}
