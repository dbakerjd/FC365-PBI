import { Component, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { FormlyFieldConfig } from '@ngx-formly/core';

@Component({
  selector: 'app-create-opportunity',
  templateUrl: './create-opportunity.component.html',
  styleUrls: ['./create-opportunity.component.scss']
})
export class CreateOpportunityComponent implements OnInit {
  fields: FormlyFieldConfig[] = [{
    fieldGroup: [{
      key: 'name',
      type: 'input',
      templateOptions: {
          label: 'Opportunity Name:',
          placeholder: 'Opportunity Name'
      }
    },{
      key: 'molecule',
      type: 'input',
      templateOptions: {
          label: 'Molecule Name:',
          placeholder: 'Molecule Name'
      }
    },{
      key: 'owner',
      type: 'input',
      templateOptions: {
          label: 'Opportunity Owner:',
          placeholder: 'Opportunity Owner'
      }
    },{
      key: 'therapy',
      type: 'select',
      templateOptions: {
          label: 'Therapy Area:',
          options: [{
              name: 'Bone / Osteoporosis',
              value: '1'
          },{
              name: 'Cardiovascular',
              value: '2'
          },{
              name: 'Dermatology',
              value: '3'
          }],
          valueProp: 'value',
          labelProp: 'name'
      }
    },{
      key: 'indication',
      type: 'select',
      templateOptions: {
          label: 'Indication Name:',
          options: [{
              name: 'Bone Cancer',
              value: '1'
          },{
              name: 'Bone Density',
              value: '2'
          },{
              name: 'Bone Infections',
              value: '3'
          }],
          valueProp: 'value',
          labelProp: 'name'
      }
    },{
      key: 'type',
      type: 'select',
      templateOptions: {
          label: 'Opportunity Type:',
          options: [{
              name: 'Acquisition',
              value: '1'
          },{
              name: 'Licensing',
              value: '2'
          },{
              name: 'Product Development',
              value: '3'
          }],
          valueProp: 'value',
          labelProp: 'name'
      }
    },{
      key: 'start_date',
      type: 'datepicker',
      templateOptions: {
          label: 'Project Start Date:'
      }
    },{
      key: 'end_date',
      type: 'datepicker',
      templateOptions: {
          label: 'Project End Date:'
      }
    }]
  }];

  form: FormGroup = new FormGroup({});
  model: any; 
  constructor() { }

  ngOnInit(): void {
  }

}
