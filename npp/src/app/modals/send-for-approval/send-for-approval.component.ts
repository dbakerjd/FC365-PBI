import { Component, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { FormlyFieldConfig } from '@ngx-formly/core';

@Component({
  selector: 'app-send-for-approval',
  templateUrl: './send-for-approval.component.html',
  styleUrls: ['./send-for-approval.component.scss']
})
export class SendForApprovalComponent implements OnInit {
  fields: FormlyFieldConfig[] = [{
    fieldGroup: [{
      key: 'comments',
      type: 'textarea',
      templateOptions: {
          label: 'Comments:',
          placeholder: 'Please enter comments',
          rows: 3
      }
    }]
  }];

  form: FormGroup = new FormGroup({});
  model: any; 

  constructor() { }

  ngOnInit(): void {
  }

}
