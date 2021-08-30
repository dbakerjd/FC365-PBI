import { Component, Inject, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';

@Component({
  selector: 'app-share-document',
  templateUrl: './share-document.component.html',
  styleUrls: ['./share-document.component.scss']
})
export class ShareDocumentComponent implements OnInit {
  fields: FormlyFieldConfig[] = [];

  form: FormGroup = new FormGroup({});
  model: any = {}; 

  constructor(@Inject(MAT_DIALOG_DATA) public data: any) { }

  ngOnInit(): void {
    this.fields = [{
      fieldGroup: [{
        key: 'fileId',
        type: 'input',
        defaultValue: this.data.fileId,
        hideExpression: true
      },{
        key: 'StageUsersId',
        type: 'ngsearchable',
        templateOptions: {
            label: 'Stage Users:',
            placeholder: 'Stage Users',
            filterLocally: false,
            query: 'siteusers',
            multiple: true
        }
      }]
    }];
  }

  onSubmit() {
    console.log(this.model)
  }
}
