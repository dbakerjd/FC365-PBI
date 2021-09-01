import { Component, Inject, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { NPPFile, SharepointService } from 'src/app/services/sharepoint.service';

@Component({
  selector: 'app-share-document',
  templateUrl: './share-document.component.html',
  styleUrls: ['./share-document.component.scss']
})
export class ShareDocumentComponent implements OnInit {
  fields: FormlyFieldConfig[] = [];
  form: FormGroup = new FormGroup({});
  model: any = {}; 

  file: NPPFile | undefined = undefined;

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    private readonly sharepoint: SharepointService
  ) { }

  ngOnInit(): void {
    this.file = this.data.file;
    this.fields = [{
      fieldGroup: [{
        key: 'userId',
        type: 'ngsearchable',
        templateOptions: {
            label: 'Stage Users:',
            placeholder: 'Stage Users',
            filterLocally: true,
            options: this.data.folderUsersList,
            multiple: false,
            labelProp: 'Title',
            valueProp: 'Id',
            required: true,
        }
      }]
    }];
  }

  async onSubmit() {
    const fileId = this.file?.ListItemAllFields?.ID;
    if (fileId && this.model.userId) {
      const userFrom = await this.sharepoint.getCurrentUserInfo();
      await this.sharepoint.createNotification(
        this.model.userId, 
        `The file "${this.file?.Name}" was shared with you by ${userFrom.Title}`
      );
    }
  }
}
