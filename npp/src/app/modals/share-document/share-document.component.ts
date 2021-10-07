import { Component, Inject, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { ToastrService } from 'ngx-toastr';
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
    private readonly sharepoint: SharepointService,
    private readonly toastr: ToastrService
  ) { }

  ngOnInit(): void {
    this.file = this.data.file;
    this.fields = [{
      fieldGroup: [{
        key: 'userId',
        type: 'ngsearchable',
        templateOptions: {
            label: 'Document Users:',
            placeholder: 'Users with access',
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
      const created = await this.sharepoint.createNotification(
        this.model.userId, 
        `The file "${this.file?.Name}" was shared with you by ${userFrom.Title}`
      );
      if (created) this.toastr.success("The file was shared successfully");
      else this.toastr.error("The file couldn't be shared", "Try again")
    }
  }
}
