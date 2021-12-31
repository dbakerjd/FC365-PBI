import { Component, Inject, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
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
    private readonly toastr: ToastrService,
    public dialogRef: MatDialogRef<ShareDocumentComponent>
  ) { }

  ngOnInit(): void {
    this.file = this.data.file;
    this.fields = [{
      fieldGroup: [{
        key: 'usersId',
        type: 'ngsearchable',
        templateOptions: {
            label: 'Document Users:',
            placeholder: 'Users with access',
            filterLocally: true,
            options: this.data.folderUsersList,
            multiple: true,
            labelProp: 'Title',
            valueProp: 'Id',
            required: true,
        }
      }]
    }];
  }

  async onSubmit() {
    const fileId = this.file?.ListItemAllFields?.ID;
    if (fileId && this.model.usersId) {
      const userFrom = await this.sharepoint.getCurrentUserInfo();
      for (const userId of this.model.usersId) {
        await this.sharepoint.createNotification(
          userId, 
          `The file "${this.file?.Name}" was shared with you by ${userFrom.Title}`
        );
      }
      this.toastr.success("The file was shared successfully");
      // else this.toastr.error("The file couldn't be shared", "Try again");
      this.dialogRef.close();
    }
  }
}
