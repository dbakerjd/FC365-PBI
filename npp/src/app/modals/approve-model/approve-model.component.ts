import { Component, Inject, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { NotificationsService } from 'src/app/services/notifications.service';
import { Opportunity } from '@shared/models/entity';
import { NPPFile } from '@shared/models/file-system';
import { FilesService } from 'src/app/services/files.service';
import { StringMapperService } from '@services/string-mapper.service';

@Component({
  selector: 'app-approve-model',
  templateUrl: './approve-model.component.html',
  styleUrls: ['./approve-model.component.scss']
})
export class ApproveModelComponent implements OnInit {

  file: NPPFile | null = null;
  entity: Opportunity | null = null; 
  rootFolder: string = "";
  approving = false;
  departmentID: number = 0;
  currentGate: any = null;

  fields: FormlyFieldConfig[] = [{
    fieldGroup: [{
      key: 'comments',
      type: 'textarea',
      templateOptions: {
          label: 'Comments:',
          placeholder: 'Please enter comment.',
          rows: 3
      }
    }]
  }];

  form: FormGroup = new FormGroup({});
  model: any = {}; 

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<ApproveModelComponent>,
    private readonly notifications: NotificationsService,
    private readonly files: FilesService,
    public readonly stringMapper: StringMapperService
  ) { }

  ngOnInit(): void {
    this.file = this.data.file;
    this.entity = this.data.entity;
    this.rootFolder = this.data.rootFolder;
    this.departmentID = this.data.DepartmentID;
    this.currentGate = this.data.currentGate;
  }

  async onSubmit() {
    try {
      if (this.file && this.entity) {
        let commentsStr = '';
        this.approving = true;
        if(this.model.comments) {
          commentsStr = await this.files.addFileComment(this.file, this.model.comments);
        }
        const result = await this.files.setFileApprovalStatus(this.rootFolder, this.file, this.entity, "Approved", commentsStr);
        let groups = [
          `DU-${this.entity.ID}-${this.departmentID}-${this.file.ListItemAllFields?.EntityGeographyId}`,
          `OO-${this.entity.ID}`
        ];
        if(this.currentGate) {
          groups.push(`SU-${this.entity.ID}-${this.currentGate?.StageNameId}`);
        }
        await this.notifications.modelApprovedNotification(this.file.Name, this.entity.ID, groups);
        this.approving = false;
        this.dialogRef.close({
          success: result,
          comments: commentsStr
        });
      }
    } catch(e) {
      this.approving = false;
    }
    
  }


}
