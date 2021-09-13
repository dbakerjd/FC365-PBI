import { Component, Inject, OnInit } from '@angular/core';
import { FormControl, FormGroup } from '@angular/forms';
import { MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { ToastrService } from 'ngx-toastr';
import { NPPFolder, SelectInputList, SharepointService } from 'src/app/services/sharepoint.service';

@Component({
  selector: 'app-folder-permissions',
  templateUrl: './folder-permissions.component.html',
  styleUrls: ['./folder-permissions.component.scss']
})
export class FolderPermissionsComponent implements OnInit {
  
  form = new FormGroup({});
  model: any = { };
  fields: FormlyFieldConfig[] = [];
  opportunityId: number | null = null;
  currentUsersList: any[] = []; // save departments users before changes
  loading = true;
  updating = false;

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<FolderPermissionsComponent>,
    private sharepoint: SharepointService, 
    private readonly toastr: ToastrService
  ) { }

  async ngOnInit() {
    if (!this.data.opportunityId) return;

    this.opportunityId = this.data.opportunityId; // pot ser només ID?

    let formlyFields: any = [
      {
        key: 'category',
        type: 'select',
        templateOptions: {
          label: 'Department/Folder:',
          options: this.data?.folderList.map((f: NPPFolder) => {
            return {
              'name': f.Title,
              'value': f.DepartmentID,
            };
          }),
          valueProp: 'value',
          labelProp: 'name',
          required: true,
        }
      }
    ];

    for (const f of this.data?.folderList) {
      const DUGroupName = `DU-${this.opportunityId}-${f.DepartmentID}`;
      const defaultUsersList: SelectInputList[] = (await this.sharepoint.getGroupMembers(DUGroupName))
        .map(el => { return { value: el.Id, label: el.Title ? el.Title : '' }});

      // save current users list for department
      this.currentUsersList[f.DepartmentID] = defaultUsersList.map(el => el.value);
      
      formlyFields.push({
        key: 'DepartmentUsersId.' + f.DepartmentID,
        type: 'ngsearchable',
        templateOptions: {
            label: 'Department Users:',
            placeholder: 'Users with access to ' + f.Title + ' files',
            filterLocally: false,
            query: 'siteusers',
            multiple: true,
            options: defaultUsersList,
        },
        expressionProperties: {
          'templateOptions.disabled': '!model.category'
        },
        hideExpression: 'model.category != ' + f.DepartmentID,
        defaultValue: defaultUsersList.map(el => el.value)
      })
    }

    this.fields = [
      {
        fieldGroup: formlyFields
      }
    ];

    this.loading = false;
  }

  async onSubmit() {
    if (this.form.invalid || !this.opportunityId || this.updating) {
      return;
    }

    this.updating = this.dialogRef.disableClose = true;
    let success = true;
    for (const key in this.model.DepartmentUsersId) {
      success = success && await this.sharepoint.updateDepartmentUsers(
        this.opportunityId, 
        `DU-${this.opportunityId}-${key}`, 
        this.currentUsersList[+key], 
        this.model.DepartmentUsersId[key]
      );
      if (success) this.currentUsersList[+key] = this.model.DepartmentUsersId[key]; // update current list
      else break;
    }

    this.updating = this.dialogRef.disableClose = false;

    if (success) this.toastr.success('All the Department user permissions has been updated', 'Folder access');
    else this.toastr.error('An error occurred updating users permissions', 'Try Again!');
  }

}
