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
  modelKeys: number[] = [];
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

    const geographiesList = (await this.sharepoint.getOpportunityGeographies(1)).map(el => { return { label: el.Title, value: el.Id }});

    this.opportunityId = this.data.opportunityId;

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
      },
      {
        key: 'geography',
        type: 'select',
        templateOptions: {
          label: 'Geography:',
          options: geographiesList,
          required: true,
        },
        "hideExpression": (model: any) => {
          return !this.data?.folderList.find((f: NPPFolder) => { return f.DepartmentID === model.category})?.containsModels;
        },
      },
    ];

    let stageGroups = [];
    for (const f of this.data?.folderList) {
      if (f.containsModels) {
        this.modelKeys.push(f.DepartmentID); // needed in onSubmit() to identify the key has subkeys
        for (const geo of geographiesList) {
          stageGroups.push({ 
            departmentID: f.DepartmentID,
            geoID: geo.value,
            group: `DU-${this.opportunityId}-${f.DepartmentID}-${geo.value}`,
            folder: f
          });
        }
      } else {
        stageGroups.push({ 
          departmentID: f.DepartmentID,
          group: `DU-${this.opportunityId}-${f.DepartmentID}`,
          folder: f
        });
      }
    }

    for (const sg of stageGroups) {
      const defaultUsersList: SelectInputList[] = (await this.sharepoint.getGroupMembers(sg.group))
        .map(el => { return { value: el.Id, label: el.Title ? el.Title : '' } });

      // save current users list for department
      this.currentUsersList.push({ 
        departmentID: sg.departmentID, 
        geoID: sg.geoID, 
        list: defaultUsersList.map(el => el.value) 
      });

      // create formly field
      let hideExpression = 'model.category != ' + sg.folder.DepartmentID;
      let formlyKey = 'DepartmentUsersId.' + sg.departmentID;
      if (sg.geoID) {
        hideExpression += ' || model.geography != ' + sg.geoID;
        formlyKey += '.' + sg.geoID;
      }

      formlyFields.push({
        key: formlyKey,
        type: 'ngsearchable',
        templateOptions: {
          label: 'Department Users:',
          placeholder: 'Users with access to ' + sg.folder.Title + ' files',
          filterLocally: false,
          query: 'siteusers',
          multiple: true,
          options: defaultUsersList,
        },
        expressionProperties: {
          'templateOptions.disabled': '!model.category'
        },
        hideExpression: hideExpression,
        defaultValue: defaultUsersList.map(el => el.value)
      });
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
      if (this.modelKeys.includes(+key)) {
        // is a department with geographies
        for (const geoKey in this.model.DepartmentUsersId[key]) {
          const currentList = this.currentUsersList.find(el => el.geoID == geoKey && el.departmentID == key);
          success = success && await this.sharepoint.updateDepartmentUsers(
            this.opportunityId,
            `DU-${this.opportunityId}-${key}-${geoKey}`,
            currentList.list,
            this.model.DepartmentUsersId[key][geoKey]
          );
          if (success) currentList.list = this.model.DepartmentUsersId[key][geoKey]; // update current list
          else break;
        }
      } else {
        const currentList = this.currentUsersList.find(el => el.departmentID == key);
        success = success && await this.sharepoint.updateDepartmentUsers(
          this.opportunityId,
          `DU-${this.opportunityId}-${key}`,
          currentList.list,
          this.model.DepartmentUsersId[key]
        );
        if (success) currentList.list = this.model.DepartmentUsersId[key]; // update current list
        else break;
      }
    }

    this.updating = this.dialogRef.disableClose = false;

    if (success) this.toastr.success('All the Department user permissions has been updated', 'Folder access');
    else this.toastr.error('An error occurred updating users permissions', 'Try Again!');
  }

}
