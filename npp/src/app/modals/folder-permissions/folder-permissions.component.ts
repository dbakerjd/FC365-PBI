import { Component, Inject, OnInit } from '@angular/core';
import { FormControl, FormGroup } from '@angular/forms';
import { MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { ToastrService } from 'ngx-toastr';
import { ErrorService } from '@services/app/error.service';
import { NotificationsService } from 'src/app/services/notifications.service';
import { Opportunity } from '@shared/models/entity';
import { NPPFolder } from '@shared/models/file-system';
import { AppDataService } from '@services/app/app-data.service';
import { PermissionsService } from 'src/app/services/permissions.service';
import { SelectInputList } from '@shared/models/app-config';

@Component({
  selector: 'app-folder-permissions',
  templateUrl: './folder-permissions.component.html',
  styleUrls: ['./folder-permissions.component.scss']
})
export class FolderPermissionsComponent implements OnInit {
  
  form = new FormGroup({});
  model: any = { };
  fields: FormlyFieldConfig[] = [];
  entityId: number = 0;
  stageId: number = 0;
  currentUsersList: any[] = []; // save departments users before changes
  modelKeys: number[] = [];
  loading = true;
  updating = false;
  entity: Opportunity | null = null;

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<FolderPermissionsComponent>,
    private permissions: PermissionsService, 
    private readonly notifications: NotificationsService,
    private readonly toastr: ToastrService,
    private readonly error: ErrorService,
    private readonly appData: AppDataService
  ) {
    this.entity = this.data.entity;
  }

  async ngOnInit() {
    if (!this.entity) return;
  
    const geographiesList = (await this.appData.getEntityGeographies(this.entity.ID))
      .map(el => { return { label: el.Title, value: el.Id } });
    const categoryOptions = this.data?.folderList.map((f: NPPFolder) => {
      return {'name': f.Title, 'value': f.DepartmentID};
    });

    this.entityId = this.entity.ID;
    this.stageId = this.data.stageId;

    let formlyFields: any = [
      {
        key: 'category',
        type: 'select',
        templateOptions: {
          label: 'Department/Folder:',
          options: [...categoryOptions, ...[{'name': 'Reports Only', 'value': 'reports'}]],
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
          return !this.data?.folderList.find((f: NPPFolder) => { return f.DepartmentID === model.category})?.containsModels && model.category !== 'reports';
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
            group: `DU-${this.entityId}-${f.DepartmentID}-${geo.value}`,
            folder: f
          });
        }
      } else {
        stageGroups.push({ 
          departmentID: f.DepartmentID,
          group: `DU-${this.entityId}-${f.DepartmentID}`,
          folder: f
        });
      }
    }

    // Reports Only options
    for (const geo of geographiesList) {
      stageGroups.push({ 
        departmentID: 'reports',
        geoID: geo.value,
        group: `RG-${this.entityId}-${geo.value}`,
        folder: { Title: 'Reports Only', 'DepartmentID': 'reports' }
      });
    }

    for (const sg of stageGroups) {
      const defaultUsersList: SelectInputList[] = (await this.appData.getGroupMembers(sg.group))
        .map(el => { return { value: el.Email, label: el.Title ? el.Title : '' } });

      // save current users list for department
      this.currentUsersList.push({ 
        departmentID: sg.departmentID, 
        geoID: sg.geoID, 
        list: defaultUsersList.map(el => el.value) 
      });

      // create formly field
      let hideExpression = 'model.category != ';
      hideExpression += typeof sg.folder.DepartmentID === 'number' ?  + sg.folder.DepartmentID : `'${sg.folder.DepartmentID}'`;
      let formlyKey = 'DepartmentUsersId.' + sg.departmentID;
      if (sg.geoID) {
        hideExpression += ' || model.geography != ' + sg.geoID;
        formlyKey += '.' + sg.geoID;
      }

      formlyFields.push({
        key: formlyKey,
        type: 'userssearchable',
        templateOptions: {
          label: 'Department Users:',
          placeholder: 'Users with access to ' + sg.folder.Title + ' files',
          multiple: true,
          options: defaultUsersList,
        },
        expressionProperties: {
          'templateOptions.disabled': '!model.category && model.category !== 0'
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
    if (this.form.invalid || !this.entityId || this.updating) {
      return;
    }

    // console.log('model', this.model); return;
    try {

      
      this.updating = this.dialogRef.disableClose = true;

      let success = true;
      for (const key in this.model.DepartmentUsersId) {
        if (this.modelKeys.includes(+key)) {
          // is a department with geographies
          for (const geoKey in this.model.DepartmentUsersId[key]) {
            const currentList = this.currentUsersList.find(el => el.geoID == geoKey && el.departmentID == key);
            success = success && await this.permissions.updateDepartmentUsers(
              this.entityId,
              +key,
              +geoKey,
              currentList.list,
              this.model.DepartmentUsersId[key][geoKey]
            );
            if (success) {
              // notifications
              const addedUsers: string[] = this.model.DepartmentUsersId[key][geoKey].filter((item: string) => currentList.list.indexOf(item) < 0);
              await this.notifications.modelFolderAccessNotification(addedUsers, this.entityId);
              // update current list
              currentList.list = this.model.DepartmentUsersId[key][geoKey];
            }
            else break;
          }
        } else if (key === 'reports') {
          for (const geoKey in this.model.DepartmentUsersId[key]) {
            const currentList = this.currentUsersList.find(el => el.geoID == geoKey && el.departmentID == key);
            success = success && await this.permissions.updateReportsOnlyUsers(
              this.entityId,
              +geoKey,
              currentList.list,
              this.model.DepartmentUsersId[key][geoKey]
            );
          }
        } else {
          const currentList = this.currentUsersList.find(el => el.departmentID == key);
          success = success && await this.permissions.updateDepartmentUsers(
            this.entityId,
            +key,
            null,
            currentList.list,
            this.model.DepartmentUsersId[key]
          );
          if (success) {
            //notifications
            const addedUsers: string[] = this.model.DepartmentUsersId[key].filter((item: string) => currentList.list.indexOf(item) < 0);
            await this.notifications.folderAccessNotification(addedUsers, this.entityId, +key);
            // update current list
            currentList.list = this.model.DepartmentUsersId[key]; 
          }
          else break;
        }
      }

      this.updating = this.dialogRef.disableClose = false;

      if (success) this.toastr.success('All the Department user permissions has been updated', 'Folder access');
      else this.toastr.error('An error occurred updating users permissions', 'Try Again!');
    } catch(e) {
      this.error.handleError(e);
      this.updating = false;
    }
  }
}
