import { Component, Inject, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { ToastrService } from 'ngx-toastr';
import { InlineNppDisambiguationService } from 'src/app/services/inline-npp-disambiguation.service';
import { Brand, Opportunity, SelectInputList, SharepointService } from 'src/app/services/sharepoint.service';

@Component({
  selector: 'app-external-folder-permissions',
  templateUrl: './external-folder-permissions.component.html',
  styleUrls: ['./external-folder-permissions.component.scss']
})
export class ExternalFolderPermissionsComponent implements OnInit {

  form = new FormGroup({});
  model: any = { };
  fields: FormlyFieldConfig[] = [];
  entityId: number = 0;
  entity: Brand | Opportunity | null = null;
  currentUsersList: any[] = []; // save departments users before changes
  modelKeys: number[] = [];
  loading = true;
  updating = false;

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<ExternalFolderPermissionsComponent>,
    private sharepoint: SharepointService,
    private disambiguator: InlineNppDisambiguationService, 
    private readonly toastr: ToastrService
  ) { }

  async ngOnInit() {
    if (!this.data.entity || !this.data.entity.ID) return;

    const geographiesList = (await this.disambiguator.getEntityGeographies(this.data.entity.ID))
      .map(el => { return { label: el.Title, value: el.Id } });

    this.entity = this.data.entity;
    if(this.entity) this.entityId = this.entity.ID;

    let formlyFields: any = [
      {
        key: 'geography',
        type: 'select',
        templateOptions: {
          label: 'Geography:',
          options: geographiesList,
          required: true,
        }
      },
    ];

    let groups = [];
    for (const geo of geographiesList) {
      groups.push({ 
        geoID: geo.value,
        group: `BU-${this.entityId}-${geo.value}`,
        geoName: geo.label
      });
    }

    for (const g of groups) {
      const defaultUsersList: SelectInputList[] = (await this.sharepoint.getGroupMembers(g.group))
        .map(el => { return { value: el.Id, label: el.Title ? el.Title : '' } });

      // save current users list for department
      this.currentUsersList.push({ 
        geoID: g.geoID, 
        list: defaultUsersList.map(el => el.value) 
      });

      let formlyKey = 'GeoUsersId.';
      let hideExpression = '';
      if (g.geoID) {
        hideExpression = 'model.geography != ' + g.geoID;
        formlyKey += g.geoID;
      }

      formlyFields.push({
        key: formlyKey,
        type: 'ngsearchable',
        templateOptions: {
          label: 'Users:',
          placeholder: 'Users with access to ' + g.geoName + ' files',
          filterLocally: false,
          query: 'siteusers',
          multiple: true,
          options: defaultUsersList,
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

    this.updating = this.dialogRef.disableClose = true;

    let success = true;
    for (const key in this.model.GeoUsersId) {
      const currentList = this.currentUsersList.find(el => el.geoID == key);
      success = success && await this.disambiguator.updateEntityGeographyUsers(
        this.entityId,
        +key,
        currentList.list,
        this.model.GeoUsersId[key]
      );
      if (success) currentList.list = this.model.GeoUsersId[key]; // update current list
      else break;
    }

    this.updating = this.dialogRef.disableClose = false;

    if (success) this.toastr.success('All the  permissions has been updated', 'Folder access');
    else this.toastr.error('An error occurred updating users permissions', 'Try Again!');
  }

}
