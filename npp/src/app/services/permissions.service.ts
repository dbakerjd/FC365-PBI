import { Injectable } from '@angular/core';
import { Country, EntityGeography, Opportunity } from '@shared/models/entity';
import { GroupPermission, User } from '@shared/models/user';
import { AppDataService } from './app-data.service';
import { SharepointService } from './sharepoint.service';
import * as SPFolders from '@shared/sharepoint/folders';
import { SystemFolder } from '@shared/models/file-system';
import { GEOGRAPHIES_LIST_NAME } from '@shared/sharepoint/list-names';
import { ToastrService } from 'ngx-toastr';

interface SPGroup {
  Id: number;
  Title: string;
  Description: string;
  LoginName: string;
  OnlyAllowMembersViewMembership: boolean;
}

interface SPGroupListItem {
  type: string;
  data: SPGroup;
}

@Injectable({
  providedIn: 'root'
})
export class PermissionsService {

  constructor(
    // private readonly sharepoint: SharepointService,
    private readonly appData: AppDataService,
    private readonly toastr: ToastrService
  ) { }


}
