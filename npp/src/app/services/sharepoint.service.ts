import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Observable, of } from 'rxjs';
import { catchError, filter } from 'rxjs/operators';
import { ErrorService } from './error.service';
import { LicensingService } from './licensing.service';
import { TeamsService } from './teams.service';
import { map } from 'rxjs/operators';


export interface Opportunity {
  ID: number;
  Title: string;
  MoleculeName: string;
  OpportunityOwnerId: number;
  OpportunityOwner?: User;
  ProjectStartDate: Date;
  ProjectEndDate: Date;
  OpportunityTypeId: number;
  OpportunityType?: OpportunityType;
  OpportunityStatus: "Processing" | "Archive" | "Active" | "Approved";
  IndicationId: number;
  Indication: Indication;
  Modified: Date;
  AuthorId: number;
  Author?: User;
  // users?: User[];
  progress?: number;
}

export interface OpportunityInput {
  Title: string;
  MoleculeName: string;
  OpportunityOwnerId: number;
  ProjectStartDate: Date;
  ProjectEndDate: Date;
  OpportunityTypeId: number;
  IndicationId: number;
}

export interface StageInput {
  Title: string;
  StageUsersId: number[];
  StageReview: Date;
  OpportunityNameId?: number;
  StageNameId?: number;
}

export interface Action {
  Id: number,
  StageNameId: number;
  OpportunityNameId: number;
  Title: string;
  ActionNameId: string;
  ActionDueDate: Date;
  Complete: boolean;
  Timestamp: Date;
  TargetUserId: Number;
  TargetUser: User;
  status?: string;
}

export interface MasterAction {
  Id: number,
  Title: string;
  ActionNumber: number;
  StageNameId: number;
  OpportunityTypeId: number;
  DueDays: number;
}

export interface OpportunityType {
  ID: number;
  Title: string;
  StageType: string;
}

export interface Indication {
  ID: number;
  Title: string;
  TherapyArea: string;
}

export interface User {
  Id: number;
  LoginName?: string;
  FirstName?: string;
  LastName?: string;
  Title?: string;
  Email?: string;
  profilePicUrl?: string;
}

export interface Stage {
  ID: number;
  Title: string;
  OpportunityNameId: number;
  StageNameId: number;
  StageReview: Date;
  StageUsersId: number[];
  Created: Date;
  actions?: Action[];
  folders?: NPPFolder[];
}

export interface NPPFile {
  Name: string;
  ServerRelativeUrl: string;
  TimeLastModified: Date;
  ListItemAllFields?: NPPFileMetadata;
}

export interface NPPFileMetadata {
  ID: number;
  ApprovalStatusId: number;
  OpportunityNameId: number;
  StageNameId: number;
  ModelApprovalComments: string;
  ApprovalStatus: string;
  CountryId?: number[];
  ModelScenarioId?: number[];
  AuthorId: number;
  Author: User;
  TargetUserId: number;
  TargetUser?: User;
}

export interface NPPFolder {
  ID: number;
  Title: string;
  StageNameId: number;
  containsModels?: boolean;
}

export interface Country {
  ID: number;
  Title: string;
}

export interface SelectInputList {
  label: string;
  value: any;
  group?: string;
}

export interface SharepointResult {
  'odata.metadata': string;
  value: any;
}

export interface FilterTerm {
  term: string;
  field?: string;
  matchCase?: boolean;
}

export interface SPGroup {
  Id: number;
  Title: string;
  Description: string;
  LoginName: string;
  OnlyAllowMembersViewMembership: boolean;
}

export interface SPGroupListItem {
  type: string;
  data: SPGroup;
}

export interface GroupPermission {
  Id: number;
  Title: string;
  ListName: string;
  Permission: string;
  ListFilter: 'List' | 'Item';
}

const OPPORTUNITES_LIST_NAME = 'Opportunities';
const OPPORTUNITY_STAGES_LIST_NAME = 'Opportunity Stages';
const OPPORTUNITY_ACTIONS_LIST_NAME = 'Opportunity Action List';
const OPPORTUNITIES_LIST = "lists/getbytitle('"+OPPORTUNITES_LIST_NAME+"')";
const OPPORTUNITY_STAGES_LIST = "lists/getbytitle('"+OPPORTUNITY_STAGES_LIST_NAME+"')";
const OPPORTUNITY_ACTIONS_LIST = "lists/getbytitle('"+OPPORTUNITY_ACTIONS_LIST_NAME+"')";
const MASTER_OPPORTUNITY_TYPES_LIST = "lists/getbytitle('Master Opportunity Type List')";
const MASTER_THERAPY_AREAS_LIST = "lists/getbytitle('Master Therapy Areas')";
const MASTER_STAGES_LIST = "lists/getbytitle('Master Stage List')";
const MASTER_ACTION_LIST = "lists/getbytitle('Master Action List')";
const MASTER_FOLDER_LIST = "lists/getByTitle('Master Folder List')";
const MASTER_GROUP_TYPES_LIST = "lists/getByTitle('Master Group Types List')";
const COUNTRIES_LIST = "lists/getByTitle('Countries')";
const MASTER_SCENARIOS_LIST = "lists/getByTitle('Master Scenarios')";
const USER_INFO_LIST = "lists/getByTitle('User Information List')";
const NOTIFICATIONS_LIST = "lists/getByTitle('Notifications')";
const FILES_FOLDER = "Current Opportunity Library";

@Injectable({
  providedIn: 'root'
})
export class SharepointService {

  // local "cache"
  masterOpportunitiesTypes: OpportunityType[] = [];
  masterGroupTypes: GroupPermission[] = [];
  masterCountriesList: SelectInputList[] = [];
  masterScenariosList: SelectInputList[] = [];
  masterTherapiesList: SelectInputList[] = [];
  masterIndications: {
    therapy: string;
    indications: Indication[]
  }[] = [];
  masterFolders: {
    stage: number;
    folders: NPPFolder[]
  }[] = [];

  constructor(private teams: TeamsService, private http: HttpClient, private error: ErrorService, private licensing: LicensingService) { }

  query(partial: string, conditions: string = '', count: number | 'all' = 'all', filter?: FilterTerm): Observable<any> {
    //TODO implement usage of count

    let filterUri = '';
    if (filter && filter.term) {
      filter.field = filter.field ? filter.field : 'Title';
      filter.matchCase = filter.matchCase ? filter.matchCase : false;

      if (filter.matchCase) {
        filterUri = `$filter=substringof('${filter.term}',${filter.field})`;
      } else {
        let capitalized = filter.term.charAt(0).toUpperCase() + filter.term.slice(1);
        filterUri = `$filter=substringof('${filter.term}',${filter.field}) or substringof('${capitalized}',${filter.field})`;
      }
    }
    let endpoint = this.licensing.getSharepointUri() + partial;
    if (conditions || filterUri) endpoint += '?';
    if (conditions) endpoint += conditions;
    if (filterUri) endpoint += conditions ? '&' + filterUri : filterUri;
    console.log('endpoint query', endpoint);
    try {
      return this.http.get(endpoint);
    } catch (e) {
      if(e.status == 401) {
        // await this.teams.refreshToken(true); 
      }
      return of([]);
    }
  }

  async getAllItems(list: string, conditions: string = ''): Promise<any[]> {
    try {
      let endpoint = this.licensing.getSharepointUri() + list + '/items';
      if (conditions) endpoint += '?' + conditions;
      let lists = await this.http.get(endpoint).toPromise() as SharepointResult; 
      if (lists.value && lists.value.length > 0) {
        return lists.value;
      }
      return [];
    } catch (e) {
      if(e.status == 401) {
        // await this.teams.refreshToken(true); 
      }
      return [];
    }
  }

  async getOneItem(list: string, conditions: string = ''): Promise<any> {
    try {
      let endpoint = this.licensing.getSharepointUri() + list + '/items';
      if (conditions) endpoint += '?' + conditions;
      let lists = await this.http.get(endpoint).toPromise() as SharepointResult; 
      if (lists.value && lists.value.length == 1) {
        return lists.value[0];
      }
      return null;
    } catch (e) {
      if(e.status == 401) {
        // await this.teams.refreshToken(true); 
      }
      return null;
    }
  }

  async getOneItemById(id: number, list: string, conditions: string = ''): Promise<any> {
    try {
      let endpoint = this.licensing.getSharepointUri() + list + `/items(${id})`;
      if (conditions) endpoint += '?' + conditions;
      return await this.http.get(endpoint).toPromise(); 
    } catch (e) {
      if(e.status == 401) {
        // await this.teams.refreshToken(true); 
      }
      return null;
    }
    return null;
  }

  async countItems(list: string, conditions: string = ''): Promise<number> {
    try {
      let endpoint = this.licensing.getSharepointUri() + list + '/ItemCount';
      if (conditions) endpoint += '?' + conditions;
      let lists = await this.http.get(endpoint).toPromise() as SharepointResult; 
      if (lists.value) {
        return lists.value;
      }
      return 0;
    } catch (e) {
      if(e.status == 401) {
        // await this.teams.refreshToken(true); 
      }
      return 0;
    }
  }

  async createItem(list: string, data: any): Promise<any> {
    try {
      return await this.http.post(
        this.licensing.getSharepointUri() + list + "/items", 
        data
      ).toPromise();
    } catch (e) {
      if(e.status == 401) {
        // await this.teams.refreshToken(true);
      }
      return {};
    }
  }

  async updateItem(id: number, list: string, data: any): Promise<boolean> {
    try {
      await this.http.post(
        this.licensing.getSharepointUri() + list + `/items(${id})`, 
        data,
        {
          headers: new HttpHeaders({
            'If-Match': '*',
            'X-HTTP-Method': "MERGE"
          })
        }
      ).toPromise();
    } catch (e) {
      if(e.status == 401) {
        // await this.teams.refreshToken(true);
      }
      return false;
    }
    return true;
  }


  /** FILES */
  
  getBaseFilesFolder(): string {
    return FILES_FOLDER;
  }

  async createFolder(newFolderUrl: string) {
    try {
      return await this.http.post(
        this.licensing.getSharepointUri() + "/folders", 
        {
          ServerRelativeUrl: FILES_FOLDER + newFolderUrl
        }
      ).toPromise();
    } catch (e) {
      if(e.status == 401) {
        // await this.teams.refreshToken(true);
        console.log('The folder cannot be created');
      }
      return false;
    }
  }

  async readFile(fileUri: string): Promise<any> {
    try {
      return this.http.get(
        this.licensing.getSharepointUri() + `GetFileByServerRelativeUrl('${fileUri}')/$value`, 
        { responseType: 'arraybuffer' }
      ).toPromise();
    } catch (e) {
      if(e.status == 401) {
        // await this.teams.refreshToken(true); 
      }
      return [];
    }
  }

  async deleteFile(fileUri: string): Promise<boolean> {
    try {
      await this.http.post(
        this.licensing.getSharepointUri() + `GetFileByServerRelativeUrl('${fileUri}')`, 
        null,
        {
          headers: new HttpHeaders({
            'If-Match': '*',
            'X-HTTP-Method': "DELETE"
          }),
        }
      ).toPromise();
    } catch (e) {
      if(e.status == 401) {
        // await this.teams.refreshToken(true); 
      }
      return false;
    }
    return true;
  }

  async uploadFile(fileData: string, folder: string, fileName: string, metadata?: any): Promise<any> {
    let uploaded: any = await this.uploadFileQuery(fileData, folder, fileName);

    if (metadata && uploaded.ListItemAllFields?.ID/* && uploaded.ServerRelativeUrl*/) {

      // GetFileByServerRelativeUrl('/Folder Name/{file_name}')/CheckOut()
      // GetFileByServerRelativeUrl('/Folder Name/{file_name}')/CheckIn(comment='Comment',checkintype=0)

      await this.updateItem(uploaded.ListItemAllFields.ID, `lists/getbytitle('${FILES_FOLDER}')`, metadata);
    }
    return uploaded;
  }

  /** Impossible to expand ListItemAllFields/Author in one query using Sharepoint REST API */
  async readFolderFiles(folder: string, expandProperties = false): Promise<NPPFile[]> {
    let files: NPPFile[] = []
    const result = await this.query(
      `GetFolderByServerRelativeUrl('${this.getBaseFilesFolder()}/${folder}')/Files`, 
      '$expand=ListItemAllFields',
    ).toPromise();

    if (result.value) {
      files = result.value;
    }
    if (expandProperties && files.length > 0) {
      for (let i = 0; i < files.length; i++) {
        const fileItems = files[i].ListItemAllFields;
        if (fileItems) {
          Object.assign(fileItems, this.getFileInfo(fileItems.ID));
        }
      }
    }
    return files;
  }

  async getSubfolders(folder: string): Promise<any> {
    let subfolders: any[] = [];
    const result = await this.query(
      `GetFolderByServerRelativeUrl('${this.getBaseFilesFolder()}/${folder}')/folders`, 
      '$expand=ListItemAllFields',
    ).toPromise();
    if (result.value) {
      subfolders = result.value;
    }
    return subfolders;
  }

  async getFileInfo(fileId: number): Promise<NPPFile> {
    return await this.query(
      `lists/getbytitle('${FILES_FOLDER}')` + `/items(${fileId})`, 
      '$select=*,Author/Id,Author/FirstName,Author/LastName,StageName/Id,StageName/Title,TargetUser/FirstName,TargetUser/LastName, \
        Country/Title, ModelScenario/Title \
        &$expand=StageName,Author,TargetUser,Country,ModelScenario',
      'all'
    ).toPromise();
  }

  private async uploadFileQuery(fileData: string, folder: string, filename: string) {
    try {
      let url = `GetFolderByServerRelativeUrl('${folder}')/Files/add(url='${filename}',overwrite=true)?$expand=ListItemAllFields`;
      return await this.http.post(
        this.licensing.getSharepointUri() + url, 
        fileData,
        {
          headers: { 'Content-Type': 'blob' }
        }
      ).toPromise();
    } catch (e) {
      if(e.status == 401) {
        // await this.teams.refreshToken(true);
      }
      return {};
    }
  }

  /** OPPORTUNITIES */

  async testEndpoint() {
    /*
    // "e78fab3e-248e-442c-b2e1-15f309e9d276"
    let r = await this.http.post('https://login.microsoftonline.com/e78fab3e-248e-442c-b2e1-15f309e9d276/oauth2/token', {
      client_id: "b431132e-d7ea-4206-a0a9-5403adf64155/.default",
      // client_secret: qDk2.~ZH0_aVijr.~2K2R4II-1~keYjwI2
      grant_type: 'client_credentials',
      resource: 'https://janddconsulting.onmicrosoft.com/NPPProvisioning-API',
      // scope: b431132e-d7ea-4206-a0a9-5403adf64155./default,
    }).toPromise();
    console.log('r', r);
    */
    
    // "eyJ0eXAiOiJKV1QiLCJub25jZSI6Im81RUREbWJaTVhuVXZ4UmU2Q05hYU0tQ2VOZkdEQXFQWGxISEVWRU5qZ1kiLCJhbGciOiJSUzI1NiIsIng1dCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyIsImtpZCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9lNzhmYWIzZS0yNDhlLTQ0MmMtYjJlMS0xNWYzMDllOWQyNzYvIiwiaWF0IjoxNjMwOTQzNjczLCJuYmYiOjE2MzA5NDM2NzMsImV4cCI6MTYzMDk0NzU3MywiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFTUUEyLzhUQUFBQW1pL1B1M2VEcDNITlVONWF0SmFxOUZnS1Iyemp6elVEZEtxVm1qcnh6a2c9IiwiYW1yIjpbInB3ZCJdLCJhcHBfZGlzcGxheW5hbWUiOiJOUFAgRGVtbyIsImFwcGlkIjoiMTc1MzRjYTItZjRmOC00M2MwLTg2MTItNzJiZGQyOWE5ZWU4IiwiYXBwaWRhY3IiOiIwIiwiZmFtaWx5X25hbWUiOiJNYcOxw6kiLCJnaXZlbl9uYW1lIjoiQWxiZXJ0IiwiaWR0eXAiOiJ1c2VyIiwiaXBhZGRyIjoiMzcuMTQuMTA4Ljc3IiwibmFtZSI6IkFsYmVydCBNYcOxw6kiLCJvaWQiOiI5MzJiNmNkMC03ODY4LTQ4MWEtOTM3NC1hOTQ2Njg4M2M0ZjMiLCJwbGF0ZiI6IjgiLCJwdWlkIjoiMTAwMzIwMDE2QjFDMDBBMCIsInJoIjoiMC5BWUVBUHF1UDU0NGtMRVN5NFJYekNlblNkcUpNVXhmNDlNQkRoaEp5dmRLYW51aUJBQzQuIiwic2NwIjoiQWxsU2l0ZXMuRnVsbENvbnRyb2wgQWxsU2l0ZXMuTWFuYWdlIE15RmlsZXMuUmVhZCBNeUZpbGVzLldyaXRlIG9wZW5pZCBwcm9maWxlIFNpdGVzLlJlYWQuQWxsIFNpdGVzLlNlYXJjaC5BbGwgVXNlci5SZWFkIFVzZXIuUmVhZC5BbGwgZW1haWwiLCJzdWIiOiJfdG5keWdwWnRqTzV1ckhTbThyS01TQTF0VW9NS3JjRDB4aEZBczZ0TWg4IiwidGVuYW50X3JlZ2lvbl9zY29wZSI6IkVVIiwidGlkIjoiZTc4ZmFiM2UtMjQ4ZS00NDJjLWIyZTEtMTVmMzA5ZTlkMjc2IiwidW5pcXVlX25hbWUiOiJhbGJlcnRAYmV0YXNvZnR3YXJlc2wub25taWNyb3NvZnQuY29tIiwidXBuIjoiYWxiZXJ0QGJldGFzb2Z0d2FyZXNsLm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6IlQ3ejV5dWx1bjBpcjB5RE5ldGU4QUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdLCJ4bXNfc3QiOnsic3ViIjoiQTBCeHN4allIWXBueTJ1TzFNYWx1WF9iVDh4RjJYMXdxVHhrVmotcFNUdyJ9LCJ4bXNfdGNkdCI6MTYyMzIyODYyM30.EFMaV_HCwk6xj1Wegx0ISJs_oONuHErAwazGOV3lmbSMLqUtgF7_fYsHVEtB1giqrlwmpDLcuZTvSpcBI2UiyzFUNsn2n5OKkmJIGPX3my_ZjkkT_7uGa2YWKknBCuLW6g-nZjsJ3QkjxzNRMBQ0erZIRhrPUzSlbT3L45QfB2pY7JyzcT4K_TXVoY1tHOsrsxT5JTgV-t1O9Z8qL22qO873bmh-hvB3fAZh80gF3F8EkDcx5XB2gsWXwLB536Hr24PFWuPGoWBe-CjPFSfdoG_88hT8bbmcD__bZUntBnylqQWcxWc2Sc-N-ezgG7LZ7Cv1xgpkEkBA4c3fOv-GHQ"
    let result = await this.http.get(
      `https://nppprovisioning20210831.azurewebsites.net/api/NewOpportunity?StageID=18&OppID=2&siteUrl=https://betasoftwaresl.sharepoint.com/sites/JDNPPApp`
    ).toPromise();
  }

  async getOpportunities(): Promise<Opportunity[]> {
    return await this.getAllItems(
      OPPORTUNITIES_LIST, 
      "$select=*,OpportunityType/Title,Indication/TherapyArea,Indication/Title,OpportunityOwner/FirstName,OpportunityOwner/LastName,OpportunityOwner/ID,OpportunityOwner/EMail&$expand=OpportunityType,Indication,OpportunityOwner"
      );
  }

  async createOpportunity(op: OpportunityInput, st: StageInput): Promise<Opportunity> {
    let opportunity = await this.createItem(OPPORTUNITIES_LIST, { OpportunityStatus: "Processing", ...op });
    let stageType = await this.getStageType(op.OpportunityTypeId);
    let masterStage = await this.getOneItem(MASTER_STAGES_LIST, `$select=ID&$filter=(StageType eq '${stageType}') and (StageNumber eq 1)`);
    let stage = await this.createItem(OPPORTUNITY_STAGES_LIST, { ...st, OpportunityNameId: opportunity.ID, StageNameId: masterStage.ID });

    this.initializeOpportunity(opportunity, stage);

    // set active (TODO when finished)
    this.updateItem(
      opportunity.ID,
      OPPORTUNITIES_LIST,
      {
        OpportunityStatus: "Active"
      }
    );
    return opportunity;
  }

  async updateOpportunity(oppId: number, oppData: OpportunityInput): Promise<boolean> {
    return await this.updateItem(oppId, OPPORTUNITIES_LIST, oppData);
  }

  async setOpportunityStatus(opportunityId: number, status: string) {
    const allowedStatus = ["Processing", "Archive", "Active", "Approved"];
    if (allowedStatus.includes(status)) {
      return this.updateItem(opportunityId, OPPORTUNITIES_LIST, {
        OpportunityStatus: status
      });
    }
    return false;
  }

  private async initializeOpportunity(opportunity: Opportunity, stage: Stage) {
    const groups = await this.createOpportunityGroups(opportunity.OpportunityOwnerId, opportunity.ID, stage.StageNameId);
    console.log('groups created');

    let permissions;
    // add groups to lists
    permissions = (await this.getGroupPermissions()).filter(el => el.ListFilter === 'List');
    this.setPermissions(permissions, groups);
    /*
    for (const gp of permissions) {
      const group = groups.find(gr => gr.type === gp.Title); // get created group involved on the permission
      if (group) {
        await this.addRolePermissionToList(`lists/getbytitle('${gp.ListName}')`, group.data.Id);
      }
    }
    */

    // add groups to the Opportunity
    permissions = await this.getGroupPermissions(OPPORTUNITES_LIST_NAME);
    console.log('permissions', permissions);
    this.setPermissions(permissions, groups, opportunity.ID);

    /*
    for (const gp of permissions) {
      const group = groups.find(gr => gr.type === gp.Title); // get created group involved on the permission
      if (group) {
        if (gp.ListFilter === 'Item') 
          await this.addRolePermissionToList(OPPORTUNITIES_LIST, group.data.Id, opportunity.ID);
        else 
          await this.addRolePermissionToList(OPPORTUNITIES_LIST, group.data.Id);
      }
    }
    */

    // add groups to the Stage
    permissions = await this.getGroupPermissions(OPPORTUNITY_STAGES_LIST_NAME);
    console.log('permissions', permissions);
    this.setPermissions(permissions, groups, stage.ID);

    /*
    for (const gp of permissions) {
      const group = groups.find(gr => gr.type === gp.Title); // get created group involved on the permission
      if (group) {
        if (gp.ListFilter === 'Item') 
          await this.addRolePermissionToList(OPPORTUNITY_STAGES_LIST, group.data.Id, stage.ID);
        else 
          await this.addRolePermissionToList(OPPORTUNITY_STAGES_LIST, group.data.Id);
      }
    }
    */

    // Actions
    const stageActions = await this.createStageActions(opportunity, stage);

    // add groups to the Actions
    permissions = await this.getGroupPermissions(OPPORTUNITY_ACTIONS_LIST_NAME);
    for (const action of stageActions) {
      this.setPermissions(permissions, groups, action.Id);
    }

    /*
    for (const action of stageActions) {
      for (const gp of permissions) {
        const group = groups.find(gr => gr.type === gp.Title); // get created group involved on the permission
        if (group) {
          if (gp.ListFilter === 'Item') 
            await this.addRolePermissionToList(OPPORTUNITY_ACTIONS_LIST, group.data.Id, action.Id);
          else 
            await this.addRolePermissionToList(OPPORTUNITY_ACTIONS_LIST, group.data.Id);
        }
      }
    } 
    */

    // Folders
    this.createStageFolders(stage);
    return true;
  }

  private async createStageActions(opportunity: Opportunity, stage: Stage) {
    const masterActions: MasterAction[] = await this.getAllItems(
      MASTER_ACTION_LIST, 
      `$filter=StageNameId eq ${stage.StageNameId} and OpportunityTypeId eq ${opportunity.OpportunityTypeId}&$orderby=ActionNumber asc`
    );

    let actions: Action[] = [];
    for (const ma of masterActions) {
      const a = await this.createAction(ma, opportunity.ID);
      if (a.Id) actions.push(a);
    }
    return actions;
  }

  private async createAction(ma: MasterAction, oppId: number): Promise<Action> {
    let dueDate = new Date();
    dueDate.setDate(dueDate.getDate() + ma.DueDays);
    return await this.createItem(
      OPPORTUNITY_ACTIONS_LIST,
      {
        Title: ma.Title,
        StageNameId: ma.StageNameId,
        OpportunityNameId: oppId,
        ActionNameId: ma.Id,
        ActionDueDate: dueDate
      }
    );
  }

  private async createStageFolders(stage: Stage) {
    const masterFolders = await this.getAllItems(
      MASTER_FOLDER_LIST,
      `$filter=StageNameId eq ${stage.StageNameId}`
    );

    await this.createFolder(`/${stage.OpportunityNameId}`);
    await this.createFolder(`/${stage.OpportunityNameId}/${stage.StageNameId}`);

    masterFolders.forEach(async (mf) => {
      await this.createFolder(`/${stage.OpportunityNameId}/${stage.StageNameId}/${mf.Id}`);
    });
  }

  async createOpportunityGroups(ownerId: number, oppId: number, masterStageId: number): Promise<SPGroupListItem[]> {
    let group;
    let groups: SPGroupListItem[] = [];
    const owner = await this.getUserInfo(ownerId);
    if (!owner.LoginName) return [];

    // Opportunity Owner (OO)
    group = await this.createGroup(`OO-${oppId}`);
    if (group) {
      groups.push({ type: 'OO', data: group });
      await this.addUserToGroup(owner.LoginName, group.Id);
    }

    // Opportunity Users (OU)
    group = await this.createGroup(`OU-${oppId}`);
    if (group) {
      groups.push({ type: 'OU', data: group });
      await this.addUserToGroup(owner.LoginName, group.Id);
    }

    // Stage Owners (SO)
    group = await this.createGroup(`SO-${oppId}-${masterStageId}`);
    if (group) {
      groups.push({ type: 'SO', data: group });
      await this.addUserToGroup(owner.LoginName, group.Id);
    }
    return groups;
  }



  /** PERMISSIONS */
  async createGroup(name: string, description: string = ''): Promise<SPGroup | null> {
    try {
      return await this.http.post(
        this.licensing.getSharepointUri() + 'sitegroups',
        {
          Title: name,
          Description: description,
          OnlyAllowMembersViewMembership: false
        }
      ).toPromise() as SPGroup;
    } catch (e) {
      if (e.status == 401) {
        // await this.teams.refreshToken(true); 
      }
      return null;
    }
  }

  async addUserToGroup(loginName: string, groupId: number): Promise<boolean> {
    try {
      await this.http.post(
        this.licensing.getSharepointUri() + `sitegroups(${groupId})/users`,
        { LoginName: loginName }
      ).toPromise();
      return true;
    } catch (e) {
      if(e.status == 401) {
        // await this.teams.refreshToken(true); 
      }
      return false;
    }
  }

  async readGroups(): Promise<SPGroup[]> {
    return await this.query('sitegroups').toPromise();
  }

  async getGroupPermissions(list: string = ''): Promise<GroupPermission[]> {
    if (this.masterGroupTypes.length < 1) {
      this.masterGroupTypes = await this.getAllItems(MASTER_GROUP_TYPES_LIST);
    }
    if (list) {
      return this.masterGroupTypes.filter(el => el.ListName === list);
    }
    return this.masterGroupTypes;
  }

  /* set permissions related to working groups a list or item */
  private async setPermissions(permissions: GroupPermission[], workingGroups: SPGroupListItem[], id: number | null = null) {
    for (const gp of permissions) {
      const group = workingGroups.find(gr => gr.type === gp.Title); // get created group involved on the permission
      if (group) {
        if (gp.ListFilter === 'List') 
          await this.addRolePermissionToList(`lists/getbytitle('${gp.ListName}')`, group.data.Id);
        else if (id) 
          await this.addRolePermissionToList(`lists/getbytitle('${gp.ListName}')`, group.data.Id, id);
      }
    }
  }

  private async addRolePermissionToList(list: string, groupId: number, id: number = 0): Promise<boolean> {
    const baseUrl = this.licensing.getSharepointUri() + list + (id === 0 ? '' : `/items(${id})`);
    try {
      await this.http.post(
        baseUrl + `/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)`,
        null).toPromise();
      await this.http.post(
        baseUrl + `/roleassignments/addroleassignment(principalid=${groupId},roledefid=1073741826)`,
        null).toPromise();
      return true;
    } catch (e) {
      if (e.status == 401) {
        // await this.teams.refreshToken(true); 
      }
      return false;
    }
  }

  async testAddGroup() {
    const group = await this.query("sitegroups/getbyname('Beta Test Group')/id").toPromise();
    console.log('group', group);

    await this.http.post(
      this.licensing.getSharepointUri() + `GetFolderByServerRelativeUrl('Current Opportunity Library/2/3')/ListItemAllFields/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)`,
      null).toPromise();
    const result = await this.http.post(
      this.licensing.getSharepointUri() + `GetFolderByServerRelativeUrl('Current Opportunity Library/2/3')/ListItemAllFields/roleassignments/addroleassignment(principalid=${group.value},roledefid=1073741826)`,
      null).toPromise();
  }

  async testAddGroupToOpportunity() {
    const group = await this.query("sitegroups/getbyname('Beta Test Group')/id").toPromise();
    console.log('group', group);

    await this.http.post(
      this.licensing.getSharepointUri()  + OPPORTUNITIES_LIST + `/items(27)/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)`,
      null).toPromise();
    const result = await this.http.post(
      this.licensing.getSharepointUri() + OPPORTUNITIES_LIST + `/items(27)/roleassignments/addroleassignment(principalid=${group.value},roledefid=1073741826)`,
      null).toPromise();

  }


  async getIndications(therapy: string = 'all'): Promise<Indication[]> {
    let cache = this.masterIndications.find(i => i.therapy == therapy);
    if (cache) {
      return cache.indications;
    }
    let max = await this.countItems(MASTER_THERAPY_AREAS_LIST);
    let cond = "$skiptoken=Paged=TRUE&$top="+max;
    if (therapy !== 'all') {
      cond += `&$filter=TherapyArea eq '${therapy}'`;
    }
    let results = await this.getAllItems(MASTER_THERAPY_AREAS_LIST, cond);
    this.masterIndications.push({
      therapy: therapy,
      indications: results
    });
    return results;
  }

  async getIndicationsList(therapy?: string): Promise<SelectInputList[]> {
    let indications = await this.getIndications(therapy);

    if (therapy) {
      return indications.map(el => { return {value: el.ID, label: el.Title}})
    }
    return indications.map(el => { return {value: el.ID, label: el.Title, group: el.TherapyArea}})
  }

  async getTherapiesList(): Promise<SelectInputList[]> {
    if (this.masterTherapiesList.length < 1) {
      let count = await this.countItems(MASTER_THERAPY_AREAS_LIST);
      let indications: Indication[] = await this.getAllItems(MASTER_THERAPY_AREAS_LIST, "$orderby=TherapyArea asc&$skiptoken=Paged=TRUE&$top="+count);

      return indications
        .map(v => v.TherapyArea)
        .filter((value, index, self) => self.indexOf(value) === index)
        .map(v => { return { label: v, value: v }});
    }
    return this.masterTherapiesList;
  }

  async getStages(opportunityId: number): Promise<Stage[]> {
    return await this.getAllItems(OPPORTUNITY_STAGES_LIST, "$filter=OpportunityNameId eq "+opportunityId);
  }

  async getActions(opportunityId: number, stageId?: number): Promise<Action[]> {
    let filterConditions = `(OpportunityNameId eq ${opportunityId})`;
    if (stageId) filterConditions += ` and (StageNameId eq ${stageId})`;
    return await this.getAllItems(OPPORTUNITY_ACTIONS_LIST, `$select=*,TargetUser/ID,TargetUser/FirstName,TargetUser/LastName&$filter=${filterConditions}&$orderby=StageNameId%20asc&$expand=TargetUser`);
  }

  async completeAction(actionId: number, userId: number): Promise<boolean> {
    const data = {
      TargetUserId: userId,
      Timestamp: new Date(),
      Complete: true
    };
    return await this.updateItem(actionId, OPPORTUNITY_ACTIONS_LIST, data);
  }

  async uncompleteAction(actionId: number): Promise<boolean> {
    const data = {
      TargetUserId: null,
      Timestamp: null,
      Complete: false
    };
    return await this.updateItem(actionId, OPPORTUNITY_ACTIONS_LIST, data);
  }

  async setActionDueDate(actionId: number, newDate: string) {
    return await this.updateItem(actionId, OPPORTUNITY_ACTIONS_LIST, { ActionDueDate: newDate });
  }

  async updateStage(stageId: number, data: any) {
    return await this.updateItem(stageId, OPPORTUNITY_STAGES_LIST, data);
  }

  async createStage(data: StageInput): Promise<boolean> {
    return await this.createItem(OPPORTUNITY_STAGES_LIST, data);
  }

  async getLists() {
   /* try {
      let lists = await this.query('lists').toPromise();
      return lists;
    } catch (e) {
      if(e.status == 401) {
        this.teams.loginAgain();
      }
      return [];
    }*/

    /*
    this.http.get('https://graph.microsoft.com/v1.0/me').subscribe(
      r => {
        console.log('r grapg', r);
      }
    );
    */
    this.http.get('https://betasoftwaresl.sharepoint.com/sites/JDNPPApp/_api/web/lists').subscribe(
      r => {
        console.log('r sharepoint', r);
      }
    );
    /*
    this.http.get(this.licensing.siteUrl + "lists/getbytitle('Master Opportunity Type List')/items")
      .subscribe(profile => {
        console.log('response', profile);
      });
    */
  }

  async getOpportunityTypes(): Promise<OpportunityType[]> {
    if (this.masterOpportunitiesTypes.length < 1) {
      this.masterOpportunitiesTypes = await this.getAllItems(MASTER_OPPORTUNITY_TYPES_LIST);
    }
    return this.masterOpportunitiesTypes;
  }

  async getOpportunityTypesList(): Promise<SelectInputList[]> {
    return (await this.getOpportunityTypes()).map(t => {return {value: t.ID, label: t.Title}});
  }

  async getOpportunityFields() {
    return [
      { value: 'title', label: 'Opportunity Name' },
      { value: 'projectStart', label: 'Project Start Date' },
      { value: 'projectEnd', label: 'Project End Date' },
      { value: 'opportunityType', label: 'Project Type' },
    ];
  }

  async getOpportunity(id: number): Promise<Opportunity> {
    return await this.getOneItem(OPPORTUNITIES_LIST, "$filter=Id eq "+id+"&$select=*,OpportunityType/Title,Indication/TherapyArea,Indication/Title,Author/FirstName,Author/LastName,Author/ID,Author/EMail,OpportunityOwner/ID,OpportunityOwner/FirstName,OpportunityOwner/EMail,OpportunityOwner/LastName&$expand=OpportunityType,Indication,Author,OpportunityOwner");
  }

  // async getFiles(id: number) {
  //   return this.files.filter(f => f.parentId == id);
  // }



  async getStageType(OpportunityTypeId: number): Promise<string> {
    let result: OpportunityType | undefined;
    if (this.masterOpportunitiesTypes.length > 0) {
      result = this.masterOpportunitiesTypes.find(ot => ot.ID === OpportunityTypeId);
    } else {
      result = await this.getOneItem(MASTER_OPPORTUNITY_TYPES_LIST, "$filter=Id eq "+OpportunityTypeId+"&$select=StageType");
    }
    if (result == null) {
      return '';
    }
    return result.StageType;
  }

  async getNextStage(stageId: number): Promise<Stage | null> {
    let current = await this.getOneItemById(stageId, MASTER_STAGES_LIST);
    console.log('current', current);
    return await this.getOneItem(MASTER_STAGES_LIST, `$filter=StageNumber eq ${current.StageNumber + 1} and StageType eq '${current.StageType}'`);
  }

  async getStageFolders(opportunityId: number, masterStageId: number, allFolders = false): Promise<NPPFolder[]> {
    let masterFolders = [];
    let cache = this.masterFolders.find(f => f.stage == masterStageId);
    if (cache) {
      masterFolders = cache.folders;
    } else {
      masterFolders = await this.getAllItems(MASTER_FOLDER_LIST, "$filter=StageNameId eq "+masterStageId);
      for (let index = 0; index < masterFolders.length; index++) {
        masterFolders[index].containsModels = masterFolders[index].Title === 'Forecast Models';
      }
      this.masterFolders.push({
        stage: masterStageId,
        folders: masterFolders
      });
    }
    
    if (allFolders) {
      return masterFolders;
    } else {
      // only folders user can access
      const allowedFolders = await this.getSubfolders(`/${opportunityId}/${masterStageId}`);
      return masterFolders.filter(f => allowedFolders.some((af: any)=> +af.Name === f.ID));
    }
  }

  async getCountriesList(): Promise<SelectInputList[]> {
    if (this.masterCountriesList.length < 1) {
      this.masterCountriesList = (await this.getAllItems(COUNTRIES_LIST, "$orderby=Title asc")).map(t => {return {value: t.ID, label: t.Title}});
    }
    return this.masterCountriesList;
  }

  async getScenariosList(): Promise<SelectInputList[]> {
    if (this.masterScenariosList.length < 1) {
      this.masterScenariosList = (await this.getAllItems(MASTER_SCENARIOS_LIST)).map(t => {return {value: t.ID, label: t.Title}});
    }
    return this.masterScenariosList;
  }

  /** todel */
  async getTest() {
    console.log('files', await this.getAllItems(`lists/getbytitle('Current Opportunity Library')`));
  }

  
  /** USERS */

  async getUserProfilePic(userId: number): Promise<string> {
    let queryObj = await this.getOneItemById(userId, USER_INFO_LIST, `$select=Id,Picture`);
    return queryObj.Picture?.Url;
  }

  async getUsersFromGroup(groupName: string): Promise<User[]> {
    // return this.getAllItems("sitegropus/getbyname('Beta Test Group')", 
    let users = await this.query(`sitegroups/getbyname('${groupName}')/users`).toPromise();
    if (users && users.value.length > 0) {
      return users.value;
    }      
    return [];
  }

  async getCurrentUserInfo(): Promise<User> {
    let account = localStorage.getItem('sharepointAccount');
    if(account) {
      return JSON.parse(account);
    } else {
      let account = await this.query('currentuser', '?$select=Title,Email,Id,FirstName,LastName').toPromise();
      console.log('account sharepoint', account);
      account['ID'] = account.Id; // set for User interface
      localStorage.setItem('sharepointAccount', JSON.stringify(account));
      return account;
    }
  }

  async getUserInfo(userId: number): Promise<User> {
    return await this.query(`siteusers/getbyid('${userId}')`).toPromise();
  }

  removeCurrentUserInfo() {
    localStorage.removeItem('sharepointAccount');
  }

  /** NOTIFICATIONS */
  async createNotification(userId: number, text: string): Promise<boolean> {
    return this.createItem(NOTIFICATIONS_LIST, {
      Title: text,
      TargetUserId: userId
    });
  }

  /** LISTS */
  searchByTermInputList(query: string, field: string, term: string, matchCase = false): Observable<SelectInputList[]> {
    return this.query(query, '', 'all', { term, field, matchCase })
      .pipe(
        map((res: any) => {
          return res.value.map(
            (el: any) => { return { value: el.Id, label: el.Title } as SelectInputList }
          );
        })
      );
  }

  async getUsersList(usersId: number[]): Promise<SelectInputList[]> {
    const conditions = usersId.map(e => { return '(Id eq ' + e + ')' }).join(' or ');
    const users = await this.query('siteusers', '$filter='+conditions).toPromise();
    if (users.value) {
      return users.value.map((u: User) => { return { label: u.Title, value: u.Id }});
    }
    return [];
  }

}
