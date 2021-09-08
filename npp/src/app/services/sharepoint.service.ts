import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Observable, of } from 'rxjs';
import { catchError, filter } from 'rxjs/operators';
import { ErrorService } from './error.service';
import { LicensingService } from './licensing.service';
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
  IsSiteAdmin?: boolean;
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
  DepartmentID?: number;
  containsModels?: boolean;
}

export interface SystemFolder {
  Name: string;
  ServerRelativeUrl: string;
  ItemCount: number;
  DepartmentID?: number;
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
  SPRoleDefinitions: {
    name: string;
    id: number;
  }[] = [];

  constructor(private http: HttpClient, private error: ErrorService, private licensing: LicensingService) { }

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

  async createFolder(newFolderUrl: string): Promise<SystemFolder | null> {
    try {
      return await this.http.post(
        this.licensing.getSharepointUri() + "folders", 
        {
          ServerRelativeUrl: FILES_FOLDER + newFolderUrl
        }
      ).toPromise() as SystemFolder;
    } catch (e) {
      if(e.status == 401) {
        // await this.teams.refreshToken(true);
        console.log('The folder cannot be created');
      }
      return null;
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

  async getOpportunities(): Promise<Opportunity[]> {
    return await this.getAllItems(
      OPPORTUNITIES_LIST, 
      "$select=*,OpportunityType/Title,Indication/TherapyArea,Indication/Title,OpportunityOwner/FirstName,OpportunityOwner/LastName,OpportunityOwner/ID,OpportunityOwner/EMail&$expand=OpportunityType,Indication,OpportunityOwner"
      );
  }

  async createOpportunity(op: OpportunityInput, st: StageInput): Promise<{ opportunity: Opportunity, stage: Stage } | false> {
    let opportunity = await this.createItem(OPPORTUNITIES_LIST, { OpportunityStatus: "Processing", ...op });
    let stageType = await this.getStageType(op.OpportunityTypeId);
    let masterStage = await this.getOneItem(MASTER_STAGES_LIST, `$select=ID&$filter=(StageType eq '${stageType}') and (StageNumber eq 1)`);
    let stage = await this.createItem(OPPORTUNITY_STAGES_LIST, { ...st, OpportunityNameId: opportunity.ID, StageNameId: masterStage.ID });

    return { opportunity, stage };
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

  async initializeOpportunity(opportunity: Opportunity, stage: Stage): Promise<boolean> {
    const groups = await this.createOpportunityGroups(opportunity.OpportunityOwnerId, opportunity.ID, stage.StageNameId);

    let permissions;
    // add groups to lists
    permissions = (await this.getGroupPermissions()).filter(el => el.ListFilter === 'List');
    await this.setPermissions(permissions, groups);
  
    // add groups to the Opportunity
    permissions = await this.getGroupPermissions(OPPORTUNITES_LIST_NAME);
    await this.setPermissions(permissions, groups, opportunity.ID);

    // add groups to the Stage
    permissions = await this.getGroupPermissions(OPPORTUNITY_STAGES_LIST_NAME);
    console.log('permissions', permissions);
    await this.setPermissions(permissions, groups, stage.ID);

    // add stage users to group OU
    const OUGroup = groups.find(g => g.type === 'OU');
    if (OUGroup) {
      for (const userId of stage.StageUsersId) {
        const user = await this.getUserInfo(userId);
        if (user.LoginName) await this.addUserToGroup(user.LoginName, OUGroup.data.Id);
      }
    }

    // Actions
    const stageActions = await this.createStageActions(opportunity, stage);

    // add groups to the Actions
    permissions = await this.getGroupPermissions(OPPORTUNITY_ACTIONS_LIST_NAME);
    for (const action of stageActions) {
      await this.setPermissions(permissions, groups, action.Id);
    }

    // Folders
    const folders = await this.createStageFolders(stage);

    // add groups to folders
    permissions = await this.getGroupPermissions(FILES_FOLDER);
    for (const f of folders) {
      if (f.DepartmentID) {
        let folderGroups = [...groups]; // copy default groups
        const DUGroup = await this.createGroup(`DU-${opportunity.ID}-${f.DepartmentID}`, 'Department ID ' + f.DepartmentID);
        if (DUGroup) folderGroups.push( { type: 'DU', data: DUGroup} );
        await this.setPermissions(permissions, folderGroups, f.ServerRelativeUrl);
      }
    }
    return true;
  }

  private async createStageActions(opportunity: Opportunity, stage: Stage): Promise<Action[]> {
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

  private async createStageFolders(stage: Stage): Promise<SystemFolder[]> {
    const masterFolders = await this.getStageFolders(stage.StageNameId);

    await this.createFolder(`/${stage.OpportunityNameId}`);
    await this.createFolder(`/${stage.OpportunityNameId}/${stage.StageNameId}`);

    let folders: SystemFolder[] = [];

    for (const mf of masterFolders) {
      const folder = await this.createFolder(`/${stage.OpportunityNameId}/${stage.StageNameId}/${mf.ID}`);
      if (folder) {
        folder.DepartmentID = mf.DepartmentID;
        folders.push(folder);
      }
    }
    return folders;
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

  async getGroupId(name: string): Promise<number | null> {
    try {
      const result = await this.query(`sitegroups/getbyname('${name}')/id`).toPromise();
      return result.value;
    } catch (e) {
      return null;
    }
  }

  async getRoleDefinitionId(name: string): Promise<number | null> {
    const cache = this.SPRoleDefinitions.find(g => g.name === name);
    if (cache) {
      return cache.id;
    } else {
      try {
        const result = await this.query(`roledefinitions/getbyname('${name}')/id`).toPromise();
        this.SPRoleDefinitions.push({name, id: result.value }); // add for local caching
        return result.value;
      }
      catch (e) {
        return null;
      }
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
  private async setPermissions(permissions: GroupPermission[], workingGroups: SPGroupListItem[], itemOrFolder: number | string | null = null) {
    for (const gp of permissions) {
      const group = workingGroups.find(gr => gr.type === gp.Title); // get created group involved on the permission
      if (group) {
        if (gp.ListName === FILES_FOLDER && typeof itemOrFolder == 'string') {
          await this.addRolePermissionToFolder(itemOrFolder, group.data.Id, gp.Permission);
        } else {
          if (gp.ListFilter === 'List')
            await this.addRolePermissionToList(`lists/getbytitle('${gp.ListName}')`, group.data.Id, gp.Permission);
          else if (typeof itemOrFolder == 'number')
            await this.addRolePermissionToList(`lists/getbytitle('${gp.ListName}')`, group.data.Id, gp.Permission, itemOrFolder);
        }
      }
    }
  }

  private async addRolePermissionToList(list: string, groupId: number, roleName: string, id: number = 0): Promise<boolean> {
    const baseUrl = this.licensing.getSharepointUri() + list + (id === 0 ? '' : `/items(${id})`);
    return await this.setRolePermission(baseUrl, groupId, roleName);
  }

  private async addRolePermissionToFolder(folderUrl: string, groupId: number, roleName: string): Promise<boolean> {
    const baseUrl = this.licensing.getSharepointUri() + `GetFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields`;
    return await this.setRolePermission(baseUrl, groupId, roleName);
  }

  private async setRolePermission(baseUrl: string, groupId: number, roleName: string) {
    // const roleId = 1073741826; // READ
    const roleId = await this.getRoleDefinitionId(roleName);
    try {
      await this.http.post(
        baseUrl + `/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)`,
        null).toPromise();
      await this.http.post(
        baseUrl + `/roleassignments/addroleassignment(principalid=${groupId},roledefid=${roleId})`,
        null).toPromise();
      return true;
    } catch (e) {
      if (e.status == 401) {
        // await this.teams.refreshToken(true); 
      }
      return false;
    }
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

  /** get stage folders. If opportunityId, only the folders with permission. Otherwise, all master folders of stage */
  async getStageFolders(masterStageId: number, opportunityId: number | null = null): Promise<NPPFolder[]> {
    let masterFolders = [];
    let cache = this.masterFolders.find(f => f.stage == masterStageId);
    if (cache) {
      masterFolders = cache.folders;
    } else {
      masterFolders = await this.getAllItems(MASTER_FOLDER_LIST, "$filter=StageNameId eq "+masterStageId);
      console.log('master folders', masterFolders);
      for (let index = 0; index < masterFolders.length; index++) {
        masterFolders[index].containsModels = masterFolders[index].Title === 'Forecast Models';
      }
      this.masterFolders.push({
        stage: masterStageId,
        folders: masterFolders
      });
    }
    
    if (opportunityId) {
      // only folders user can access
      const allowedFolders = await this.getSubfolders(`/${opportunityId}/${masterStageId}`);
      return masterFolders.filter(f => allowedFolders.some((af: any)=> +af.Name === f.ID));
    }
    return masterFolders;
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
