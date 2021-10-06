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
  StageUsersId: number[];
  StageReview: Date;
  Title?: string;
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
  LinkingUri: string;
  TimeLastModified: Date;
  ListItemAllFields?: NPPFileMetadata;
}

export interface NPPFileMetadata {
  ID: number;
  OpportunityNameId: number;
  StageNameId: number;
  ModelApprovalComments: string;
  ApprovalStatusId?: number;
  ApprovalStatus?: any;
  CountryId?: number[];
  GeographyId?: number;
  Geography?: OpportunityGeography;
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
  GeographyID?: boolean;
}

export interface Country {
  ID: number;
  Title: string;
}

export interface OpportunityGeography {
  Id: number;
  Title: string;
  GeographyId: number;
  Geography?: MasterGeography;
  StageId: number;
}

export interface MasterGeography {
  Id: number;
  Title: string;
  CountryId: number[];
}

export interface NPPNotification {
  Id: number;
  Title: string;
  TargetUserId: number;
  TargetUser?: User;
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
const GEOGRAPHIES_LIST_NAME = 'Opportunity Geographies';
const OPPORTUNITIES_LIST = "lists/getbytitle('"+OPPORTUNITES_LIST_NAME+"')";
const OPPORTUNITY_STAGES_LIST = "lists/getbytitle('"+OPPORTUNITY_STAGES_LIST_NAME+"')";
const OPPORTUNITY_ACTIONS_LIST = "lists/getbytitle('"+OPPORTUNITY_ACTIONS_LIST_NAME+"')";
const MASTER_OPPORTUNITY_TYPES_LIST = "lists/getbytitle('Master Opportunity Type List')";
const MASTER_THERAPY_AREAS_LIST = "lists/getbytitle('Master Therapy Areas')";
const MASTER_STAGES_LIST = "lists/getbytitle('Master Stage List')";
const MASTER_ACTION_LIST = "lists/getbytitle('Master Action List')";
const MASTER_FOLDER_LIST = "lists/getByTitle('Master Folder List')";
const MASTER_GROUP_TYPES_LIST = "lists/getByTitle('Master Group Types List')";
const MASTER_APPROVAL_STATUS_LIST = "lists/getByTitle('Master Approval Status')";
const MASTER_GEOGRAPHIES_LIST = "lists/getByTitle('Master Geographies')";
const COUNTRIES_LIST = "lists/getByTitle('Countries')";
const GEOGRAPHIES_LIST = "lists/getByTitle('"+GEOGRAPHIES_LIST_NAME+"')";
const MASTER_SCENARIOS_LIST = "lists/getByTitle('Master Scenarios')";
const USER_INFO_LIST = "lists/getByTitle('User Information List')";
const NOTIFICATIONS_LIST = "lists/getByTitle('Notifications')";
const FILES_FOLDER = "Current Opportunity Library";
const FORECAST_MODELS_FOLDER_NAME = 'Forecast Models';

@Injectable({
  providedIn: 'root'
})
export class SharepointService {

  // local "cache"
  masterOpportunitiesTypes: OpportunityType[] = [];
  masterGroupTypes: GroupPermission[] = [];
  masterCountriesList: SelectInputList[] = [];
  masterGeographiesList: SelectInputList[] = [];
  masterScenariosList: SelectInputList[] = [];
  masterTherapiesList: SelectInputList[] = [];
  masterApprovalStatusList: any[] = [];
  masterGeographies: MasterGeography[] = [];
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

  async test() {
    // const r = await this.query('siteusers').toPromise();
    // const r = await this.query('siteusers', "$filter=isSiteAdmin eq true").toPromise();
    // const r = await this.query("/_vti_bin/ListData.svc/UserInformationList?$filter=IsSiteAdmin eq true").toPromise();
    // const r = await this.getAllItems(USER_INFO_LIST, "$filter=IsSiteAdmin eq true");
    const siteTitle = await this.query('title').toPromise();
    const r = await this.getGroupMembers(siteTitle.value + ' Owners');
    console.log('users', r);
  }

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
    let endpoint = this.licensing.getSharepointApiUri() + partial;
    if (conditions || filterUri) endpoint += '?';
    if (conditions) endpoint += conditions;
    if (filterUri) endpoint += conditions ? '&' + filterUri : filterUri;
    try {
      return this.http.get(endpoint);
    } catch (e) {
      if(e.status == 401) {
        // await this.teams.refreshToken(true); 
      }
      return of([]);
    }
  }

  private async getAllItems(list: string, conditions: string = ''): Promise<any[]> {
    try {
      let endpoint = this.licensing.getSharepointApiUri() + list + '/items';
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

  private async getOneItem(list: string, conditions: string = ''): Promise<any> {
    try {
      let endpoint = this.licensing.getSharepointApiUri() + list + '/items';
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

  private async getOneItemById(id: number, list: string, conditions: string = ''): Promise<any> {
    try {
      let endpoint = this.licensing.getSharepointApiUri() + list + `/items(${id})`;
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

  private async countItems(list: string, conditions: string = ''): Promise<number> {
    try {
      let endpoint = this.licensing.getSharepointApiUri() + list + '/ItemCount';
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

  private async createItem(list: string, data: any): Promise<any> {
    try {
      return await this.http.post(
        this.licensing.getSharepointApiUri() + list + "/items", 
        data
      ).toPromise();
    } catch (e) {
      if(e.status == 401) {
        // await this.teams.refreshToken(true);
      }
      return null;
    }
  }

  private async updateItem(id: number, list: string, data: any): Promise<boolean> {
    try {
      await this.http.post(
        this.licensing.getSharepointApiUri() + list + `/items(${id})`, 
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

  /** --- OPPORTUNITIES --- **/

  async getOpportunities(expand = true): Promise<Opportunity[]> {
    if (expand) {
      return await this.getAllItems(
        OPPORTUNITIES_LIST,
        "$select=*,OpportunityType/Title,Indication/TherapyArea,Indication/Title,OpportunityOwner/FirstName,OpportunityOwner/LastName,OpportunityOwner/ID,OpportunityOwner/EMail&$expand=OpportunityType,Indication,OpportunityOwner"
      );
    }
    return await this.getAllItems(OPPORTUNITIES_LIST);
  }

  async createOpportunity(opp: OpportunityInput, st: StageInput, stageStartNumber: number = 1):
    Promise<{ opportunity: Opportunity, stage: Stage } | false> {

    const opportunity = await this.createItem(OPPORTUNITIES_LIST, { OpportunityStatus: "Processing", ...opp });
    if (!opportunity) return false;

    // get master stage info
    const stageType = await this.getStageType(opp.OpportunityTypeId);
    const masterStage = await this.getMasterStage(stageType, stageStartNumber);

    const stage = await this.createStage(
      { ...st, Title: masterStage.Title, OpportunityNameId: opportunity.ID, StageNameId: masterStage.ID }
    );
    if (!stage) return false; // TODO remove opportunity

    return { opportunity, stage };
  }

  async createGeographies(oppId: number, geographies: number[], countries: number[]) {
    const geographiesList = await this.getGeographiesList();
    const countriesList = await this.getCountriesList();
    
    for (const g of geographies) {
      const newGeo = await this.createItem(GEOGRAPHIES_LIST, {
        Title: geographiesList.find(el => el.value == g)?.label,
        OpportunityId: oppId,
        GeographyId: g
      });
    }
    for (const c of countries) {
      await this.createItem(GEOGRAPHIES_LIST, {
        Title: countriesList.find(el => el.value == c)?.label,
        OpportunityId: oppId,
        CountryId: c
      });
    }
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

    // add groups to the Opp geographies
    permissions = await this.getGroupPermissions(GEOGRAPHIES_LIST_NAME);
    const oppGeographies = await this.getAllItems(GEOGRAPHIES_LIST, '$filter=OpportunityId eq ' + opportunity.ID);
    for (const oppGeo of oppGeographies) {
      await this.setPermissions(permissions, groups, oppGeo.Id);
    }

    await this.initializeStage(opportunity, stage, oppGeographies);

    return true;
  }

  async updateOpportunity(oppId: number, oppData: OpportunityInput): Promise<boolean> {
    const oppBeforeChanges: Opportunity = await this.getOneItemById(oppId, OPPORTUNITIES_LIST);
    const success = await this.updateItem(oppId, OPPORTUNITIES_LIST, oppData);

    if (success && oppBeforeChanges.OpportunityOwnerId !== oppData.OpportunityOwnerId) { // owner changed
      return this.changeOpportunityOwnerPermissions(oppId, oppBeforeChanges.OpportunityOwnerId, oppData.OpportunityOwnerId);
    }

    return success;
  }

  async getOpportunity(id: number): Promise<Opportunity> {
    return await this.getOneItem(OPPORTUNITIES_LIST, "$filter=Id eq "+id+"&$select=*,OpportunityType/Title,Indication/TherapyArea,Indication/Title,Author/FirstName,Author/LastName,Author/ID,Author/EMail,OpportunityOwner/ID,OpportunityOwner/FirstName,OpportunityOwner/EMail,OpportunityOwner/LastName&$expand=OpportunityType,Indication,Author,OpportunityOwner");
  }

  async setOpportunityStatus(opportunityId: number, status: "Processing" | "Archive" | "Active" | "Approved") {
    return this.updateItem(opportunityId, OPPORTUNITIES_LIST, {
      OpportunityStatus: status
    });
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

  async getOpportunityGeographies(oppId: number) {
    return await this.getAllItems(
      GEOGRAPHIES_LIST,
      `$filter=OpportunityId eq ${oppId}`,
    );
  }

  private async createOpportunityGroups(ownerId: number, oppId: number, masterStageId: number): Promise<SPGroupListItem[]> {
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

    return groups;
  }

  async getOpportunityTypes(type: string | null = null): Promise<OpportunityType[]> {
    if (this.masterOpportunitiesTypes.length < 1) {
      this.masterOpportunitiesTypes = await this.getAllItems(MASTER_OPPORTUNITY_TYPES_LIST);
    }
    if (type) {
      return this.masterOpportunitiesTypes.filter(el => el.StageType === type);
    }
    return this.masterOpportunitiesTypes;
  }

  async getOpportunityFields() {
    return [
      { value: 'title', label: 'Opportunity Name' },
      { value: 'projectStart', label: 'Project Start Date' },
      { value: 'projectEnd', label: 'Project End Date' },
      { value: 'opportunityType', label: 'Project Type' },
      { value: 'molecule', label: 'Molecule' },
      { value: 'indication', label: 'Indication' },
    ];
  }


  /** --- STAGES --- **/

  async createStage(data: StageInput): Promise<Stage | null> {
    if (!data.Title && data.StageNameId) {
      // get from master list
      const masterStage = await this.getOneItemById(data.StageNameId, MASTER_STAGES_LIST);
      Object.assign(data, { Title: masterStage.Title });
    }
    return await this.createItem(OPPORTUNITY_STAGES_LIST, data);
  }

  async updateStage(stageId: number, data: any): Promise<boolean> {
    const currentStage = await this.getOneItemById(stageId, OPPORTUNITY_STAGES_LIST);
    let success = await this.updateItem(stageId, OPPORTUNITY_STAGES_LIST, data);

    return success && await this.changeStageUsersPermissions(
      currentStage.OpportunityNameId,
      currentStage.StageNameId,
      currentStage.StageUsersId,
      data.StageUsersId
    );
  }

  async getStages(opportunityId: number): Promise<Stage[]> {
    return await this.getAllItems(OPPORTUNITY_STAGES_LIST, "$filter=OpportunityNameId eq "+opportunityId);
  }

  async getFirstStage(opp: Opportunity) {
    const stageType = await this.getStageType(opp.OpportunityTypeId);
    const firstMasterStage = await this.getMasterStage(stageType, 1);
    return await this.getOneItem(
      OPPORTUNITY_STAGES_LIST, 
      `$filter=OpportunityNameId eq ${opp.ID} and StageNameId eq ${firstMasterStage.ID}`
    );
  }

  async initializeStage(opportunity: Opportunity, stage: Stage, geographies: OpportunityGeography[]): Promise<boolean> {
    const OUGroup = await this.createGroup('OU-'+opportunity.ID);
    const OOGroup = await this.createGroup('OO-'+opportunity.ID);
    const SUGroup = await this.createGroup(`SU-${opportunity.ID}-${stage.StageNameId}`);

    if (!OUGroup || ! OOGroup || !SUGroup) return false; // something happened with groups

    const owner = await this.getUserInfo(opportunity.OpportunityOwnerId);
    if (!owner.LoginName) return false;

    await this.addUserToGroup(owner.LoginName, OUGroup.Id);
    await this.addUserToGroup(owner.LoginName, OOGroup.Id);
    // await this.addUserToGroup(owner.LoginName, SUGroup.Id); // not needed

    let groups: SPGroupListItem[] = [];
    groups.push({ type: 'OU', data: OUGroup });
    groups.push({ type: 'OO', data: OOGroup });
    groups.push({ type: 'SU', data: SUGroup });

    // add groups to the Stage
    let permissions = await this.getGroupPermissions(OPPORTUNITY_STAGES_LIST_NAME);
    await this.setPermissions(permissions, groups, stage.ID);

    // add stage users to group OU and SU
    for (const userId of stage.StageUsersId) {
      const user = await this.getUserInfo(userId);
      if (user.LoginName) {
        await this.addUserToGroup(user.LoginName, OUGroup.Id);
        await this.addUserToGroup(user.LoginName, SUGroup.Id);
      }
    }

    // Actions
    const stageActions = await this.createStageActions(opportunity, stage);

    // add groups into Actions
    permissions = await this.getGroupPermissions(OPPORTUNITY_ACTIONS_LIST_NAME);
    for (const action of stageActions) {
      await this.setPermissions(permissions, groups, action.Id);
    }

    // Folders
    const folders = await this.createStageFolders(stage, geographies, groups);

    // add groups to folders
    permissions = await this.getGroupPermissions(FILES_FOLDER);
    for (const f of folders) {
      if (f.DepartmentID) {
        let folderGroups = [...groups]; // copy default groups
        let DUGroup = await this.createGroup(`DU-${opportunity.ID}-${f.DepartmentID}`, 'Department ID ' + f.DepartmentID);
        if (DUGroup) folderGroups.push( { type: 'DU', data: DUGroup} );
        await this.setPermissions(permissions, folderGroups, f.ServerRelativeUrl);
      }
    }
    return true;
  }

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
    return await this.getMasterStage(current.StageType, current.StageNumber + 1);
  }

  /** get stage folders. If opportunityId, only the folders with permission. Otherwise, all master folders of stage */
  async getStageFolders(masterStageId: number, opportunityId: number | null = null): Promise<NPPFolder[]> {
    let masterFolders = [];
    let cache = this.masterFolders.find(f => f.stage == masterStageId);
    if (cache) {
      masterFolders = cache.folders;
    } else {
      masterFolders = await this.getAllItems(MASTER_FOLDER_LIST, "$filter=StageNameId eq "+masterStageId);
      for (let index = 0; index < masterFolders.length; index++) {
        masterFolders[index].containsModels = masterFolders[index].Title === FORECAST_MODELS_FOLDER_NAME;
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

  private async getMasterStage(stageType: string, stageNumber: number = 1): Promise<any> {
    return await this.getOneItem(
      MASTER_STAGES_LIST, 
      `$select=ID,Title&$filter=(StageType eq '${stageType}') and (StageNumber eq ${stageNumber})`
    );
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

  private async createStageFolders(stage: Stage, geographies: OpportunityGeography[], groups: SPGroupListItem[]): Promise<SystemFolder[]> {
    let oppId = stage.OpportunityNameId;
    
    const masterFolders = await this.getStageFolders(stage.StageNameId);
    await this.createFolder(`/${stage.OpportunityNameId}`);
    await this.createFolder(`/${stage.OpportunityNameId}/${stage.StageNameId}`);

    let folders: SystemFolder[] = [];

    for (const mf of masterFolders) {
      const folder = await this.createFolder(`/${stage.OpportunityNameId}/${stage.StageNameId}/${mf.ID}`);
      if (folder) {
        if (mf.Title !== FORECAST_MODELS_FOLDER_NAME) {
          folder.DepartmentID = mf.DepartmentID;
          folders.push(folder);
        } else {
          for(let geo of geographies) {
            const folder = await this.createFolder(`/${stage.OpportunityNameId}/${stage.StageNameId}/${mf.ID}/${geo.Id}`);
            if (folder) {
              const OUGroup = groups.find(el => el.type=="OU");
              const OOGroup = groups.find(el => el.type=="OO");
              const SUGroup = groups.find(el => el.type=="SU");
              if (!OUGroup || !OOGroup || !SUGroup) throw new Error("Error creating group permissions.");
              
              // department group name
              let groupName = `DU-${oppId}-${mf.ID}-${geo.Id}`;
              const permissions = await this.getGroupPermissions(FILES_FOLDER);
              let DUGroup = await this.createGroup(groupName, 'Department ID ' + mf.ID + ' / Geography ID ' + geo.Id);
                
              if (DUGroup) {
                let folderGroups: SPGroupListItem[] = [...groups, { type: 'DU', data: DUGroup }];
                await this.setPermissions(permissions, folderGroups, folder.ServerRelativeUrl);
              } else {
                throw new Error("Error creating geography group permissions.")
              }
            }
          }
        }
      }
    }
    return folders;
  }

  /** --- OPPORTUNITY ACTIONS --- **/

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

  async getActions(opportunityId: number, stageId?: number): Promise<Action[]> {
    let filterConditions = `(OpportunityNameId eq ${opportunityId})`;
    if (stageId) filterConditions += ` and (StageNameId eq ${stageId})`;
    return await this.getAllItems(
      OPPORTUNITY_ACTIONS_LIST, 
      `$select=*,TargetUser/ID,TargetUser/FirstName,TargetUser/LastName&$filter=${filterConditions}&$orderby=StageNameId%20asc&$expand=TargetUser`
    );
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

  /** --- FILES --- **/
  
  getBaseFilesFolder(): string {
    return FILES_FOLDER;
  }

  async createFolder(newFolderUrl: string): Promise<SystemFolder | null> {
    try {
      return await this.http.post(
        this.licensing.getSharepointApiUri() + "folders", 
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
        this.licensing.getSharepointApiUri() + `GetFileByServerRelativeUrl('${fileUri}')/$value`, 
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
        this.licensing.getSharepointApiUri() + `GetFileByServerRelativeUrl('${fileUri}')`, 
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

  async renameFile(fileUri: string, newName: string): Promise<boolean> {
    try {
      await this.http.post(
        this.licensing.getSharepointApiUri() + `GetFileByServerRelativeUrl('${fileUri}')/ListItemAllFields`, 
        {
          Title: newName,
          FileLeafRef: newName
        },
        {
          headers: new HttpHeaders({
            'If-Match': '*',
            'X-HTTP-Method': "MERGE"
          }),
        }
      ).toPromise();
    } catch (e) {
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

  async existsFile(filename: string, folder: string): Promise<boolean> {
    try {
      let file = await this.query(
        `GetFolderByServerRelativeUrl('${folder}')/Files`,
        `$expand=ListItemAllFields&$filter=Name eq '${filename}'`,
      ).toPromise();
      return file.value.length > 0;
    } catch (e) {
      return false;
    }
  }

  async cloneForecastModel(originFile: NPPFile, newFilename: string, newScenarios: number[], comments = ''): Promise<boolean> {

    const destinationFolder = originFile.ServerRelativeUrl.replace('/' + originFile.Name, '/');

    let success = await this.cloneFile(originFile.ServerRelativeUrl, destinationFolder, newFilename);
    if (!success) return false;

    let newFileInfo = await this.query(
      `GetFolderByServerRelativeUrl('${destinationFolder}')/Files`,
      `$expand=ListItemAllFields&$filter=Name eq '${newFilename}'`,
    ).toPromise();

    if (newFileInfo.value[0].ListItemAllFields && originFile.ListItemAllFields) {
      const newData = {
        ModelScenarioId: newScenarios,
        ModelApprovalComments: comments ? comments : originFile.ListItemAllFields?.ModelApprovalComments,
        ApprovalStatusId: await this.getApprovalStatusId("In Progress")
      }
      success = await this.updateItem(newFileInfo.value[0].ListItemAllFields.ID, `lists/getbytitle('${FILES_FOLDER}')`, newData);
    }

    return success;
  }

  async addScenarioSufixToFilename(originFilename: string, scenarioId: number): Promise<string | false> {
    const scenarios = await this.getScenariosList();
    const extension = originFilename.split('.').pop();
    if (!extension) return false;

    const baseFileName = originFilename.substring(0, originFilename.length - (extension.length + 1));
    return baseFileName
        + '-' + scenarios.find(el => el.value === scenarioId)?.label.replace(/ /g, '').toLocaleLowerCase()
        + '.' + extension;
  }

  async cloneFile(originServerRelativeUrl: string, destinationFolder: string, newFileName: string): Promise<boolean> {
    const originUrl = `getfilebyserverrelativeurl('${originServerRelativeUrl}')/`;
    let destinationUrl = `copyTo('${destinationFolder + newFileName}')`;
    try {
      const r = await this.http.post(
        this.licensing.getSharepointApiUri() + originUrl + destinationUrl,
        null
      ).toPromise();
      return true;
    }
    catch (e) {
      return false;
    }
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
        let fileItems = files[i].ListItemAllFields;
        if (fileItems) {
          fileItems = Object.assign(fileItems, await this.getFileInfo(fileItems.ID));
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
        Country/Title, Geography/Title, ModelScenario/Title, ApprovalStatus/Title \
        &$expand=StageName,Author,TargetUser,Country,Geography,ModelScenario,ApprovalStatus',
      'all'
    ).toPromise();
  }

  async setApprovalStatus(fileId: number, status: string, comments: string | null = null): Promise<boolean> {
    const statusId = await this.getApprovalStatusId(status);
    if (!statusId) return false;

    let data = { ApprovalStatusId: statusId };
    if (comments) Object.assign(data, { ModelApprovalComments: comments });

    return await this.updateItem(fileId, `lists/getbytitle('${FILES_FOLDER}')`, data);
  }

  async getApprovalStatusId(status: string): Promise <number | null> {
    if (this.masterApprovalStatusList.length < 1) {
      this.masterApprovalStatusList = await this.getAllItems(MASTER_APPROVAL_STATUS_LIST);
    }

    const approvalStatus = this.masterApprovalStatusList.find(el => el.Title == status);
    if (approvalStatus) {
      return approvalStatus.Id;
    }
    return null;
  }

  private async uploadFileQuery(fileData: string, folder: string, filename: string) {
    try {
      let url = `GetFolderByServerRelativeUrl('${folder}')/Files/add(url='${filename}',overwrite=true)?$expand=ListItemAllFields`;
      return await this.http.post(
        this.licensing.getSharepointApiUri() + url, 
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

  /** --- PERMISSIONS --- **/

  async createGroup(name: string, description: string = ''): Promise<SPGroup | null> {
    // if exists, return grup
    const group = await this.getGroup(name);
    if (group) return group;

    // otherwise, create group
    try {
      return await this.http.post(
        this.licensing.getSharepointApiUri() + 'sitegroups',
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

  async getGroup(name: string): Promise<SPGroup | null> {
    try {
      const result = await this.query(`sitegroups/getbyname('${name}')`).toPromise();
      return result;
    } catch (e) {
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

  async getGroups(): Promise<SPGroup[]> {
    const groups = await this.query('sitegroups').toPromise();
    if (groups.value) {
      return groups.value;
    }
    return [];
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

  async updateDepartmentUsers(
    oppId: number, 
    stageId: number, 
    departmentId: number, 
    folderDepartmentId: number, 
    geoId: number | null, 
    currentUsersList: number[], 
    newUsersList: number[]
  ): Promise<boolean> {
    // groups needed
    const OUGroup = await this.getGroup('OU-' + oppId);
    const OOGroup = await this.getGroup('OO-' + oppId);
    const SUGroup = await this.getGroup('SU-'  + oppId + '-' + stageId);
    let groupName = `DU-${oppId}-${departmentId}`;
    if (geoId) {
      groupName += `-${geoId}`;
    } 
    const DUGroup = await this.getGroup(groupName);
    
    if (!OUGroup || !OOGroup || !SUGroup || !DUGroup) throw new Error("Permission groups missing.");

    const removedUsers = currentUsersList.filter(item => newUsersList.indexOf(item) < 0);
    const addedUsers = newUsersList.filter(item => currentUsersList.indexOf(item) < 0);

    let success = true;
    for (const userId of removedUsers) {
      success = success && await this.removeUserFromGroup(DUGroup.Id, userId);
      success = success && await this.removeUserFromAllGroups(oppId, userId, ['OU']); // remove (if needed) of OU group
    }

    if (!success) return success;

    for (const userId of addedUsers) {
      const user = await this.getUserInfo(userId);
      if (user.LoginName) {
        success = success && await this.addUserToGroup(user.LoginName, DUGroup.Id);
        success = success && await this.addUserToGroup(user.LoginName, OUGroup.Id);
        if (!success) return success;
      }
    }
    return success;
  }

  async addUserToGroup(loginName: string, groupId: number): Promise<boolean> {
    try {
      await this.http.post(
        this.licensing.getSharepointApiUri() + `sitegroups(${groupId})/users`,
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

  async removeUserFromGroup(group: string | number, userId: number): Promise<boolean> {
    let url = '';
    if (typeof group == 'string') {
      url = this.licensing.getSharepointApiUri() + `sitegroups//getbyname('${group}')/users/removebyid(${userId})`;
    } else if (typeof group == 'number') {
      url = this.licensing.getSharepointApiUri() + `sitegroups(${group})/users/removebyid(${userId})`;
    }
    try {
      await this.http.post(
        url,
        null,
        {
          headers: new HttpHeaders({
            'If-Match': '*',
            'X-HTTP-Method': "DELETE"
          })
        }
      ).toPromise();
      return true;
    } catch (e) {
      if(e.status == 400) {
        return true;
      }
      return false;
    }
  }

  async getUserGroups(userId: number): Promise<SPGroup[]> {
    const user = await this.query(`siteusers/getbyid('${userId}')?$expand=groups`).toPromise();
    if (user.Groups.length > 0) {
      return user.Groups;
    }
    return [];
  }

  async getGroupMembers(groupName: string): Promise<User[]> {
    try {
      let users = await this.query(`sitegroups/getbyname('${groupName}')/users`).toPromise();
      if (users && users.value.length > 0) {
        return users.value;
      }
      return [];
    } catch (e) {
      return [];
    }
  }


  /** todel */
  async deleteAllGroups() {
    const groups = await this.getGroups();
    console.log('groups', groups);
    for (const g of groups) {
      if (g.Title.startsWith('DU') || g.Title.startsWith('OO') || g.Title.startsWith('OU') || g.Title.startsWith('SU')) {
        this.deleteGroup(g.Id);
      }
    }
  }

  async deleteGroup(id: number) {
    try {
      await this.http.post(
        this.licensing.getSharepointApiUri() + `/sitegroups/removebyid(${id})`, 
        null,
        {
          headers: new HttpHeaders({
            'If-Match': '*',
            'X-HTTP-Method': "DELETE"
          })
        }
      ).toPromise();
      return true;
    } catch (e) {
      return false;
    }
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
    const baseUrl = this.licensing.getSharepointApiUri() + list + (id === 0 ? '' : `/items(${id})`);
    return await this.setRolePermission(baseUrl, groupId, roleName);
  }

  private async addRolePermissionToFolder(folderUrl: string, groupId: number, roleName: string): Promise<boolean> {
    const baseUrl = this.licensing.getSharepointApiUri() + `GetFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields`;
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

  private async changeOpportunityOwnerPermissions(oppId: number, currentOwnerId: number, newOwnerId: number): Promise<boolean> {

    const newOwner = await this.getUserInfo(newOwnerId);
    const OOGroup = await this.getGroup('OO-'+oppId); // Opportunity Owner (OO)
    const OUGroup = await this.getGroup('OU-'+oppId); // Opportunity Users (OO)
    if (!newOwner.LoginName || !OOGroup || !OUGroup) return false;

    let success = await this.removeUserFromAllGroups(oppId, currentOwnerId, ['OO', 'OU']);
    
    success = await this.addUserToGroup(newOwner.LoginName, OOGroup.Id) && success;
    return await this.addUserToGroup(newOwner.LoginName, OUGroup.Id) && success;
  }

  private async changeStageUsersPermissions(oppId: number, masterStageId: number, currentUsers: number[], newUsers: number[]): Promise<boolean> {
    const removedUsers = currentUsers.filter(item => newUsers.indexOf(item) < 0);
    const addedUsers = newUsers.filter(item => currentUsers.indexOf(item) < 0);

    let success = true;
    for (const userId of removedUsers) {
      success = success && await this.removeUserFromAllGroups(oppId, userId, ['SU'], masterStageId.toString());
      success = success && await this.removeUserFromAllGroups(oppId, userId, ['OU']); // remove (if needed) of OU group
    }

    if (!success) return false;

    if (addedUsers.length > 0) {
      const OUGroup = await this.getGroup('OU-' + oppId);
      const SUGroup = await this.getGroup(`SU-${oppId}-${masterStageId}`);
      if (!OUGroup || !SUGroup) return false;

      for (const userId of addedUsers) {
        const user = await this.getUserInfo(userId);
        if (user.LoginName) {
          success = success && await this.addUserToGroup(user.LoginName, OUGroup.Id);
          success = success && await this.addUserToGroup(user.LoginName, SUGroup.Id);
          if (!success) return false;
        }
      }
    }
    return success;
  }

  private async removeUserFromAllGroups(oppId: number, userId: number, groups: string[], sufix: string = ''): Promise<boolean> {
    const userGroups = await this.getUserGroups(userId);
    const involvedGroups = userGroups.filter(userGroup => {
      for (const groupType of groups) {
        if (userGroup.Title.startsWith(groupType + '-' + oppId + (sufix ? '-' + sufix : ''))) return true;
      }
      return false;
    });
    let success = true;
    for (const ig of involvedGroups) {
      if (!ig.Title.startsWith('OU')) success = await this.removeUserFromGroup(ig.Title, userId) && success;
    }

    if (!success) return false;

    // has to be removed of OU -> extra check if the user is not in any opportunity group
    if (involvedGroups.some(ig => ig.Title.startsWith('OU'))) {
      const updatedGroups = await this.getUserGroups(userId);
      if (updatedGroups.filter(userGroup => userGroup.Title.split('-')[1] === oppId.toString()).length === 1) {
        // not involved in any group of the opportunity
        success = await this.removeUserFromGroup('OU-' + oppId, userId);
      }
    }
    return success;
  }
  
  /** --- USERS --- **/

  async getUserProfilePic(userId: number): Promise<string> {
    let queryObj = await this.getOneItem(USER_INFO_LIST, `$filter=Id eq ${userId}&$select=Picture`);
    return queryObj.Picture.Url;
  }

  async getCurrentUserInfo(): Promise<User> {
    let account = localStorage.getItem('sharepointAccount');
    if(account) {
      return JSON.parse(account);
    } else {
      let account = await this.query('currentuser', '$select=Title,Email,Id,FirstName,LastName,IsSiteAdmin').toPromise();
      account['ID'] = account.Id; // set for User interface
      localStorage.setItem('sharepointAccount', JSON.stringify(account));
      return account;
    }
  }

  async getUserInfo(userId: number): Promise<User> {
    return await this.query(`siteusers/getbyid('${userId}')`).toPromise();
  }

  async getSiteOwners(): Promise<User[]> {
    const siteTitle = await this.query('title').toPromise();
    if (siteTitle.value) {
      return (await this.getGroupMembers(siteTitle.value + ' Owners'))
      .filter((m: any) => m.Title != 'System Account' && m.UserId); // only "real" users
    }
    return [];
  }

  removeCurrentUserInfo() {
    localStorage.removeItem('sharepointAccount');
  }

  /** --- NOTIFICATIONS --- */

  async getUserNotifications(userId: number): Promise<NPPNotification[]> {
    return await this.getAllItems(
      NOTIFICATIONS_LIST,
      `$filter=TargetUserId eq '${userId}'`
    );
  }

  async createNotification(userId: number, text: string): Promise<NPPNotification> {
    return await this.createItem(NOTIFICATIONS_LIST, {
      Title: text,
      TargetUserId: userId
    });
  }

  /** --- SELECT LISTS --- */

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

  async getOpportunityTypesList(type: string | null = null): Promise<SelectInputList[]> {
    return (await this.getOpportunityTypes(type)).map(t => {return {value: t.ID, label: t.Title}});
  }

  async getUsersList(usersId: number[]): Promise<SelectInputList[]> {
    const conditions = usersId.map(e => { return '(Id eq ' + e + ')' }).join(' or ');
    const users = await this.query('siteusers', '$filter='+conditions).toPromise();
    if (users.value) {
      return users.value.map((u: User) => { return { label: u.Title, value: u.Id }});
    }
    return [];
  }

  async getCountriesList(): Promise<SelectInputList[]> {
    if (this.masterCountriesList.length < 1) {
      this.masterCountriesList = (await this.getAllItems(COUNTRIES_LIST, "$orderby=Title asc")).map(t => {return {value: t.ID, label: t.Title}});
    }
    return this.masterCountriesList;
  }

  async getGeographiesList(): Promise<SelectInputList[]> {
    if (this.masterGeographiesList.length < 1) {
      this.masterGeographiesList = (await this.getAllItems(MASTER_GEOGRAPHIES_LIST, "$orderby=Title asc")).map(t => {return {value: t.ID, label: t.Title}});
    }
    return this.masterGeographiesList;
  }

  /** Accessible Geographies for the user (subfolders with read/write permission) */
  async getAccessibleGeographiesList(oppId: number, stageId: number, departmentID: number): Promise<SelectInputList[]> {
    
    const geographiesList = await this.getAllItems(GEOGRAPHIES_LIST, '$filter=OpportunityId eq ' + oppId);
    
    const geoFoldersWithAccess = await this.getSubfolders(`/${oppId}/${stageId}/${departmentID}`);
    return geographiesList.filter(mf => geoFoldersWithAccess.some((gf: any) => +gf.Name === mf.Id))
      .map(t => {return {value: t.Id, label: t.Title}});
  }

  async getScenariosList(): Promise<SelectInputList[]> {
    if (this.masterScenariosList.length < 1) {
      this.masterScenariosList = (await this.getAllItems(MASTER_SCENARIOS_LIST)).map(t => {return {value: t.ID, label: t.Title}});
    }
    return this.masterScenariosList;
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

  async getSiteOwnersList(): Promise<SelectInputList[]> {
    const owners = await this.getSiteOwners();
    return owners.map(v => { return { label: v.Title ? v.Title : '', value: v.Id }})
  }

  async getMasterStageNumbers(stageType: string): Promise<SelectInputList[]> {
    const stages = await this.getAllItems(MASTER_STAGES_LIST, `$filter=StageType eq '${stageType}'`);
    return stages.map(v => { return { label: v.Title, value: v.StageNumber }});
  }



}
