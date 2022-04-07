import { Injectable } from '@angular/core';
import { AppType } from '@shared/models/app-config';
import { Action, Country, EntityGeography, Indication, MasterGeography, Opportunity, OpportunityType, Stage } from '@shared/models/entity';
import { NPPFile, NPPFileMetadata, NPPFolder, SystemFolder } from '@shared/models/file-system';
import { NPPNotification } from '@shared/models/notification';
import { PBIRefreshComponent, PBIReport } from '@shared/models/pbi';
import { GroupPermission, User } from '@shared/models/user';
import { FILES_FOLDER, FOLDER_APPROVED, FOLDER_ARCHIVED, FOLDER_DOCUMENTS, FOLDER_POWER_BI_APPROVED, FOLDER_POWER_BI_ARCHIVED, FOLDER_POWER_BI_DOCUMENTS, FOLDER_POWER_BI_WIP, FOLDER_WIP, FORECAST_MODELS_FOLDER_NAME } from '@shared/sharepoint/folders';
import * as SPLists from '@shared/sharepoint/list-names';
import { ToastrService } from 'ngx-toastr';
import { Observable } from 'rxjs';
import { map } from 'rxjs/operators';
import { GraphService } from './graph.service';
import { LicensingService } from './licensing.service';
import { ReadPermission, SelectInputList, SharepointService } from './sharepoint.service';



interface SPGroup {
  Id: number;
  Title: string;
  Description: string;
  LoginName: string;
  OnlyAllowMembersViewMembership: boolean;
}

interface MasterAction {
  Id: number,
  Title: string;
  ActionNumber: number;
  StageNameId: number;
  OpportunityTypeId: number;
  DueDays: number;
}

interface MasterApprovalStatus {
  Id: number;
  Title: string;
}

type EntityGeographyType = 'Geography' | 'Country';

type EntityGeographyInput = {
  Title: string;
  EntityNameId: number;
  GeographyId?: number;
  CountryId?: number;
  EntityGeographyType: EntityGeographyType 
}

interface OpportunityInput {
  Title: string;
  MoleculeName: string;
  EntityOwnerId: number;
  ProjectStartDate?: Date;
  ProjectEndDate?: Date;
  OpportunityTypeId: number;
  IndicationId: number;
  AppTypeId?: number;
  Year?: number;
  
}

interface StageInput {
  StageUsersId: number[];
  StageReview?: Date;
  Title?: string;
  EntityNameId?: number;
  StageNameId?: number;
}

interface BrandInput {
  Title?: string;
  EntityOwnerId?: number;
  IndicationId?: number;
  BusinessUnitId?: number;
  ForecastCycleId?: number;
  FCDueDate?: Date;
  Year?: number;
  AppTypeId?: number;
  ForecastCycleDescriptor?: string;
}

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
export class AppDataService {

  // local "cache"
  masterBusinessUnits: SelectInputList[] = [];
  masterForecastCycles: SelectInputList[] = [];
  masterOpportunitiesTypes: OpportunityType[] = [];
  masterGroupTypes: GroupPermission[] = [];
  masterCountriesList: SelectInputList[] = [];
  masterGeographiesList: SelectInputList[] = [];
  masterScenariosList: SelectInputList[] = [];
  masterTherapiesList: SelectInputList[] = [];
  masterApprovalStatusList: MasterApprovalStatus[] = [];
  masterGeographies: MasterGeography[] = [];
  masterIndications: {
    therapy: string;
    indications: Indication[]
  }[] = [];
  masterFolders: {
    stage: number;
    folders: NPPFolder[]
  }[] = [];

  public app!: AppType;

  constructor(
    private readonly sharepoint: SharepointService, 
    private readonly msgraph: GraphService,
    private readonly licensing: LicensingService,
    private readonly toastr: ToastrService
  ) { }

  async canConnectAndAccessData(): Promise<boolean> {
    try {
      const currentUser = await this.getCurrentUserInfo();
      const userInfo = await this.getUserInfo(currentUser.Id);
      return true;
    } catch (e) {
      return false;
    }
  }

  getAppType(): AppType {
    return this.app;
  }

  /** read app config values from sharepoint */
  public async getAppConfig() {
    return await this.sharepoint.getAllItems(SPLists.APP_CONFIG_LIST_NAME);
  }

  public async getApp(appId: string) {
    return await this.sharepoint.getAllItems(SPLists.MASTER_APPS_LIST_NAME, "$select=*&$filter=Title eq '" + appId + "'");
  }

  async getEntity(id: number, expand = true): Promise<Opportunity> {
    let options = "$filter=Id eq " + id;
    if (expand) {
      options += "&$select=*,ClinicalTrialPhase/Title,ForecastCycle/Title,BusinessUnit/Title,OpportunityType/Title,Indication/TherapyArea,Indication/ID,Indication/Title,Author/FirstName,Author/LastName,Author/ID,Author/EMail,EntityOwner/ID,EntityOwner/Title,EntityOwner/FirstName,EntityOwner/EMail,EntityOwner/LastName&$expand=OpportunityType,Indication,Author,EntityOwner,BusinessUnit,ClinicalTrialPhase,ForecastCycle";
    }
    return await this.sharepoint.getOneItem(SPLists.ENTITIES_LIST_NAME, options);
  }

    // /** TOCHECK getbrand o get Entity? */
    /* TODEL */
    // async getBrand(id: number): Promise<Opportunity> {
    //   let cond = "&$select=*,Indication/Title,Indication/ID,Indication/TherapyArea,EntityOwner/Title,ForecastCycle/Title,BusinessUnit/Title&$expand=EntityOwner,ForecastCycle,BusinessUnit,Indication";
     
    //   let results = await this.sharepoint.getOneItem(SPLists.ENTITIES_LIST_NAME, "$filter=Id eq "+id+cond);
      
    //   return results;
    // }

  async getAllEntities(appId: number) {
    let countCond = `$filter=AppTypeId eq ${appId}`;
    let max = await this.sharepoint.countItems(SPLists.ENTITIES_LIST_NAME, countCond);

    let cond = countCond+"&$select=*,Indication/Title,Indication/TherapyArea,EntityOwner/Title,ForecastCycle/Title,BusinessUnit/Title&$expand=EntityOwner,ForecastCycle,BusinessUnit,Indication&$skiptoken=Paged=TRUE&$top="+max;
    
    let results = await this.sharepoint.getAllItems(SPLists.ENTITIES_LIST_NAME, cond);
    
    return results;
  }

  async getAllOpportunities(expand = true, onlyActive = false): Promise<Opportunity[]> {
    let filter = undefined;
    if (expand) {
      //TODO check why OpportunityType/isInternal is failing
      filter = "$select=*,ClinicalTrialPhase/Title,OpportunityType/Title,Indication/TherapyArea,Indication/Title,EntityOwner/FirstName,EntityOwner/LastName,EntityOwner/ID,EntityOwner/EMail&$expand=OpportunityType,Indication,EntityOwner,ClinicalTrialPhase";
    }
    if (onlyActive) {
      if (!filter) filter = "$filter=AppTypeId eq '"+this.getAppType().ID+"' and OpportunityStatus eq 'Active'";
      else filter += "&$filter=AppTypeId eq '"+this.getAppType().ID+"' and OpportunityStatus eq 'Active'";
    } else {
      if (!filter) filter = "$filter=AppTypeId eq '"+this.getAppType().ID+"'";
      else filter += "&$filter=AppTypeId eq '"+this.getAppType().ID+"'";
    }

    return await this.sharepoint.getAllItems(SPLists.ENTITIES_LIST_NAME, filter);
  }

  async createEntity(opp: OpportunityInput | BrandInput): Promise<Opportunity> {
    return await this.sharepoint.createItem(SPLists.ENTITIES_LIST_NAME, { OpportunityStatus: "Processing", ...opp });
  }

  async updateEntity(id: number, data: OpportunityInput | BrandInput) {
    return await this.sharepoint.updateItem(id, SPLists.ENTITIES_LIST_NAME, data);
  }

  async deleteOpportunity(oppId: number): Promise<boolean> {
    return await this.sharepoint.deleteItem(oppId, SPLists.ENTITIES_LIST_NAME);
    // TODO Remove all related opportunity info if exists (stages, actions, files...)
  }

  async setOpportunityStatus(opportunityId: number, status: "Processing" | "Archive" | "Active" | "Approved") {
    return this.sharepoint.updateItem(opportunityId, SPLists.ENTITIES_LIST_NAME, {
      OpportunityStatus: status
    });
  }

  async getOpportunityFilterFields() {
    return [
      { value: 'title', label: 'Opportunity Name' },
      { value: 'projectStart', label: 'Project Start Date' },
      { value: 'projectEnd', label: 'Project End Date' },
      { value: 'opportunityType', label: 'Project Type' },
      { value: 'molecule', label: 'Molecule' },
      { value: 'indication', label: 'Indication' },
    ];
  }

  async getBrandFilterFields() {
    return [
      { value: 'Title', label: 'Brand Name' },
      //{ value: 'FCDueDate', label: 'Forecast Cycle Due Date' },
      { value: 'BusinessUnit.Title', label: 'Business Unit' },
      { value: 'Indication.Title', label: 'Indication Name' },
    ];
  }

  /** --- STAGES --- **/

  async createStage(data: StageInput): Promise<Stage | null> {
    if (!data.Title && data.StageNameId) {
      // get from master list
      const masterStage = await this.sharepoint.getOneItemById(data.StageNameId, SPLists.MASTER_STAGES_LIST_NAME);
      Object.assign(data, { Title: masterStage.Title });
    }
    return await this.sharepoint.createItem(SPLists.ENTITY_STAGES_LIST_NAME, data);
  }

  async updateStage(id: number, data: StageInput): Promise<boolean> {
    return await this.sharepoint.updateItem(id, SPLists.ENTITY_STAGES_LIST_NAME, data);
  }

  async getAllStages(): Promise<Stage[]> {
    return await this.sharepoint.getAllItems(SPLists.ENTITY_STAGES_LIST_NAME);
  }

  async getEntityStages(entityId: number): Promise<Stage[]> {
    return await this.sharepoint.getAllItems(SPLists.ENTITY_STAGES_LIST_NAME, "$filter=EntityNameId eq " + entityId);
  }

  async getEntityStage(id: number): Promise<Stage> {
    return await this.sharepoint.getOneItemById(id, SPLists.ENTITY_STAGES_LIST_NAME);
  }

  async getFirstStage(entity: Opportunity) {
    const stageType = await this.getStageType(entity.OpportunityTypeId);
    const firstMasterStage = await this.getMasterStage(stageType, 1);
    return await this.sharepoint.getOneItem(
      SPLists.ENTITY_STAGES_LIST_NAME,
      `$filter=EntityNameId eq ${entity.ID} and StageNameId eq ${firstMasterStage.ID}`
    );
  }

  async getUserInfo(userId: number): Promise<User> {
    return await this.sharepoint.query(`siteusers/getbyid('${userId}')`).toPromise();
  }

  async getUsers(): Promise<User[]> {
    const result = await this.sharepoint.query('siteusers').toPromise();
    if (result.value) {
      return result.value;
    }
    return [];
  }

  async getUserGroups(userId: number): Promise<SPGroup[]> {
    const user = await this.sharepoint.query(`siteusers/getbyid('${userId}')?$expand=groups`).toPromise();
    if (user.Groups.length > 0) {
      return user.Groups;
    }
    return [];
  }

  /** Adds a user to a group */
  async addUserToGroup(user: User, groupId: number): Promise<boolean> {
    return user.LoginName ? await this.sharepoint.addUserToSharepointGroup(user.LoginName, groupId) : false;
  }

  async removeUserFromGroupId(userId: number, groupId: number): Promise<boolean> {
    return await this.sharepoint.removeUserFromSharepointGroup(userId, groupId);
  }

  async removeUserFromGroupName(userId: number, groupName: string): Promise<boolean> {
    return await this.sharepoint.removeUserFromSharepointGroup(userId, groupName);
  }

  /** Create group with name. If group previously exists, get the group */
  async createGroup(name: string, description: string = ''): Promise<SPGroup | null> {
    const group = await this.getGroup(name);
    if (group) return group;

    return this.sharepoint.createGroup(name, description);
  }

  /** Delete group with id */
  async deleteGroup(id: number): Promise<boolean> {
    return this.sharepoint.deleteGroup(id);
  }

  async getGroups(): Promise<SPGroup[]> {
    const groups = await this.sharepoint.query('sitegroups').toPromise();
    if (groups.value) {
      return groups.value;
    }
    return [];
  }

  async userIsInGroup(userId: number, groupId: number): Promise<boolean> {
    try {
      const groupUsers = await this.getGroupMembers(groupId);
      return groupUsers.some(user => user.Id === userId);
    } catch (e) {
      return false;
    }
  }

  /** todel */
  async deleteAllGroups() {
    const groups = await this.getGroups();
    for (const g of groups) {
      if (g.Title.startsWith('DU') || g.Title.startsWith('OO') || g.Title.startsWith('OU') || g.Title.startsWith('SU')) {
        this.deleteGroup(g.Id);
      }
    }
  }

  /** ---- MASTER INFO ---- */
  async getMasterApprovalStatuses(): Promise<MasterApprovalStatus[]> {
    if (this.masterApprovalStatusList.length < 1) {
      this.masterApprovalStatusList = await this.sharepoint.getAllItems(SPLists.MASTER_APPROVAL_STATUS_LIST_NAME);
    }
    return this.masterApprovalStatusList;
  }

  async getMasterActions(stageNameId: number, oppType: number): Promise<MasterAction[]> {
    return await this.sharepoint.getAllItems(
      SPLists.MASTER_ACTION_LIST_NAME,
      `$filter=StageNameId eq ${stageNameId} and OpportunityTypeId eq ${oppType}&$orderby=ActionNumber asc`
    );
  }

  async getMasterApprovalStatusId(status: string): Promise<number | null> {
    const approvalStatus = (await this.getMasterApprovalStatuses()).find(el => el.Title == status);
    if (approvalStatus) {
      return approvalStatus.Id;
    }
    return null;
  }

  async getMasterGeography(id: number): Promise<MasterGeography> {
    const countryExpandOptions = '$select=*,Country/ID,Country/Title&$expand=Country';
    return await this.sharepoint.getOneItemById(id, SPLists.MASTER_GEOGRAPHIES_LIST_NAME, countryExpandOptions);
  }

  async getMasterStageFolders(masterStageId: number): Promise<NPPFolder[]> {
    return await this.sharepoint.getAllItems(SPLists.MASTER_FOLDER_LIST_NAME, "$filter=StageNameId eq " + masterStageId);
  }

  /** TOCHECK any type */
  async getMasterStage(stageType: string, stageNumber: number = 1): Promise<any> {
    return await this.sharepoint.getOneItem(
       SPLists.MASTER_STAGES_LIST_NAME,
      `$select=ID,Title&$filter=(StageType eq '${stageType}') and (StageNumber eq ${stageNumber})`
    );
  }

  async getMasterStageNumbers(stageType: string): Promise<SelectInputList[]> {
    const stages = await this.sharepoint.getAllItems(SPLists.MASTER_STAGES_LIST_NAME, `$filter=StageType eq '${stageType}'`);
    return stages.map(v => { return { label: v.Title, value: v.StageNumber } });
  }

  /** TODEL unused ? */
  // async setApprovalStatus(fileId: number, status: string, comments: string | null = null, folder: string = FILES_FOLDER): Promise<boolean> {
  //   const statusId = await this.getMasterApprovalStatusId(status);
  //   if (!statusId) return false;

  //   let data = { ApprovalStatusId: statusId };
  //   if (comments) Object.assign(data, { Comments: comments });

  //   return await this.sharepoint.updateItem(fileId, `lists/getbytitle('${folder}')`, data);
  // }

  /** TOCHECK on ha d'anar? */
  async setActionDueDate(actionId: number, newDate: string) {
    return await this.sharepoint.updateItem(actionId, SPLists.ENTITY_ACTIONS_LIST_NAME, { ActionDueDate: newDate });
  }

  async getOpportunityTypes(type: string | null = null): Promise<OpportunityType[]> {
    if (this.masterOpportunitiesTypes.length < 1) {
      this.masterOpportunitiesTypes = await this.sharepoint.getAllItems( SPLists.MASTER_OPPORTUNITY_TYPES_LIST_NAME);
    }
    if (type) {
      return this.masterOpportunitiesTypes.filter(el => el.StageType === type);
    }
    return this.masterOpportunitiesTypes;
  }

  async getOpportunityType(OpportunityTypeId: number): Promise<OpportunityType | null> {
    let result: OpportunityType | undefined;
    if (this.masterOpportunitiesTypes.length > 0) {
      result = this.masterOpportunitiesTypes.find(ot => ot.ID === OpportunityTypeId);
    } else {
      result = await this.sharepoint.getOneItem(SPLists.MASTER_OPPORTUNITY_TYPES_LIST_NAME, "$filter=Id eq " + OpportunityTypeId);
    }
    return result ? result : null;
  }

  async getStageType(OpportunityTypeId: number): Promise<string> {
    let result: OpportunityType | undefined;
    if (this.masterOpportunitiesTypes.length > 0) {
      result = this.masterOpportunitiesTypes.find(ot => ot.ID === OpportunityTypeId);
    } else {
      result = await this.sharepoint.getOneItem(SPLists.MASTER_OPPORTUNITY_TYPES_LIST_NAME, "$filter=Id eq " + OpportunityTypeId + "&$select=StageType");
    }
    return result ? result.StageType : '';
  }

  async getIndications(therapy: string = 'all'): Promise<Indication[]> {
    let cache = this.masterIndications.find(i => i.therapy == therapy);
    if (cache) {
      return cache.indications;
    }
    let max = await this.sharepoint.countItems( SPLists.MASTER_THERAPY_AREAS_LIST_NAME);
    let cond = "$skiptoken=Paged=TRUE&$top=" + max;
    if (therapy !== 'all') {
      cond += `&$filter=TherapyArea eq '${therapy}'`;
    }
    let results = await this.sharepoint.getAllItems( SPLists.MASTER_THERAPY_AREAS_LIST_NAME, cond + '&$orderby=TherapyArea asc,Title asc');
    this.masterIndications.push({
      therapy: therapy,
      indications: results
    });
    return results;
  }

  async getGroupPermissions(list: string = ''): Promise<GroupPermission[]> {
    if (this.masterGroupTypes.length < 1) {
      this.masterGroupTypes = await this.sharepoint.getAllItems(SPLists.MASTER_GROUP_TYPES_LIST_NAME);
    }
    if (list) {
      return this.masterGroupTypes.filter(el => el.ListName === list);
    }
    return this.masterGroupTypes;
  }

  async getSiteOwners(): Promise<User[]> {
    const siteTitle = await this.sharepoint.query('title').toPromise();
    if (siteTitle.value) {
      return (await this.getGroupMembers(siteTitle.value + ' Owners'))
        .filter((m: any) => m.Title != 'System Account' && m.UserId); // only "real" users
    }
    return [];
  }

  async getGroupMembers(groupNameOrId: string | number): Promise<User[]> {
    try {
      let users = [];
      if (typeof groupNameOrId == 'number') {
        users = await this.sharepoint.query(`sitegroups/getbyid('${groupNameOrId}')/users`).toPromise();
      } else {
        users = await this.sharepoint.query(`sitegroups/getbyname('${groupNameOrId}')/users`).toPromise();
      }
      if (users && users.value.length > 0) {
        return users.value;
      }
      return [];
    } catch (e) {
      return [];
    }
  }

  /** Returns the Sharepoint Group named as 'name' */
  async getGroup(name: string): Promise<SPGroup | null> {
    try {
      const result = await this.sharepoint.query(`sitegroups/getbyname('${name}')`).toPromise();
      return result;
    } catch (e) {
      return null;
    }
  }

  /** Gets the Id of the group named as 'name' */
  async getGroupId(name: string): Promise<number | null> {
    try {
      const result = await this.sharepoint.query(`sitegroups/getbyname('${name}')/id`).toPromise();
      return result.value;
    } catch (e) {
      return null;
    }
  }

  /** get stage folders. If opportunityId, only the folders with permission. Otherwise, all master folders of stage */
  async getStageFolders(masterStageId: number, opportunityId: number | null = null, businessUnitId: number | null = null): Promise<NPPFolder[]> {
    let masterFolders = [];
    let cache = this.masterFolders.find(f => f.stage == masterStageId);
    if (cache) {
      masterFolders = cache.folders;
    } else {
      masterFolders = await this.getMasterStageFolders(masterStageId);
      for (let index = 0; index < masterFolders.length; index++) {
        masterFolders[index].containsModels = masterFolders[index].Title === FORECAST_MODELS_FOLDER_NAME;
      }
      this.masterFolders.push({
        stage: masterStageId,
        folders: masterFolders
      });
    }

    if (opportunityId && (businessUnitId !== null)) {
      // only folders user can access
      const allowedDepartmentFolders = await this.getSubfolders(`/${businessUnitId}/${opportunityId}/${masterStageId}`);
      const allowedGeoFolders = await this.getSubfolders(`/${businessUnitId}/${opportunityId}/${masterStageId}/0`);
      return masterFolders.filter(f => {
        if (f.containsModels) return allowedGeoFolders.length > 0;
        else return allowedDepartmentFolders.some((af: any) => +af.Name === f.DepartmentID)
      });
    }
    return masterFolders;
  }

  async getNextStage(stageId: number): Promise<Stage | null> {
    // es pot utilitzar getMasterStage() ?
    let current = await this.sharepoint.getOneItemById(stageId, SPLists.MASTER_STAGES_LIST_NAME);
    return await this.getMasterStage(current.StageType, current.StageNumber + 1);
  }

  // TOCHECK
  // es pot substituir la primera crida  per getMasterStage() i la segona per getMasterStageFolders() ?
  /** Recupera els departaments d'una opportunity interna (si entity només els que l'usuari té accés) */
  /** crec que s'hauria de moure a entities services o permissions ? */
  public async getInternalDepartments(entityId: number | null = null, businessUnitId: number | null = null): Promise<NPPFolder[]> {
    let internalStageId = await this.sharepoint.getOneItem(SPLists.MASTER_STAGES_LIST_NAME, "$filter=Title eq 'Internal'");
    let folders = await this.sharepoint.getAllItems(SPLists.MASTER_FOLDER_LIST_NAME, "$filter=StageNameId eq " + internalStageId.ID);
    for (let index = 0; index < folders.length; index++) {
      folders[index].containsModels = folders[index].DepartmentID ? false : true;
    }

    if (entityId && (businessUnitId !== null)) {
      // only folders user can access
      const allowedFolders = await this.getSubfolders(`/${businessUnitId}/${entityId}/0`);
      return folders.filter(f => allowedFolders.some((af: any) => +af.Name === f.DepartmentID));
    }
    return folders;
  }
  
  async getAADGroupName(): Promise<string | null> {
    const AADGroup = await this.sharepoint.getOneItem(SPLists.MASTER_AAD_GROUPS_LIST_NAME, `$filter=AppTypeId eq ${this.getAppType().ID}`);
    if (AADGroup) return AADGroup.Title;
    return null;
  }

  /** --- SELECT LISTS --- */

  searchByTermInputList(query: string, field: string, term: string, matchCase = false): Observable<SelectInputList[]> {
    return this.sharepoint.query(query, '', 'all', { term, field, matchCase })
      .pipe(
        map((res: any) => {
          return res.value.map(
            (el: any) => { return { value: el.Id, label: el.Title } as SelectInputList }
          );
        })
      );
  }

  async getOpportunityTypesList(type: string | null = null): Promise<SelectInputList[]> {
    let res = await this.getOpportunityTypes(type);
    return res.map(t => { return { value: t.ID, label: t.Title, extra: t } });
  }

  async getUsersList(usersId: number[]): Promise<SelectInputList[]> {
    const conditions = usersId.map(e => { return '(Id eq ' + e + ')' }).join(' or ');
    const users = await this.sharepoint.query('siteusers', '$filter=' + conditions).toPromise();
    if (users.value) {
      return users.value.map((u: User) => { return { label: u.Title, value: u.Id } });
    }
    return [];
  }

  async getCountriesList(): Promise<SelectInputList[]> {
    if (this.masterCountriesList.length < 1) {
      let count = await this.sharepoint.countItems(SPLists.MASTER_COUNTRIES_LIST_NAME);
      this.masterCountriesList = (await this.sharepoint.getAllItems(SPLists.MASTER_COUNTRIES_LIST_NAME, `$orderby=Title asc&$top=${count}`)).map(t => { return { value: t.ID, label: t.Title } });
    }
    return this.masterCountriesList;
  }

  async getGeographiesList(): Promise<SelectInputList[]> {
    if (this.masterGeographiesList.length < 1) {
      this.masterGeographiesList = (await this.sharepoint.getAllItems( SPLists.MASTER_GEOGRAPHIES_LIST_NAME, "$orderby=Title asc")).map(t => { return { value: t.ID, label: t.Title } });
    }
    return this.masterGeographiesList;
  }

  /** Accessible Geographies for the user (subfolders with read/write permission) */
  async getAccessibleGeographiesList(entity: Opportunity, stageId: number): Promise<SelectInputList[]> {

    const geographiesList = await this.getEntityGeographies(entity.ID);

    const geoFoldersWithAccess = await this.getSubfolders(`${FILES_FOLDER}/${entity.BusinessUnitId}/${entity.ID}/${stageId}/0`, true);
    return geographiesList.filter(mf => geoFoldersWithAccess.some((gf: any) => +gf.Name === mf.Id))
      .map(t => { return { value: t.Id, label: t.Title } });
  }

  async getScenariosList(): Promise<SelectInputList[]> {
    if (this.masterScenariosList.length < 1) {
      this.masterScenariosList = (await this.sharepoint.getAllItems(SPLists.MASTER_SCENARIOS_LIST_NAME)).map(t => { return { value: t.ID, label: t.Title } });
    }
    return this.masterScenariosList;
  }

  async getClinicalTrialPhases(): Promise<SelectInputList[]> {
    return (await this.sharepoint.getAllItems(SPLists.MASTER_CLINICAL_TRIAL_PHASES_LIST_NAME)).map(t => { return { value: t.ID, label: t.Title } });
  }

  async getIndicationsList(therapy?: string): Promise<SelectInputList[]> {
    let indications = await this.getIndications(therapy);

    if (therapy) {
      return indications.map(el => { return { value: el.ID, label: el.Title } })
    }
    return indications.map(el => { return { value: el.ID, label: el.Title, group: el.TherapyArea } })
  }

  async getTherapiesList(): Promise<SelectInputList[]> {
    if (this.masterTherapiesList.length < 1) {
      let count = await this.sharepoint.countItems(SPLists.MASTER_THERAPY_AREAS_LIST_NAME);
      let indications: Indication[] = await this.sharepoint.getAllItems( SPLists.MASTER_THERAPY_AREAS_LIST_NAME, "$orderby=TherapyArea asc&$skiptoken=Paged=TRUE&$top=" + count);

      return indications
        .map(v => v.TherapyArea)
        .filter((value, index, self) => self.indexOf(value) === index)
        .map(v => { return { label: v, value: v } });
    }
    return this.masterTherapiesList;
  }

  async getSiteOwnersList(): Promise<SelectInputList[]> {
    const owners = await this.getSiteOwners();
    return owners.map(v => { return { label: v.Title ? v.Title : '', value: v.Id } })
  }

  async getBusinessUnitsList(): Promise<SelectInputList[]> {
    let cache = this.masterBusinessUnits;
    if (cache && cache.length) {
      return cache;
    }
    let max = await this.sharepoint.countItems(SPLists.MASTER_BUSINESS_UNIT_LIST_NAME);
    let cond = "$skiptoken=Paged=TRUE&$top="+max;
    let results = await this.sharepoint.getAllItems(SPLists.MASTER_BUSINESS_UNIT_LIST_NAME, cond);
    this.masterBusinessUnits = results.map(el => { return {value: el.ID, label: el.Title }});
    return this.masterBusinessUnits;
  }

  async getForecastCycles(): Promise<SelectInputList[]> {
    let cache = this.masterForecastCycles;
    if (cache && cache.length) {
      return cache;
    }
    let max = await this.sharepoint.countItems(SPLists.MASTER_FORECAST_CYCLES_LIST_NAME);
    let cond = "$skiptoken=Paged=TRUE&$top="+max;
    let results = await this.sharepoint.getAllItems(SPLists.MASTER_FORECAST_CYCLES_LIST_NAME, cond);
    this.masterForecastCycles = results.map(el => { return {value: el.ID, label: el.Title }});
    return this.masterForecastCycles;
  }

  async getEntityAccessibleGeographiesList(entity: Opportunity): Promise<SelectInputList[]> {
    const geographiesList = await this.getEntityGeographies(entity.ID);

    const geoFoldersWithAccess = await this.getSubfolders(`${FOLDER_WIP}/${entity.BusinessUnitId}/${entity.ID}/0/0`, true);
    return geographiesList.filter(mf => geoFoldersWithAccess.some((gf: any) => +gf.Name === mf.Id))
      .map(t => { return { value: t.Id, label: t.Title } });
  }

  /** Gets the profile pic of the user in Microsoft (uses MS Graph) */
  async getUserProfilePic(userId: number): Promise<Blob | null> {
    const user = await this.getUserInfo(userId);
    if (!user.Email) return null;
    return await this.msgraph.getProfilePic(user.Email);
  }

  async createEntityGeography(data: EntityGeographyInput): Promise<EntityGeography> {
    return await this.sharepoint.createItem(SPLists.GEOGRAPHIES_LIST_NAME, data);
  }

  async updateEntityGeography(id: number, data: any): Promise<boolean> {
    return await this.sharepoint.updateItem(id,  SPLists.GEOGRAPHIES_LIST_NAME, data);
  }

  async getEntityGeography(id: number): Promise<EntityGeography> {
    const countryExpandOptions = '$select=*,Country/ID,Country/Title&$expand=Country';
    return await this.sharepoint.getOneItemById(id, SPLists.GEOGRAPHIES_LIST_NAME, countryExpandOptions);
  }

  async getEntityGeographies(entityId: number, all?: boolean) {
    let filter = `$filter=EntityNameId eq ${entityId}`;
    if (!all) {
      filter += ' and Removed ne 1';
    }
    return await this.sharepoint.getAllItems(
       SPLists.GEOGRAPHIES_LIST_NAME, filter,
    );
  }
  
  async createStageActionFromMaster(ma: MasterAction, entityId: number): Promise<Action> {
    let dueDate = new Date();
    dueDate.setDate(dueDate.getDate() + ma.DueDays);
    return await this.sharepoint.createItem(
      SPLists.ENTITY_ACTIONS_LIST_NAME,
      {
        Title: ma.Title,
        StageNameId: ma.StageNameId,
        EntityNameId: entityId,
        ActionNameId: ma.Id,
        ActionDueDate: dueDate
      }
    );
  }

  async getActions(opportunityId: number, stageId?: number): Promise<Action[]> {
    let filterConditions = `(EntityNameId eq ${opportunityId})`;
    if (stageId) filterConditions += ` and (StageNameId eq ${stageId})`;
    return await this.sharepoint.getAllItems(
      SPLists.ENTITY_ACTIONS_LIST_NAME,
      `$select=*,TargetUser/ID,TargetUser/FirstName,TargetUser/LastName&$filter=${filterConditions}&$orderby=StageNameId%20asc&$expand=TargetUser`
    );
  }

  /** TOCHECK passar parametre per filtre select o no */
  async getActionsRaw(opportunityId: number, stageId?: number): Promise<Action[]> {
    let filterConditions = `(EntityNameId eq ${opportunityId})`;
    if (stageId) filterConditions += ` and (StageNameId eq ${stageId})`;
    return await this.sharepoint.getAllItems(
      SPLists.ENTITY_ACTIONS_LIST_NAME,
      `$filter=${filterConditions}&$orderby=Timestamp%20asc`
    );
  }
  
  async completeAction(actionId: number, userId: number): Promise<boolean> {
    const data = {
      TargetUserId: userId,
      Timestamp: new Date(),
      Complete: true
    };
    return await this.sharepoint.updateItem(actionId, SPLists.ENTITY_ACTIONS_LIST_NAME, data);
  }

  async uncompleteAction(actionId: number): Promise<boolean> {
    const data = {
      TargetUserId: null,
      Timestamp: null,
      Complete: false
    };
    return await this.sharepoint.updateItem(actionId, SPLists.ENTITY_ACTIONS_LIST_NAME, data);
  }

  /** Add Power BI Row Level Security Access for the user to the entity */
  async addPowerBI_RLS(user: User, entityId: number, countries: Country[]) {
    const rlsList = await this.sharepoint.getAllItems(
      SPLists.POWER_BI_ACCESS_LIST_NAME, 
      `$filter=TargetUserId eq ${user.Id} and EntityNameId eq ${entityId}`
    );
    for (const country of countries) {
      const rlsItem = rlsList.find(e => e.CountryId == country.ID);
      if (rlsItem) {
        await this.sharepoint.updateItem(rlsItem.Id, SPLists.POWER_BI_ACCESS_LIST_NAME, {
          Removed: "false"
        });
      } else {
        await this.sharepoint.createItem(SPLists.POWER_BI_ACCESS_LIST_NAME, {
          Title: user.Title,
          CountryId: country.ID,
          EntityNameId: entityId,
          TargetUserId: user.Id,
          Removed: false
        });
      }
    }
  }

  /** Remove Power BI Row Level Security Access 
   * 
   * @param entityId The entity to remove the access
   * @param countries List of countries to remove
   * @param userId Remove only the access for the user [optional]
  */
   async removePowerBI_RLS(entityId: number, countries: Country[], userId: number | null = null) {
    let conditions = `$filter=EntityNameId eq ${entityId} and Removed eq 0`;
    if (userId) {
      conditions += ` and TargetUserId eq ${userId}`;
    }
    const rlsList = await this.sharepoint.getAllItems(SPLists.POWER_BI_ACCESS_LIST_NAME, conditions);
    for (const country of countries) {
      const rlsItems = rlsList.filter(e => e.CountryId == country.ID);
      for (const rlsItem of rlsItems) {
        await this.sharepoint.updateItem(rlsItem.Id, SPLists.POWER_BI_ACCESS_LIST_NAME, {
          Removed: "true"
        });
      }
    }
  }

  async createFolder(newFolderUrl: string, isAbsolutePath: boolean = false): Promise<SystemFolder | null> {
    let basePath = FILES_FOLDER;
    if (isAbsolutePath) basePath = '';

    return await this.sharepoint.createFolder(basePath + newFolderUrl);
  }

  async getFolder(folderUrl: string) {
    return await this.sharepoint.getFolderByUrl(folderUrl);
  }

  async assignReadPermissionToFolder(folderUrl: string, groupId: number): Promise<boolean> {
    return await this.assignPermissionToFolder(folderUrl, groupId, ReadPermission);
  }

  async assignPermissionToFolder(folderUrl: string, groupId: number, permission: string) {
    return await this.sharepoint.addRolePermissionToFolder(folderUrl, groupId, permission);
  }

  async assignPermissionToList(listName: string, groupId: number, permission: string, id: number = 0) {
    return await this.sharepoint.addRolePermissionToList(`lists/getbytitle('${listName}')`, groupId, permission, id);
  }

  async getEntityForecastCycles(entity: Opportunity) {
    let filter = `$filter=EntityNameId eq ${entity.ID}`;
    
    return await this.sharepoint.getAllItems(
      SPLists.OPPORTUNITY_FORECAST_CYCLE_LIST_NAME, filter,
    ); 
  }

  async createEntityForecastCycle(entity: Opportunity) {
    return await this.sharepoint.createItem(SPLists.OPPORTUNITY_FORECAST_CYCLE_LIST_NAME, {
      EntityNameId: entity.ID,
      ForecastCycleTypeId: entity.ForecastCycleId,
      Year: entity.Year+"",
      Title: entity.ForecastCycle?.Title + ' ' + entity.Year,
      ForecastCycleDescriptor: entity.ForecastCycleDescriptor
    });    
  }

  /** ----- USERS ----- **/

  async getCurrentUserInfo(): Promise<User> {
    let sharepointUrl = this.licensing.getSharepointApiUri();
    let accountStorageKey = sharepointUrl + '-sharepointAccount';
    let account = localStorage.getItem(accountStorageKey);
    if (account) {
      return JSON.parse(account);
    } else {
      let account = await this.sharepoint.query('currentuser', '$select=Title,Email,Id,FirstName,LastName,IsSiteAdmin').toPromise();
      account['ID'] = account.Id; // set for User interface
      localStorage.setItem(accountStorageKey, JSON.stringify(account));
      return account;
    }
  }

  removeCurrentUserInfo() {
    localStorage.removeItem('sharepointAccount');
  }

  async getSeats(email: string) {
    return await this.licensing.getSeats(email);
  }

  /** TODEL ? */
  async addseattouser(email: string) {
    await this.licensing.addSeat(email);
  }

  /** TODEL ? */
  async removeseattouser(email: string) {
    await this.licensing.removeSeat(email);
  }

  /** --- NOTIFICATIONS --- */

  async getUserNotifications(userId: number, from: Date | false | null = null, limit: number | null = null): Promise<NPPNotification[]> {
    let conditions = `$filter=TargetUserId eq '${userId}'`;
    if (from) {
      conditions += `and Created gt datetime'${from.toISOString()}'`;
    } else if (from === false) {
      conditions += ` and ReadAt eq null`;
    }

    if (limit) conditions += '&$top=' + limit;

    return await this.sharepoint.getAllItems(
      SPLists.NOTIFICATIONS_LIST_NAME,
      conditions + '&$orderby=Created desc'
    );
  }

  async updateNotification(notificationId: number, data: any): Promise<boolean> {
    return await this.sharepoint.updateItem(notificationId, SPLists.NOTIFICATIONS_LIST_NAME, data);
  }

  async notificationsCount(userId: number, conditions = ''): Promise<number> {
    conditions = `$filter=TargetUserId eq '${userId}'` + ( conditions ? ' and ' + conditions : '');
    // item count de sharepoint ho retorna tot sense condicions => getAllItems + length
    return (await this.sharepoint.getAllItems(SPLists.NOTIFICATIONS_LIST_NAME, '$select=Id&' + conditions)).length;
  }

  async createNotification(userId: number, text: string): Promise<NPPNotification> {
    return await this.sharepoint.createItem(SPLists.NOTIFICATIONS_LIST_NAME, {
      Title: text,
      TargetUserId: userId
    });
  }

  /** ---- Power BI ---- **/

  async getReports(): Promise<PBIReport[]>{
    return await this.sharepoint.getAllItems(SPLists.MASTER_POWER_BI_LIST_NAME,'$orderby=SortOrder');
  }

  async getReport(id:number): Promise<PBIReport>{
    return await this.sharepoint.getOneItemById(id, SPLists.MASTER_POWER_BI_LIST_NAME);
  }

  async getReportByName(reportName:string): Promise<PBIReport>{
    let filter = `$filter=Title eq '${reportName}'`;
    let select = `$select=ID,name,GroupId,pageName,Title`;
    return await this.sharepoint.getOneItem(SPLists.MASTER_POWER_BI_LIST_NAME,`${select}&${filter}`)
  }

  async getComponents(report: PBIReport): Promise<PBIRefreshComponent[]> {
    let select = `$select=Title,ComponentType,GroupId`
    let filter = `$filter=ReportTypeId eq'${report.ID}'`;
    let order = '$orderby=ComponentOrder';
    let reportComponents: PBIRefreshComponent[];
    return reportComponents = (await this.sharepoint.getAllItems(SPLists.MASTER_POWER_BI_COMPONENTS_LIST_NAME, `${select}&${filter}&${order}`)).map(t => { return { ComponentType: t.ComponentType, GroupId: t.GroupId, ComponentName: t.Title } })
  }

  /** ---- Files ----- **/

  async readFile(fileUri: string): Promise<any> {
    return await this.sharepoint.readFile(fileUri);
  }

  /** Get all the folder files with properties, if needed */
  async getFolderFiles(folder: string, expandProperties = false): Promise<NPPFile[]> {
    let files: NPPFile[] = []
    const result = await this.sharepoint.getPathFiles(folder);

    if (result.value) {
      files = result.value;
    }

    /** Impossible to expand ListItemAllFields/Author in one query using Sharepoint REST API */
    if (expandProperties && files.length > 0) {
      for (let i = 0; i < files.length; i++) {
        let fileItems = files[i];
        if (fileItems) {
          const info = await this.getEntityFileInfo(folder, fileItems);
          fileItems = Object.assign(fileItems.ListItemAllFields, info);
        }
      }
    }
    return files;
  }

  async getFileProperties(fileUrl: string): Promise<NPPFileMetadata> {
    return await this.sharepoint.readFileMetadata(fileUrl);
  }

  async updateFilePropertiesByPath(filePath: string, properties: any) {
    await this.sharepoint.updateFileFields(filePath, properties);
  }

  async updateFilePropertiesById(fileId: number, rootFolder: string, properties: any) {
    return await this.sharepoint.updateItem(fileId, `lists/getbytitle('${rootFolder}')`, properties);
  }

  async getFileByName(path: string, filename: string) {
    return await this.sharepoint.getPathFiles(path, `$filter=Name eq '${this.clearFileName(filename)}'`);
  }

  async getFileByForecast(path: string, forecastId: number) {
    return await this.sharepoint.getPathFiles(path, `$filter=ListItemAllFields/ForecastId eq ${forecastId}`);
  }

  async deleteFile(fileUri: string): Promise<boolean> {
    return await this.sharepoint.deleteFile(fileUri);
  }

  async renameFile(fileUri: string, newName: string): Promise<boolean> {
    return await this.sharepoint.renameFile(fileUri, this.clearFileName(newName));
  }

  async copyFile(originServerRelativeUrl: string, destinationFolder: string, newFileName: string): Promise<any> {
    return await this.sharepoint.copyFile(originServerRelativeUrl, destinationFolder, this.clearFileName(newFileName));
  }

  async moveFile(originServerRelativeUrl: string, destinationFolder: string, newFilename: string = ''): Promise<any> {
    return await this.sharepoint.moveFile(originServerRelativeUrl, destinationFolder, this.clearFileName(newFilename));
  }

  async cloneFile(originServerRelativeUrl: string, destinationFolder: string, newFileName: string): Promise<boolean> {
    return await this.sharepoint.cloneFile(originServerRelativeUrl, destinationFolder, this.clearFileName(newFileName));
  }

  async existsFile(filename: string, folder: string): Promise<boolean> {
    return await this.sharepoint.existsFile(filename, folder);
  }

  async uploadFile(fileData: string, folder: string, fileName: string, metadata?: any): Promise<any> {
    let uploaded: any = await this.sharepoint.uploadFileQuery(fileData, folder, this.clearFileName(fileName));

    if (metadata && uploaded.ListItemAllFields?.ID/* && uploaded.ServerRelativeUrl*/) {

      // GetFileByServerRelativeUrl('/Folder Name/{file_name}')/CheckOut()
      // GetFileByServerRelativeUrl('/Folder Name/{file_name}')/CheckIn(comment='Comment',checkintype=0)

      await this.sharepoint.updateItem(uploaded.ListItemAllFields.ID, `lists/getbytitle('${FILES_FOLDER}')`, metadata);
    }
    return uploaded;
  }

  /** TOCHECK set type */
  async getSubfolders(folder: string, isAbsolutePath: boolean = false): Promise<any> {
    let basePath = FILES_FOLDER;
    if (isAbsolutePath) basePath = '';
    return await this.sharepoint.getPathSubfolders(basePath + folder);
  }

  async getNPPFolderByDepartment(departmentID: number): Promise<NPPFolder> {
    return await this.sharepoint.getOneItem(SPLists.MASTER_FOLDER_LIST_NAME, "$filter=Id eq " + departmentID);
  }

  /** Adds a user to a Sharepoint group. If ask for seat, also try to assign a seat for the user */
  async addUserToGroupAndSeat(user: User, groupId: number, askForSeat = false): Promise<boolean> {
    try {
      if (askForSeat) {
        //check if is previously in the group, to avoid ask again for the same seat
        if (await this.userIsInGroup(user.Id, groupId)) {
          return true;
        }
        await this.askSeatForUser(user);
      }
      return await this.addUserToGroup(user, groupId);
    } catch (e: any) {
      if (e.status === 422) {
        this.toastr.warning(`Sorry, there are no more free seats for user <${user.Title}>. This \
          user could not be assigned.`, "No Seats Available!", {
          disableTimeOut: true,
          closeButton: true
        });
        return false;
      }
      return false;
    }
  }
  
  /** Remove a user from a Sharepoint group. If removeSeat, also free his seat */
  async removeUserFromGroup(group: string | number, userId: number, removeSeat = false): Promise<boolean> {
    try {
      if (removeSeat) {
        const user = await this.getUserInfo(userId);
        await this.removeUserSeat(user);
      }
      if (typeof group == 'string') {
        return await this.removeUserFromGroupName(userId, group);
      } else {
        return await this.removeUserFromGroupId(userId, group);
      }
    } catch (e: any) {
      if (e.status == 400) {
        return true;
      }
      return false;
    }
  }
  
  private async askSeatForUser(user: User) {
    if (!user.Email) return false;
    try {
      const response = await this.licensing.addSeat(user.Email);
      if (response?.UserGroupsCount == 1) { // assigned seat for first time
        const RLSGroup = await this.getAADGroupName();
        if (RLSGroup) this.msgraph.addUserToPowerBI_RLSGroup(user.Email, RLSGroup);
      }
      return true;
    } catch (e: any) {
      if (e.status === 422) {
        this.toastr.warning(`Sorry, there are no more free seats for user <${user.Title}>. This \
        user could not be assigned.`, "No Seats Available!", {
          disableTimeOut: true,
          closeButton: true
        });
        return false;
      }
      return false;
    }
  }

  private async removeUserSeat(user: User) {
    if (!user.Email) return false;
    try {
      const response = await this.licensing.removeSeat(user.Email);
      if (response?.UserGroupsCount == 0) { // removed the last seat for user
        const RLSGroup = await this.getAADGroupName();
        if (RLSGroup) this.msgraph.removeUserToPowerBI_RLSGroup(user.Email, RLSGroup);
      }
      return true;
    } catch (e: any) {
      if (e.status == 400) {
        return true;
      }
      return false;
    }
  }

  private clearFileName(name: string): string {
    return name.replace(/[~#%&*{}:<>?+|"'/\\]/g, "");
  }

  private async getEntityFileInfo(folder: string, file: NPPFile): Promise<NPPFile> {
    let arrFolder = folder.split("/");
    let rootFolder = arrFolder[0];
    let select = '';
    switch(rootFolder) {
      case FOLDER_DOCUMENTS:
        select = '$select=*,Indication/Title,Indication/ID,Indication/TherapyArea,Author/Id,Author/FirstName,Author/LastName,Editor/Id,Editor/FirstName,Editor/LastName,EntityGeography/Title,EntityGeography/EntityGeographyType,ModelScenario/Title,ApprovalStatus/Title&$expand=Author,Editor,EntityGeography,ModelScenario,Indication,ApprovalStatus';
        break;
      case FOLDER_ARCHIVED:
        select = '$select=*,Indication/Title,Indication/ID,Indication/TherapyArea,Author/Id,Author/FirstName,Author/LastName,Editor/Id,Editor/FirstName,Editor/LastName,EntityGeography/Title,EntityGeography/EntityGeographyType,ModelScenario/Title&$expand=Author,Editor,EntityGeography,ModelScenario,Indication';  
        break;
      default:
        select = '$select=*,Indication/Title,Indication/ID,Indication/TherapyArea,Author/Id,Author/FirstName,Author/LastName,Editor/Id,Editor/FirstName,Editor/LastName,EntityGeography/Title,EntityGeography/EntityGeographyType,ModelScenario/Title,ApprovalStatus/Title&$expand=Author,Editor,EntityGeography,ModelScenario,ApprovalStatus,Indication';
        break;
    }
    
    return await this.sharepoint.query(
      `lists/getbytitle('${rootFolder}')` + `/items(${file.ListItemAllFields?.ID})`,
      select,
      'all'
    ).toPromise();
  }
  
}
