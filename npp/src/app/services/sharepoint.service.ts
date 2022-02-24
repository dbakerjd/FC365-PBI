import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Observable, of } from 'rxjs';
import { ErrorService } from './error.service';
import { LicensingService } from './licensing.service';
import { map } from 'rxjs/operators';
import { ToastrService } from 'ngx-toastr';
import { GraphService } from './graph.service';
import { ThrowStmt } from '@angular/compiler';


export interface Opportunity {
  ID: number;
  Title: string;
  MoleculeName: string;
  EntityOwnerId: number;
  EntityOwner?: User;
  ProjectStartDate: Date;
  ProjectEndDate: Date;
  OpportunityTypeId: number;
  OpportunityType?: OpportunityType;
  OpportunityStatus: "Processing" | "Archive" | "Active" | "Approved";
  ForecastCycle?:ForecastCycle;
  ForecastCycleId?: number;
  IndicationId: number[];
  Indication: Indication[];
  Modified: Date;
  AuthorId: number;
  Author?: User;
  progress?: number;
  gates?: Stage[];
  isGateType?: boolean;
  BusinessUnitId: number;
  Year: number;
  ClinicalTrialPhaseId: number;
  ClinicalTrialPhase?: ClinicalTrialPhase[];
  ForecastCycleDescriptor: string;
  AppType?: AppType;
  AppTypeId: number;
}

export interface ClinicalTrialPhase {
  ID: number;
  Title: string;
}

export interface OpportunityInput {
  Title: string;
  MoleculeName: string;
  EntityOwnerId: number;
  ProjectStartDate?: Date;
  ProjectEndDate?: Date;
  OpportunityTypeId: number;
  IndicationId: number;
  AppTypeId: number;
  Year?: number;
}

export interface StageInput {
  StageUsersId: number[];
  StageReview: Date;
  Title?: string;
  EntityNameId?: number;
  StageNameId?: number;
}

export interface Action {
  Id: number,
  StageNameId: number;
  EntityNameId: number;
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
  IsInternal: boolean;
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
  EntityNameId: number;
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
  lastComments: FileComments[];
}

export interface FileComments {
  text: string;
  email: string;
  name: string;
  createdAt: string;
}

export interface NPPFileMetadata {
  ID: number;
  EntityNameId?: number;
  StageNameId?: number;
  ApprovalStatusId?: number;
  ApprovalStatus?: any;
  EntityGeographyId?: number;
  EntityGeography?: EntityGeography;
  ModelScenarioId?: number[];
  AuthorId: number;
  Author: User;
  Comments: string;
  IndicationId: number[];
  Indication?: Indication[];
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
  GeographyID?: number;
}

export interface Country {
  ID: number;
  Title: string;
}

export interface EntityGeography {
  Attachments: boolean;
  AuthorId: number;
  ContentTypeId: number;
  CountryId: number;
  Country?: Country;
  Created: Date;
  EditorId: number;
  GeographyId: number;
  Geography?: MasterGeography;
  ID: number;
  Id: number;
  Modified: Date;
  EntityId: number;
  EntityGeographyType: string;
  ServerRedirectedEmbedUri: string;
  ServerRedirectedEmbedUrl: string;
  Title: string;
  Removed: "true" | "false";
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

export interface PBIReport {
  ID: number;
  name: string;
  GroupId: string;
  pageName: string;
  Title: string;
}

const ENTITIES_LIST_NAME = 'Entities';
const ENTITY_STAGES_LIST_NAME = 'Entity Stages';
const ENTITY_ACTIONS_LIST_NAME = 'Entity Action List';
const GEOGRAPHIES_LIST_NAME = 'Entity Geographies';
const OPPORTUNITIES_LIST = "lists/getbytitle('" + ENTITIES_LIST_NAME + "')";
const ENTITY_STAGES_LIST = "lists/getbytitle('" + ENTITY_STAGES_LIST_NAME + "')";
const ENTITY_ACTIONS_LIST = "lists/getbytitle('" + ENTITY_ACTIONS_LIST_NAME + "')";
const MASTER_OPPORTUNITY_TYPES_LIST = "lists/getbytitle('Master Opportunity Type List')";
const MASTER_THERAPY_AREAS_LIST = "lists/getbytitle('Master Therapy Areas')";
const MASTER_STAGES_LIST = "lists/getbytitle('Master Stage List')";
const MASTER_ACTION_LIST = "lists/getbytitle('Master Action List')";
const MASTER_FOLDER_LIST = "lists/getByTitle('Master Folder List')";
const MASTER_GROUP_TYPES_LIST = "lists/getByTitle('Master Group Types List')";
const MASTER_APPROVAL_STATUS_LIST = "lists/getByTitle('Master Approval Status')";
const MASTER_GEOGRAPHIES_LIST = "lists/getByTitle('Master Geographies')";
const COUNTRIES_LIST = "lists/getByTitle('Master Countries')";
const GEOGRAPHIES_LIST = "lists/getByTitle('" + GEOGRAPHIES_LIST_NAME + "')";
const MASTER_SCENARIOS_LIST = "lists/getByTitle('Master Scenarios')";
export const MASTER_CLINICAL_TRIAL_PHASES_LIST = "lists/getByTitle('Master Clinical Trial Phases')";
const USER_INFO_LIST = "lists/getByTitle('User Information List')";
const NOTIFICATIONS_LIST = "lists/getByTitle('Notifications')";
export const FILES_FOLDER = "Current Opportunity Library";
export const FORECAST_MODELS_FOLDER_NAME = 'Forecast Models';
const MASTER_POWER_BI = "lists/getbytitle('Master Power BI')";
const MASTER_AAD_GROUPS = "lists/getbytitle('Master AAD Groups')";
const POWER_BI_ACCESS_LIST = "lists/getbytitle('Power BI Access')";

export interface BusinessUnit {
  ID: number;
  Title: string;
  BUOwnerID: number;
  BUOwner?: User;
  SortOrder: number;
}

export interface ForecastCycle {
  ID: number;
  Title: string;
  ForecastCycleDescriptor: string;
  SortOrder: number;
}

export interface BrandForecastCycle {
  ID: number;
  Title: string;
  BrandId: number;
  Brand?: Brand;
  ForecastCycleTypeId: number;
  ForecastCycleType?: ForecastCycle;
  Year: string;
  ForecastCycleDescriptor: string;
}

export interface OpportunityForecastCycle {
  ID: number;
  Title: string;
  EntityId: number;
  Entity?: Opportunity;
  ForecastCycleTypeId: number;
  ForecastCycleType?: ForecastCycle;
  ForecastCycleDescriptor: string;
  Year: string;
}

export interface Brand {
  ID: number;
  Title: string;
  EntityOwnerId: number;
  EntityOwner?: User;
  BusinessUnitId: number;
  BusinessUnit?: BusinessUnit;
  ForecastCycleId: number;
  ForecastCycle?: ForecastCycle;
  IndicationId: number[];
  Indication?: Indication[];
  FCDueDate?: Date;
  Year: number;
  ForecastCycleDescriptor: string;
  AppType?: AppType;
  AppTypeId: number;
}

export interface BrandInput {
  Title: string;
  EntityOwnerId: number;
  IndicationId: number;
  BusinessUnitId: number;
  ForecastCycleId: number;
  FCDueDate?: Date;
  Year: number;
  AppTypeId: number;
}

export interface AppType {
  ID: number;
  Title: string;
}

export const ENTITIES_LIST = "lists/getbytitle('" + ENTITIES_LIST_NAME + "')";
export const BUSINESS_UNIT_LIST = "lists/getbytitle('Master Business Units')";
export const FORECAST_CYCLES_LIST = "lists/getbytitle('Master Forecast Cycles')";
export const FOLDER_APPROVED = 'Approved Models';
export const FOLDER_ARCHIVED = 'Archived Models';
export const FOLDER_WIP = 'Work in Progress';
export const FOLDER_DOCUMENTS = FILES_FOLDER;
export const FOLDER_POWER_BI_DOCUMENTS = "Power BI Current Opportunity Library";
export const FOLDER_POWER_BI_WIP = "Power BI Work In Progress";
export const FOLDER_POWER_BI_APPROVED = "Power BI Approved Models";
export const FOLDER_POWER_BI_ARCHIVED = "Power BI Archived Models";
export const BRAND_FORECAST_CYCLE = 'Archived Brand Forecast Cycles';
export const BRAND_FORECAST_CYCLE_LIST = "lists/getbytitle('" + BRAND_FORECAST_CYCLE + "')";
export const MASTER_APPS = "lists/getbytitle('Master APPs')";
export const OPPORTUNITY_FORECAST_CYCLE = 'Archived Forecast Cycles';
export const OPPORTUNITY_FORECAST_CYCLE_LIST = "lists/getbytitle('" + OPPORTUNITY_FORECAST_CYCLE + "')";

@Injectable({
  providedIn: 'root'
})
export class SharepointService {

  // local "cache"
  masterBusinessUnits: SelectInputList[] = [];
  masterForecastCycles: SelectInputList[] = [];
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
  provisioningAPI = "https://nppprovisioning20210831.azurewebsites.net/api/";
  public app: AppType | undefined;

  constructor(
    private http: HttpClient, 
    private error: ErrorService, 
    private licensing: LicensingService, 
    private readonly msgraph: GraphService,
    private readonly toastr: ToastrService
  ) { }

  async test() {
    // const r = await this.query('siteusers').toPromise();
    // const r = await this.query('siteusers', "$filter=isSiteAdmin eq true").toPromise();
    // const r = await this.query("/_vti_bin/ListData.svc/UserInformationList?$filter=IsSiteAdmin eq true").toPromise();
    // const r = await this.getAllItems(USER_INFO_LIST, "$filter=IsSiteAdmin eq true");
    // const siteTitle = await this.query('title').toPromise();
    // const r = await this.getGroupMembers(siteTitle.value + ' Owners');
    // console.log('users', r);
  }

  async canConnect(): Promise<boolean> {
    try {
      const currentUser = await this.getCurrentUserInfo();
      const userInfo = await this.getUserInfo(currentUser.Id);
      return true;
    } catch (e) {
      return false;
    }
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
    } catch (e: any) {
      this.error.handleError(e);
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
    } catch (e: any) {
      this.error.handleError(e);
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
    } catch (e: any) {
      this.error.handleError(e);
      return null;
    }
  }

  private async getOneItemById(id: number, list: string, conditions: string = ''): Promise<any> {
    try {
      let endpoint = this.licensing.getSharepointApiUri() + list + `/items(${id})`;
      if (conditions) endpoint += '?' + conditions;
      return await this.http.get(endpoint).toPromise();
    } catch (e: any) {
      this.error.handleError(e);
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
    } catch (e: any) {
      this.error.handleError(e);
      return 0;
    }
  }

  private async createItem(list: string, data: any): Promise<any> {
    try {
      return await this.http.post(
        this.licensing.getSharepointApiUri() + list + "/items",
        data
      ).toPromise();
    } catch (e: any) {
      this.error.handleError(e);
      return null;
    }
  }

  public async updateItem(id: number, list: string, data: any): Promise<boolean> {
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
    } catch (e: any) {
      this.error.handleError(e);
      return false;
    }
    return true;
  }

  public async deleteItem(id: number, list: string): Promise<boolean> {
    try {
      await this.http.post(
        this.licensing.getSharepointApiUri() + list + `/items(${id})`,
        null,
        {
          headers: new HttpHeaders({
            'If-Match': '*',
            'X-HTTP-Method': "DELETE"
          }),
        }
      ).toPromise();
      return true;
    } catch (e: any) {
      this.error.handleError(e);
      return false;
    }
  }

  public async getApp(appId: string) {
    return await this.getAllItems(MASTER_APPS, "$select=*&$filter=Title eq '"+appId+"'");
  }
  /** --- OPPORTUNITIES --- **/

  async getOpportunities(expand = true, onlyActive = false): Promise<Opportunity[]> {
    let filter = undefined;
    if (expand) {
      //TODO check why OpportunityType/isInternal is failing
      filter = "$select=*,ClinicalTrialPhase/Title,OpportunityType/Title,Indication/TherapyArea,Indication/Title,EntityOwner/FirstName,EntityOwner/LastName,EntityOwner/ID,EntityOwner/EMail&$expand=OpportunityType,Indication,EntityOwner,ClinicalTrialPhase";
    }
    if (onlyActive) {
      if (!filter) filter = "$filter=AppTypeId eq '"+this.app?.ID+"' and OpportunityStatus eq 'Active'";
      else filter += "&$filter=AppTypeId eq '"+this.app?.ID+"' and OpportunityStatus eq 'Active'";
    } else {
      if (!filter) filter = "$filter=AppTypeId eq '"+this.app?.ID+"'";
      else filter += "&$filter=AppTypeId eq '"+this.app?.ID+"'";
    }

    /*
    await this.licensing.removeSeat('aspedding@jdforecasting.com');
    // await this.licensing.removeSeat('arandall@jdforecasting.com');
    await this.licensing.removeSeat('BetaNPPDev@janddconsulting.onmicrosoft.com');
    // await this.licensing.removeSeat('awu@jdforecasting.com');
    // await this.licensing.removeSeat('cburrows@jdforecasting.com');
    // await this.licensing.removeSeat('cdavies@jdforecasting.com');
    */
    // await this.licensing.removeSeat('awu@jdforecasting.com');
    // await this.licensing.addSeat('awu@jdforecasting.com');

    return await this.getAllItems(OPPORTUNITIES_LIST, filter);
  }

  async getAllStages(): Promise<Stage[]> {
    return await this.getAllItems(ENTITY_STAGES_LIST);
  }

  async createOpportunity(opp: OpportunityInput, st: StageInput, stageStartNumber: number = 1):
    Promise<{ opportunity: Opportunity, stage: Stage | null } | false> {
    if(this.app) opp.AppTypeId = this.app.ID;
    
    // clean fields according type
    const isInternal = await this.isInternalOpportunity(opp.OpportunityTypeId);
    if (isInternal) {
      opp.ProjectStartDate = opp.ProjectEndDate = undefined;
    } else {
      opp.Year = undefined;
    }

    const opportunity = await this.createItem(OPPORTUNITIES_LIST, { OpportunityStatus: "Processing", ...opp });
    if (!opportunity) return false;

    // get master stage info
    let stage = null;

    if(!isInternal) {
      const opportunityType = await this.getOpportunityType(opp.OpportunityTypeId);
      const stageType = opportunityType?.StageType;
      if(!stageType) throw new Error("Could not determine Opportunity Type");
      const masterStage = await this.getMasterStage(stageType, stageStartNumber);
  
      stage = await this.createStage(
        { ...st, Title: masterStage.Title, EntityNameId: opportunity.ID, StageNameId: masterStage.ID }
      );
      if (!stage) this.deleteOpportunity(opportunity.ID);
    }

    return { opportunity, stage };
  }

  async createGeographies(oppId: number, geographies: number[], countries: number[]): Promise<EntityGeography[]> {
    const geographiesList = await this.getGeographiesList();
    const countriesList = await this.getCountriesList();
    let res: EntityGeography[] = [];
    for (const g of geographies) {
      let newGeo: EntityGeography = await this.createItem(GEOGRAPHIES_LIST, {
        Title: geographiesList.find(el => el.value == g)?.label,
        EntityNameId: oppId,
        GeographyId: g,
        EntityGeographyType: 'Geography'
      });
      res.push(newGeo);
    }
    for (const c of countries) {
      let newGeo: EntityGeography = await this.createItem(GEOGRAPHIES_LIST, {
        Title: countriesList.find(el => el.value == c)?.label,
        EntityNameId: oppId,
        CountryId: c,
        EntityGeographyType: 'Country'
      });
      res.push(newGeo);
    }

    return res;
  }

  async initializeOpportunity(opportunity: Opportunity, stage: Stage | null): Promise<boolean> {
    const groups = await this.createOpportunityGroups(opportunity.EntityOwnerId, opportunity.ID);
    if (groups.length < 1) return false;

    let permissions;
    // add groups to lists
    permissions = (await this.getGroupPermissions()).filter(el => el.ListFilter === 'List');
    await this.setPermissions(permissions, groups);

    // add groups to the Opportunity
    permissions = await this.getGroupPermissions(ENTITIES_LIST_NAME);
    await this.setPermissions(permissions, groups, opportunity.ID);

    // add groups to the Opp geographies
    permissions = await this.getGroupPermissions(GEOGRAPHIES_LIST_NAME);
    const oppGeographies = await this.getAllItems(GEOGRAPHIES_LIST, '$filter=EntityNameId eq ' + opportunity.ID);
    for (const oppGeo of oppGeographies) {
      await this.setPermissions(permissions, groups, oppGeo.Id);
    }

    if (stage) {
      await this.initializeStage(opportunity, stage, oppGeographies);
    } else {
      await this.initializeInternalEntityFolders(opportunity, oppGeographies);
    }
    
    return true;
  }

  /** TODEL ? */
  /*
  async initializeOpportunityAPI(opportunity: Opportunity, stage: Stage) {
    //NewOpportunity?StageID=2&OppID=1&siteUrl=https://janddconsulting.sharepoint.com/sites/NPPDemoV15
    let sharepoint = this.licensing.getSharepointUri();
    return await this.http.get(this.provisioningAPI, {
      params: {
        StageID: stage.ID,
        OppID: opportunity.ID,
        siteUrl: sharepoint ? sharepoint : ''
      }
    }).toPromise();
  }
  */

  async updateOpportunity(oppId: number, oppData: OpportunityInput): Promise<boolean> {
    const oppBeforeChanges: Opportunity = await this.getOneItemById(oppId, OPPORTUNITIES_LIST);
    const success = await this.updateItem(oppId, OPPORTUNITIES_LIST, oppData);

    if (success && oppBeforeChanges.EntityOwnerId !== oppData.EntityOwnerId) { // owner changed
      return this.changeEntityOwnerPermissions(oppId, oppBeforeChanges.EntityOwnerId, oppData.EntityOwnerId);
    }

    return success;
  }

  async deleteOpportunity(oppId: number): Promise<boolean> {
    return await this.deleteItem(oppId, OPPORTUNITIES_LIST);
    // TODO Remove all related opportunity info if exists (stages, actions, files...)
  }

  async getOpportunity(id: number): Promise<Opportunity> {
    return await this.getOneItem(OPPORTUNITIES_LIST, "$filter=Id eq " + id + "&$select=*,ClinicalTrialPhase/Title,ForecastCycle/Title,BusinessUnit/Title,OpportunityType/Title,Indication/TherapyArea,Indication/ID,Indication/Title,Author/FirstName,Author/LastName,Author/ID,Author/EMail,EntityOwner/ID,EntityOwner/Title,EntityOwner/FirstName,EntityOwner/EMail,EntityOwner/LastName&$expand=OpportunityType,Indication,Author,EntityOwner,BusinessUnit,ClinicalTrialPhase,ForecastCycle");
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
    let cond = "$skiptoken=Paged=TRUE&$top=" + max;
    if (therapy !== 'all') {
      cond += `&$filter=TherapyArea eq '${therapy}'`;
    }
    let results = await this.getAllItems(MASTER_THERAPY_AREAS_LIST, cond + '&$orderby=TherapyArea asc,Title asc');
    this.masterIndications.push({
      therapy: therapy,
      indications: results
    });
    return results;
  }

  async getOpportunityGeographies(oppId: number, all?: boolean) {
    let filter = `$filter=EntityNameId eq ${oppId}`;
    if (!all) {
      filter += ' and Deleted ne 1';
    }
    return await this.getAllItems(
      GEOGRAPHIES_LIST, filter,
    );
  }

  private async createOpportunityGroups(ownerId: number, oppId: number): Promise<SPGroupListItem[]> {
    let group;
    let groups: SPGroupListItem[] = [];
    const owner = await this.getUserInfo(ownerId);
    if (!owner.LoginName) return [];

    // Opportunity Users (OU)
    group = await this.createGroup(`OU-${oppId}`);
    if (group) {
      groups.push({ type: 'OU', data: group });
      if (!await this.addUserToGroup(owner, group.Id, true)) {
        return [];
      }
    }

    // Opportunity Owner (OO)
    group = await this.createGroup(`OO-${oppId}`);
    if (group) {
      groups.push({ type: 'OO', data: group });
      await this.addUserToGroup(owner, group.Id);
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
    return await this.createItem(ENTITY_STAGES_LIST, data);
  }

  async updateStage(stageId: number, data: any): Promise<boolean> {
    const currentStage = await this.getOneItemById(stageId, ENTITY_STAGES_LIST);
    let success = await this.updateItem(stageId, ENTITY_STAGES_LIST, data);

    return success && await this.changeStageUsersPermissions(
      currentStage.EntityNameId,
      currentStage.StageNameId,
      currentStage.StageUsersId,
      data.StageUsersId
    );
  }

  async getStages(opportunityId: number): Promise<Stage[]> {
    return await this.getAllItems(ENTITY_STAGES_LIST, "$filter=EntityNameId eq " + opportunityId);
  }

  async getFirstStage(opp: Opportunity) {
    const stageType = await this.getStageType(opp.OpportunityTypeId);
    const firstMasterStage = await this.getMasterStage(stageType, 1);
    return await this.getOneItem(
      ENTITY_STAGES_LIST,
      `$filter=EntityNameId eq ${opp.ID} and StageNameId eq ${firstMasterStage.ID}`
    );
  }

  async initializeInternalEntityFolders(opportunity: Opportunity, geographies: EntityGeography[]) {
    const OUGroup = await this.createGroup('OU-' + opportunity.ID);
    const OOGroup = await this.createGroup('OO-' + opportunity.ID);

    if (!OUGroup || !OOGroup) return false; // something happened with groups

    const owner = await this.getUserInfo(opportunity.EntityOwnerId);
    if (!owner.LoginName) return false;

    if (!await this.addUserToGroup(owner, OUGroup.Id, true)) {
      return false;
    }
    await this.addUserToGroup(owner, OOGroup.Id);
    
    let groups: SPGroupListItem[] = [];
    groups.push({ type: 'OU', data: OUGroup });
    groups.push({ type: 'OO', data: OOGroup });

    // Folders
    const folders = await this.createInternalFolders(opportunity, groups, geographies);

    // add groups to folders
    const RefDocsPermissions = await this.getGroupPermissions(FILES_FOLDER);
    await this.createFolderGroups(opportunity.ID, RefDocsPermissions, folders.rw.filter(el => el.DepartmentID), groups);
    const WIPpermissions = await this.getGroupPermissions(FOLDER_WIP);
    await this.createFolderGroups(opportunity.ID, WIPpermissions, folders.rw.filter(el => el.GeographyID), groups);
    const approvedPermissions = await this.getGroupPermissions(FOLDER_APPROVED);
    await this.createFolderGroups(opportunity.ID, approvedPermissions, folders.ro.filter(el => el.ServerRelativeUrl.includes(FOLDER_APPROVED)), groups);
    const archivedPermissions = await this.getGroupPermissions(FOLDER_ARCHIVED);
    await this.createFolderGroups(opportunity.ID, archivedPermissions, folders.ro.filter(el => el.ServerRelativeUrl.includes(FOLDER_ARCHIVED)), groups);
      
    return true;
  }

  async initializeStage(opportunity: Opportunity, stage: Stage, geographies: EntityGeography[]): Promise<boolean> {
    const OUGroup = await this.createGroup('OU-' + opportunity.ID);
    const OOGroup = await this.createGroup('OO-' + opportunity.ID);
    const SUGroup = await this.createGroup(`SU-${opportunity.ID}-${stage.StageNameId}`);

    if (!OUGroup || !OOGroup || !SUGroup) return false; // something happened with groups

    const owner = await this.getUserInfo(opportunity.EntityOwnerId);
    if (!owner.LoginName) return false;

    if (!await this.addUserToGroup(owner, OUGroup.Id, true)) {
      return false;
    }
    await this.addUserToGroup(owner, OOGroup.Id);
    
    let groups: SPGroupListItem[] = [];
    groups.push({ type: 'OU', data: OUGroup });
    groups.push({ type: 'OO', data: OOGroup });
    groups.push({ type: 'SU', data: SUGroup });

    // add groups to the Stage
    let permissions = await this.getGroupPermissions(ENTITY_STAGES_LIST_NAME);
    await this.setPermissions(permissions, groups, stage.ID);

    // add stage users to group OU and SU
    let addedSU = [];
    for (const userId of stage.StageUsersId) {
      const user = await this.getUserInfo(userId);
      if (!await this.addUserToGroup(user, OUGroup.Id, true)) continue;
      await this.addUserToGroup(user, SUGroup.Id);
      addedSU.push(user.Id);
    }
    if (addedSU.length < 1) {
      // add owner as stage user to don't leave the field blank
      // owner has seat assigned in this point
      await this.addUserToGroup(owner, SUGroup.Id);
      await this.updateItem(stage.ID, ENTITY_STAGES_LIST, { StageUsersId: [owner.Id]});
    } else if (addedSU.length < stage.StageUsersId.length) {
      // update with only the stage users with seat
      await this.updateItem(stage.ID, ENTITY_STAGES_LIST, { StageUsersId: addedSU});
    }

    // Actions
    const stageActions = await this.createStageActions(opportunity, stage);

    // add groups into Actions
    permissions = await this.getGroupPermissions(ENTITY_ACTIONS_LIST_NAME);
    for (const action of stageActions) {
      await this.setPermissions(permissions, groups, action.Id);
    }

    // Folders
    const folders = await this.createStageFolders(opportunity, stage, geographies, groups);

    // add groups to folders
    permissions = await this.getGroupPermissions(FILES_FOLDER);
    await this.createFolderGroups(opportunity.ID, permissions, folders, groups);
    return true;
  }

  /** Creates the DU folder groups and sets permissions for a list of folders 
   * 
   * @param oppId The opportunity ID containing the folders
   * @param permissions List of group permissions to set
   * @param folders List of folders to create the groups
   * @param baseGroups Base of groups to include in the permissions
  */
  private async createFolderGroups(oppId: number, permissions: GroupPermission[], folders: SystemFolder[], baseGroups: SPGroupListItem[]) {
    for (const f of folders) {
      let folderGroups = [...baseGroups]; // copy default groups
      if (f.DepartmentID) {
        let DUGroup = await this.createGroup(`DU-${oppId}-${f.DepartmentID}`, 'Department ID ' + f.DepartmentID);
        if (DUGroup) folderGroups.push({ type: 'DU', data: DUGroup });
      } else if (f.GeographyID) {
        let DUGroup = await this.createGroup(`DU-${oppId}-0-${f.GeographyID}`, 'Geography ID ' + f.GeographyID);
        if (DUGroup) folderGroups.push({ type: 'DU', data: DUGroup });
      }
      await this.setPermissions(permissions, folderGroups, f.ServerRelativeUrl);
    }
  }

  async getStageType(OpportunityTypeId: number): Promise<string> {
    let result: OpportunityType | undefined;
    if (this.masterOpportunitiesTypes.length > 0) {
      result = this.masterOpportunitiesTypes.find(ot => ot.ID === OpportunityTypeId);
    } else {
      result = await this.getOneItem(MASTER_OPPORTUNITY_TYPES_LIST, "$filter=Id eq " + OpportunityTypeId + "&$select=StageType");
    }
    if (result == null) {
      return '';
    }
    return result.StageType;
  }

  async isInternalOpportunity(oppTypeId: number): Promise<boolean> {
    const oppType = await this.getOpportunityType(oppTypeId);
    if (oppType?.IsInternal) {
      return oppType.IsInternal;
    }
    return false;
  }

  async getOpportunityType(OpportunityTypeId: number): Promise<OpportunityType | null> {
    let result: OpportunityType | undefined;
    if (this.masterOpportunitiesTypes.length > 0) {
      result = this.masterOpportunitiesTypes.find(ot => ot.ID === OpportunityTypeId);
    } else {
      result = await this.getOneItem(MASTER_OPPORTUNITY_TYPES_LIST, "$filter=Id eq " + OpportunityTypeId);
    }
    if (result == null) {
      return null;
    }
    return result;
  }

  async getNextStage(stageId: number): Promise<Stage | null> {
    let current = await this.getOneItemById(stageId, MASTER_STAGES_LIST);
    return await this.getMasterStage(current.StageType, current.StageNumber + 1);
  }

  public async getInternalDepartments(entityId: number | null = null, businessUnitId: number | null = null): Promise<NPPFolder[]> {
    let internalStageId = await this.getOneItem(MASTER_STAGES_LIST, "$filter=Title eq 'Internal'");
    let folders = await this.getAllItems(MASTER_FOLDER_LIST, "$filter=StageNameId eq " + internalStageId.ID);
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

  /** get stage folders. If opportunityId, only the folders with permission. Otherwise, all master folders of stage */
  async getStageFolders(masterStageId: number, opportunityId: number | null = null, businessUnitId: number | null = null): Promise<NPPFolder[]> {
    let masterFolders = [];
    let cache = this.masterFolders.find(f => f.stage == masterStageId);
    if (cache) {
      masterFolders = cache.folders;
    } else {
      masterFolders = await this.getAllItems(MASTER_FOLDER_LIST, "$filter=StageNameId eq " + masterStageId);
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

  private async createStageFolders(opportunity: Opportunity, stage: Stage, geographies: EntityGeography[], groups: SPGroupListItem[]): Promise<SystemFolder[]> {

    const OUGroup = groups.find(el => el.type == "OU");
    if (!OUGroup) throw new Error("Error creating group permissions.");

    const masterFolders = await this.getStageFolders(stage.StageNameId);
    const buFolder = await this.createFolder(`/${opportunity.BusinessUnitId}`);
    const oppFolder = await this.createFolder(`/${opportunity.BusinessUnitId}/${stage.EntityNameId}`);
    const stageFolder = await this.createFolder(`/${opportunity.BusinessUnitId}/${stage.EntityNameId}/${stage.StageNameId}`);
    if (!oppFolder || !stageFolder) throw new Error("Error creating opportunity folders.");

    // assign OU to parent folders
    await this.addRolePermissionToFolder(oppFolder.ServerRelativeUrl, OUGroup.data.Id, 'ListRead');
    await this.addRolePermissionToFolder(stageFolder.ServerRelativeUrl, OUGroup.data.Id, 'ListRead');

    let folders: SystemFolder[] = [];

    for (const mf of masterFolders) {
      let folder = await this.createFolder(`/${opportunity.BusinessUnitId}/${stage.EntityNameId}/${stage.StageNameId}/${mf.DepartmentID}`);
      if (folder) {
        if (mf.DepartmentID) {
          folder.DepartmentID = mf.DepartmentID;
          folders.push(folder);
          folder = await this.createFolder(`/${opportunity.BusinessUnitId}/${stage.EntityNameId}/${stage.StageNameId}/${mf.DepartmentID}/0`);
          if (folder) {
            folder = await this.createFolder(`/${opportunity.BusinessUnitId}/${stage.EntityNameId}/${stage.StageNameId}/${mf.DepartmentID}/0/0`);
          }
        } else {
          for (let geo of geographies) {
            let folder = await this.createFolder(`/${opportunity.BusinessUnitId}/${stage.EntityNameId}/${stage.StageNameId}/${mf.DepartmentID}/${geo.Id}`);
            if (folder) {
              folder.DepartmentID = 0;
              folder.GeographyID = geo.ID;
              folders.push(folder);
              folder = await this.createFolder(`/${opportunity.BusinessUnitId}/${stage.EntityNameId}/${stage.StageNameId}/${mf.DepartmentID}/${geo.Id}/0`);
            }
          }
        }
      }
    }

    return folders;
  }

  async getNPPFolderByDepartment(departmentID: number): Promise<NPPFolder> {
    return await this.getOneItem(MASTER_FOLDER_LIST, "$filter=Id eq " + departmentID);
  }

  /** --- OPPORTUNITY ACTIONS --- **/

  private async createAction(ma: MasterAction, oppId: number): Promise<Action> {
    let dueDate = new Date();
    dueDate.setDate(dueDate.getDate() + ma.DueDays);
    return await this.createItem(
      ENTITY_ACTIONS_LIST,
      {
        Title: ma.Title,
        StageNameId: ma.StageNameId,
        EntityNameId: oppId,
        ActionNameId: ma.Id,
        ActionDueDate: dueDate
      }
    );
  }

  async getActions(opportunityId: number, stageId?: number): Promise<Action[]> {
    let filterConditions = `(EntityNameId eq ${opportunityId})`;
    if (stageId) filterConditions += ` and (StageNameId eq ${stageId})`;
    return await this.getAllItems(
      ENTITY_ACTIONS_LIST,
      `$select=*,TargetUser/ID,TargetUser/FirstName,TargetUser/LastName&$filter=${filterConditions}&$orderby=StageNameId%20asc&$expand=TargetUser`
    );
  }

  async getActionsRaw(opportunityId: number, stageId?: number): Promise<Action[]> {
    let filterConditions = `(EntityNameId eq ${opportunityId})`;
    if (stageId) filterConditions += ` and (StageNameId eq ${stageId})`;
    return await this.getAllItems(
      ENTITY_ACTIONS_LIST,
      `$filter=${filterConditions}&$orderby=Timestamp%20asc`
    );
  }

  async completeAction(actionId: number, userId: number): Promise<boolean> {
    const data = {
      TargetUserId: userId,
      Timestamp: new Date(),
      Complete: true
    };
    return await this.updateItem(actionId, ENTITY_ACTIONS_LIST, data);
  }

  async uncompleteAction(actionId: number): Promise<boolean> {
    const data = {
      TargetUserId: null,
      Timestamp: null,
      Complete: false
    };
    return await this.updateItem(actionId, ENTITY_ACTIONS_LIST, data);
  }

  async setActionDueDate(actionId: number, newDate: string) {
    return await this.updateItem(actionId, ENTITY_ACTIONS_LIST, { ActionDueDate: newDate });
  }

  /** --- FILES --- **/

  getBaseFilesFolder(): string {
    return FILES_FOLDER;
  }

  async createFolder(newFolderUrl: string, isAbsolutePath: boolean = false): Promise<SystemFolder | null> {
    try {
      let basePath = FILES_FOLDER;
      if(isAbsolutePath) basePath = '';

      return await this.http.post(
        this.licensing.getSharepointApiUri() + "folders",
        {
          ServerRelativeUrl: basePath + newFolderUrl
        }
      ).toPromise() as SystemFolder;
    } catch (e: any) {
      console.log("Error creating folder: "+e.message);
      this.error.handleError(e);
      return null;
    }
  }

  async getFolderByUrl(folderUrl: string): Promise<SystemFolder | null> {
    try {
      let folder = await this.query(
        `GetFolderByServerRelativeUrl('${folderUrl}')`
      ).toPromise();
      return folder ? folder : null;
    } catch (e) {
      return null;
    }
  }

  async readFile(fileUri: string): Promise<any> {
    try {
      return this.http.get(
        this.licensing.getSharepointApiUri() + `GetFileByServerRelativeUrl('${fileUri}')/$value`,
        { responseType: 'arraybuffer' }
      ).toPromise();
    } catch (e: any) {
      this.error.handleError(e);
      return [];
    }
  }

  async deleteFile(fileUri: string, checkCSV: boolean = true): Promise<boolean> {
    try {
      //First check if it has related CSV files to remove
      if(checkCSV) {
        await this.deleteRelatedCSV(fileUri);
      }
      //then remove
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
    } catch (e: any) {
      this.error.handleError(e);
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
    let uploaded: any = await this.uploadFileQuery(fileData, folder, this.clearFileName(fileName));

    if (metadata && uploaded.ListItemAllFields?.ID/* && uploaded.ServerRelativeUrl*/) {

      // GetFileByServerRelativeUrl('/Folder Name/{file_name}')/CheckOut()
      // GetFileByServerRelativeUrl('/Folder Name/{file_name}')/CheckIn(comment='Comment',checkintype=0)

      await this.updateItem(uploaded.ListItemAllFields.ID, `lists/getbytitle('${FILES_FOLDER}')`, metadata);
    }
    return uploaded;
  }

  async uploadInternalFile(fileData: string, folder: string, fileName: string, metadata?: any): Promise<any> {
    if(metadata) {
      let scenarios = metadata.ModelScenarioId;
      if(scenarios) {
        let file = await this.getFileByScenarios(folder, scenarios);
        if(file) this.deleteFile(file?.ServerRelativeUrl);
      }
    }
    
    let uploaded: any = await this.uploadFileQuery(fileData, folder, this.clearFileName(fileName));

    if (metadata && uploaded.ListItemAllFields?.ID/* && uploaded.ServerRelativeUrl*/) {

      // GetFileByServerRelativeUrl('/Folder Name/{file_name}')/CheckOut()
      // GetFileByServerRelativeUrl('/Folder Name/{file_name}')/CheckIn(comment='Comment',checkintype=0)
      let arrFolder = folder.split("/");
      let rootFolder = arrFolder[0];
      if(!metadata.Comments) {
        metadata.Comments = " ";
      }
      await this.updateItem(uploaded.ListItemAllFields.ID, `lists/getbytitle('${rootFolder}')`, metadata);
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
        Comments: comments ? comments : null,
        ApprovalStatusId: await this.getApprovalStatusId("In Progress")
      }
      success = await this.updateItem(newFileInfo.value[0].ListItemAllFields.ID, `lists/getbytitle('${FILES_FOLDER}')`, newData);
    }

    return success;
  }

  async cloneEntityForecastModel(originFile: NPPFile, newFilename: string, newScenarios: number[], authorId: number, comments = ''): Promise<boolean> {

    const destinationFolder = originFile.ServerRelativeUrl.replace('/' + originFile.Name, '/');

    let success = await this.cloneFile(originFile.ServerRelativeUrl, destinationFolder, newFilename);
    if (!success) return false;

    let newFileInfo = await this.query(
      `GetFolderByServerRelativeUrl('${destinationFolder}')/Files`,
      `$expand=ListItemAllFields&$filter=Name eq '${this.clearFileName(newFilename)}'`,
    ).toPromise();

    if (newFileInfo.value[0].ListItemAllFields && originFile.ListItemAllFields) {
      const newData:any = {
        ModelScenarioId: newScenarios,
        Comments: comments ? comments : null,
        ApprovalStatusId: await this.getApprovalStatusId("In Progress")
      }
      
      let arrFolder = destinationFolder.split("/");
      let rootFolder = arrFolder[3];
      
      success = await this.updateItem(newFileInfo.value[0].ListItemAllFields.ID, `lists/getbytitle('${rootFolder}')`, newData);
      if(success && authorId) {
        const user = await this.getUserInfo(authorId);
        if (user.LoginName)
          await this.updateReadOnlyField(rootFolder, newFileInfo.value[0].ListItemAllFields.ID, 'Editor', user.LoginName);
      }
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

  /** TODEL ? */
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

  /** Impossible to expand ListItemAllFields/Author in one query using Sharepoint REST API */

  async readEntityFolderFiles(folder: string, expandProperties = false): Promise<NPPFile[]> {
    let files: NPPFile[] = []
    const result = await this.query(
      `GetFolderByServerRelativeUrl('${folder}')/Files`,
      '$expand=ListItemAllFields',
    ).toPromise();

    if (result.value) {
      files = result.value;
    }

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

  async getEntityFileFromURL(url: string)  {
    return await this.query(
      `GetFileByServerRelativeUrl('${url}')/listItemAllFields`
    ).toPromise();
  }
  

  async getSubfolders(folder: string, isAbsolutePath: boolean = false): Promise<any> {
    let subfolders: any[] = [];
    let basePath = this.getBaseFilesFolder();
    if(isAbsolutePath) basePath = '';
    const result = await this.query(
      `GetFolderByServerRelativeUrl('${basePath}${folder}')/folders`,
      '$expand=ListItemAllFields',
    ).toPromise();
    if (result.value) {
      subfolders = result.value;
    }
    return subfolders;
  }

  /** TODEL */
  async getFileInfo(fileId: number): Promise<NPPFile> {
    return await this.query(
      `lists/getbytitle('${FILES_FOLDER}')` + `/items(${fileId})`,
      '$select=*,Author/Id,Author/FirstName,Author/LastName,StageName/Id,StageName/Title, \
        EntityGeography/Title,EntityGeography/EntityGeographyType,ModelScenario/Title,ApprovalStatus/Title \
        &$expand=StageName,Author,EntityGeography,ModelScenario,ApprovalStatus',
      'all'
    ).toPromise();
  }

  async setApprovalStatus(fileId: number, status: string, comments: string | null = null, folder: string = FILES_FOLDER): Promise<boolean> {
    const statusId = await this.getApprovalStatusId(status);
    if (!statusId) return false;

    let data = { ApprovalStatusId: statusId };
    if (comments) Object.assign(data, { Comments: comments });

    return await this.updateItem(fileId, `lists/getbytitle('${folder}')`, data);
  }

  async getApprovalStatusId(status: string): Promise<number | null> {
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
      let url = `GetFolderByServerRelativeUrl('${folder}')/Files/add(url='${this.clearFileName(filename)}',overwrite=true)?$expand=ListItemAllFields`;
      return await this.http.post(
        this.licensing.getSharepointApiUri() + url,
        fileData,
        {
          headers: { 'Content-Type': 'blob' }
        }
      ).toPromise();
    } catch (e: any) {
      this.error.handleError(e);
      return {};
    }
  }

  /** --- PERMISSIONS --- **/

  /** Create a Sharepoint group. If previously exists, gets the Group */
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
    } catch (e: any) {
      this.error.handleError(e);
      return null;
    }
  }

  /** Returns the Sharepoint Group named as 'name' */
  async getGroup(name: string): Promise<SPGroup | null> {
    try {
      const result = await this.query(`sitegroups/getbyname('${name}')`).toPromise();
      return result;
    } catch (e) {
      return null;
    }
  }

  /** Gets the Id of the group named as 'name' */
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
        this.SPRoleDefinitions.push({ name, id: result.value }); // add for local caching
        return result.value;
      }
      catch (e) {
        return null;
      }
    }
  }

  /** Sets the access for the entity departments groups updating their members */
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
    let SUGroup = null;
    if(stageId) SUGroup = await this.getGroup('SU-' + oppId + '-' + stageId);
    let groupName = `DU-${oppId}-${departmentId}`;
    let geoCountriesList: Country[] = [];
    if (geoId) {
      groupName += `-${geoId}`;
      geoCountriesList = await this.getCountriesOfEntityGeography(geoId);
    }
    const DUGroup = await this.getGroup(groupName);

    if (!OUGroup || !OOGroup || (!SUGroup && stageId) || !DUGroup) throw new Error("Permission groups missing.");

    const removedUsers = currentUsersList.filter(item => newUsersList.indexOf(item) < 0);
    const addedUsers = newUsersList.filter(item => currentUsersList.indexOf(item) < 0);

    let success = true;
    for (const userId of removedUsers) {
      success = success && await this.removeUserFromGroup(DUGroup.Id, userId);
      if (success && geoId) { // it's model folder
        this.removePowerBI_RLS(oppId, geoCountriesList, userId);
      }
      success = success && await this.removeUserFromAllGroups(oppId, userId, ['OU']); // remove (if needed) of OU group
    }

    if (!success) return success;

    for (const userId of addedUsers) {
      const user = await this.getUserInfo(userId);
      if (!await this.addUserToGroup(user, OUGroup.Id, true)) {
        continue;
      }
      success = success && await this.addUserToGroup(user, DUGroup.Id);
      if (success && geoId) { // it's model folder
        this.addPowerBI_RLS(user, oppId, geoCountriesList);
      }
      if (!success) return success;
    }
    return success;
  }

  /** Adds a user to a Sharepoint group. If ask for seat, also try to assign a seat for the user */
  async addUserToGroup(user: User, groupId: number, askForSeat = false): Promise<boolean> {
    try {
      if (askForSeat && user.Email) {
        //check if is previously in the group, to avoid ask again for the same seat
        if (await this.isInGroup(user.Id, groupId)) {
          return true;
        }
        const response = await this.licensing.addSeat(user.Email);
        if (response?.UserGroupsCount == 1) { // assigned seat for first time
          const RLSGroup = await this.getAADGroupName();
          if (RLSGroup) this.msgraph.addUserToPowerBI_RLSGroup(user.Email, RLSGroup);
        }
      }
      await this.http.post(
        this.licensing.getSharepointApiUri() + `sitegroups(${groupId})/users`,
        { LoginName: user.LoginName }
      ).toPromise();
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

  /** Remove a user from a Sharepoint group. If removeSeat, also free his seat */
  async removeUserFromGroup(group: string | number, userId: number, removeSeat = false): Promise<boolean> {
    let url = '';
    if (typeof group == 'string') {
      url = this.licensing.getSharepointApiUri() + `sitegroups/getbyname('${group}')/users/removebyid(${userId})`;
    } else if (typeof group == 'number') {
      url = this.licensing.getSharepointApiUri() + `sitegroups(${group})/users/removebyid(${userId})`;
    }
    try {
      if (removeSeat) {
        const user = await this.getUserInfo(userId);
        if (user.Email) {
          const response = await this.licensing.removeSeat(user.Email);
          if (response?.UserGroupsCount == 0) { // removed the last seat for user
            const RLSGroup = await this.getAADGroupName();
            if (RLSGroup) this.msgraph.removeUserToPowerBI_RLSGroup(user.Email, RLSGroup);
          }
        }
      }
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
    } catch (e: any) {
      if (e.status == 400) {
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

  async getGroupMembers(groupNameOrId: string | number): Promise<User[]> {
    try {
      let users = [];
      if (typeof groupNameOrId == 'number') {
        users = await this.query(`sitegroups/getbyid('${groupNameOrId}')/users`).toPromise();
      } else {
        users = await this.query(`sitegroups/getbyname('${groupNameOrId}')/users`).toPromise();
      }
      if (users && users.value.length > 0) {
        return users.value;
      }
      return [];
    } catch (e) {
      return [];
    }
  }

  async isInGroup(userId: number, groupId: number): Promise<boolean> {
    try {
      const groupUsers = await this.getGroupMembers(groupId);
      return groupUsers.some(user => user.Id === userId);
    } catch (e) {
      return false;
    }
  }

  /** Add Power BI Row Level Security Access for the user to the entity */
  async addPowerBI_RLS(user: User, entityId: number, countries: Country[]) {
    const rlsList = await this.getAllItems(POWER_BI_ACCESS_LIST, `$filter=TargetUserId eq ${user.Id} and EntityNameId eq ${entityId}`);
    for (const country of countries) {
      const rlsItem = rlsList.find(e => e.CountryId == country.ID);
      if (rlsItem) {
        await this.updateItem(rlsItem.Id, POWER_BI_ACCESS_LIST, {
          Removed: "false"
        });
      } else {
        await this.createItem(POWER_BI_ACCESS_LIST, {
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
    const rlsList = await this.getAllItems(POWER_BI_ACCESS_LIST, conditions);
    for (const country of countries) {
      const rlsItems = rlsList.filter(e => e.CountryId == country.ID);
      for (const rlsItem of rlsItems) {
        await this.updateItem(rlsItem.Id, POWER_BI_ACCESS_LIST, {
          Removed: "true"
        });
      }
    }
  }


  async getAADGroupName(): Promise<string | null> {
    const AADGroup = await this.getOneItem(MASTER_AAD_GROUPS, `$filter=AppTypeId eq ${this.app?.ID}`);
    if (AADGroup) return AADGroup.Title;
    return null;
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

  /** Deletes the sharepoint group by Id */
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
    let folders = [FILES_FOLDER, FOLDER_APPROVED, FOLDER_ARCHIVED, FOLDER_WIP];
    for (const gp of permissions) {
      const group = workingGroups.find(gr => gr.type === gp.Title); // get created group involved on the permission
      if (group) {
        if ((folders.indexOf(gp.ListName) != -1) && typeof itemOrFolder == 'string') {
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
    // permissions to folders without inherit
    let success = await this.setRolePermission(baseUrl, groupId, roleName, false);
    return success && await this.removeRolePermission(baseUrl, (await this.getCurrentUserInfo()).Id);
  }

  private async setRolePermission(baseUrl: string, groupId: number, roleName: string, inherit = true) {
    // const roleId = 1073741826; // READ
    const roleId = await this.getRoleDefinitionId(roleName);
    try {
      await this.http.post(
        baseUrl + `/breakroleinheritance(copyRoleAssignments=${inherit ? 'true' : 'false'},clearSubscopes=${inherit ? 'true' : 'false'})`,
        null).toPromise();
      await this.http.post(
        baseUrl + `/roleassignments/addroleassignment(principalid=${groupId},roledefid=${roleId})`,
        null).toPromise();
      return true;
    } catch (e: any) {
      this.error.handleError(e);
      return false;
    }
  }

  private async removeRolePermission(baseUrl: string, groupId: number) {
    try {
      await this.http.post(
        baseUrl + `/roleassignments/removeroleassignment(principalid=${groupId})`,
        null).toPromise();
      return true;
    } catch (e: any) {
      this.error.handleError(e);
      return false;
    }
  }

  private async changeEntityOwnerPermissions(oppId: number, currentOwnerId: number, newOwnerId: number): Promise<boolean> {

    const newOwner = await this.getUserInfo(newOwnerId);
    const OOGroup = await this.getGroup('OO-' + oppId); // Opportunity Owner (OO)
    const OUGroup = await this.getGroup('OU-' + oppId); // Opportunity Users (OO)
    if (!newOwner.LoginName || !OOGroup || !OUGroup) return false;

    let success = await this.removeUserFromAllGroups(oppId, currentOwnerId, ['OO', 'OU']);

    if (success = await this.addUserToGroup(newOwner, OUGroup.Id, true) && success) {
      success = await this.addUserToGroup(newOwner, OOGroup.Id) && success;
    }

    return success;
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
        if (!(success = await this.addUserToGroup(user, OUGroup.Id, true) && success)) {
          continue;
        }
        success = success && await this.addUserToGroup(user, SUGroup.Id);
        if (!success) return false;
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
        success = await this.removeUserFromGroup('OU-' + oppId, userId, true);
      }
    }
    return success;
  }

  /** --- USERS --- **/

  async getUserProfilePic(userId: number): Promise<string> {
    const user = await this.getUserInfo(userId);
    if (!user) return '';
    //TODO check why graph call is failing...
    return '';
    //return `https://graph.microsoft.com/v1.0/users/${user.Email}/photo/$value`;
  }

  async getCurrentUserInfo(): Promise<User> {
    let account = localStorage.getItem('sharepointAccount');
    if (account) {
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

  async getUsers(): Promise<User[]> {
    const result = await this.query('siteusers').toPromise();
    if (result.value) {
      return result.value;
    }
    return [];
  }

  async getSeats(email: string) {
    return await this.licensing.getSeats(email);
  }

  async addseattouser(email: string) {
    await this.licensing.addSeat(email);
  }

  async removeseattouser(email: string) {
    await this.licensing.removeSeat(email);
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

  async getUserNotifications(userId: number, from: Date | false | null = null, limit: number | null = null): Promise<NPPNotification[]> {
    let conditions = `$filter=TargetUserId eq '${userId}'`;
    if (from) {
      conditions += `and Created gt datetime'${from.toISOString()}'`;
    } else if (from === false) {
      conditions += ` and ReadAt eq null`;
    }

    if (limit) conditions += '&$top=' + limit;

    return await this.getAllItems(
      NOTIFICATIONS_LIST,
      conditions + '&$orderby=Created desc'
    );
  }

  async updateNotification(notificationId: number, data: any): Promise<boolean> {
    return await this.updateItem(notificationId, NOTIFICATIONS_LIST, data);
  }

  async notificationsCount(userId: number, conditions = ''): Promise<number> {
    conditions = `$filter=TargetUserId eq '${userId}'` + ( conditions ? ' and ' + conditions : '');
    // item count ho retorna tot sense condicions => getAllItems + length
    return (await this.getAllItems(NOTIFICATIONS_LIST, '$select=Id&' + conditions)).length;
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
    let res = await this.getOpportunityTypes(type);
    return res.map(t => { return { value: t.ID, label: t.Title, extra: t } });
  }

  async getUsersList(usersId: number[]): Promise<SelectInputList[]> {
    const conditions = usersId.map(e => { return '(Id eq ' + e + ')' }).join(' or ');
    const users = await this.query('siteusers', '$filter=' + conditions).toPromise();
    if (users.value) {
      return users.value.map((u: User) => { return { label: u.Title, value: u.Id } });
    }
    return [];
  }

  async getCountriesList(): Promise<SelectInputList[]> {
    if (this.masterCountriesList.length < 1) {
      let count = await this.countItems(COUNTRIES_LIST);
      this.masterCountriesList = (await this.getAllItems(COUNTRIES_LIST, `$orderby=Title asc&$top=${count}`)).map(t => { return { value: t.ID, label: t.Title } });
    }
    return this.masterCountriesList;
  }

  async getGeographiesList(): Promise<SelectInputList[]> {
    if (this.masterGeographiesList.length < 1) {
      this.masterGeographiesList = (await this.getAllItems(MASTER_GEOGRAPHIES_LIST, "$orderby=Title asc")).map(t => { return { value: t.ID, label: t.Title } });
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
      this.masterScenariosList = (await this.getAllItems(MASTER_SCENARIOS_LIST)).map(t => { return { value: t.ID, label: t.Title } });
    }
    return this.masterScenariosList;
  }

  async getClinicalTrialPhases(): Promise<SelectInputList[]> {
    return (await this.getAllItems(MASTER_CLINICAL_TRIAL_PHASES_LIST)).map(t => { return { value: t.ID, label: t.Title } });
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
      let count = await this.countItems(MASTER_THERAPY_AREAS_LIST);
      let indications: Indication[] = await this.getAllItems(MASTER_THERAPY_AREAS_LIST, "$orderby=TherapyArea asc&$skiptoken=Paged=TRUE&$top=" + count);

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

  async getMasterStageNumbers(stageType: string): Promise<SelectInputList[]> {
    const stages = await this.getAllItems(MASTER_STAGES_LIST, `$filter=StageType eq '${stageType}'`);
    return stages.map(v => { return { label: v.Title, value: v.StageNumber } });
  }

  /** Updates the Entity Geographies with the new sent geographies. 
   *  Creates new geographies and soft delete the old ones including their related groups
   */
  async updateEntityGeographies(entity: Opportunity, newGeographies: string[]) {
    const owner = await this.getUserInfo(entity.EntityOwnerId);
    if (!owner.LoginName) throw new Error("Could not determine entity's owner");
    
    let allGeo: EntityGeography[] = await this.getEntityGeographies(entity.ID, true);

    let neoGeo = newGeographies.filter(el => {
      let arrId = el.split("-");
      let kindOfGeo = arrId[0];
      let id = arrId[1];
      let geo = allGeo.find(el => {
        if (kindOfGeo == 'G') {
          return el.GeographyId == parseInt(id);
        } else {
          return el.CountryId == parseInt(id);
        }
      });

      return !geo;
    });

    let neoCountry = neoGeo.filter(el => {
      let arrId = el.split("-");
      let kindOfGeo = arrId[0];
      return kindOfGeo == 'C';
    }).map(el => {
      let arrId = el.split("-");
      return parseInt(arrId[1]);
    });

    let neoGeography = neoGeo.filter(el => {
      let arrId = el.split("-");
      let kindOfGeo = arrId[0];
      return kindOfGeo == 'G';
    }).map(el => {
      let arrId = el.split("-");
      return parseInt(arrId[1]);
    })

    let restoreGeo: EntityGeography[] = [];
    newGeographies.forEach(el => {
      let arrId = el.split("-");
      let kindOfGeo = arrId[0];
      let id = arrId[1];
      let geo = allGeo.find(el => {
        if (kindOfGeo == 'G') {
          return el.GeographyId == parseInt(id);
        } else {
          return el.CountryId == parseInt(id);
        }
      });

      if (geo && geo.Removed) {
        restoreGeo.push(geo);
      }
    });

    let removeGeo = allGeo.filter(el => {
      let isCountry = !!el.CountryId;
      let geo = newGeographies.find(g => {
        if (isCountry) {
          return g == 'C-' + el.CountryId;
        } else {
          return g == 'G-' + el.GeographyId;
        }
      });

      return !geo && !el.Removed;
    });

    if (removeGeo.length > 0) await this.deleteGeographies(entity, removeGeo);
    if (restoreGeo.length > 0) await this.restoreGeographies(entity, restoreGeo);
    
    let newGeos: EntityGeography[] = [];
    if (neoGeography.length > 0 || neoCountry.length > 0) {
      newGeos = await this.createGeographies(entity.ID, neoGeography, neoCountry);
    }
    if (newGeos.length < 1) return; // finish

    let OOGroup = await this.getGroup(`OO-${entity.ID}`);
    let OUGroup = await this.getGroup(`OU-${entity.ID}`);
    if (!OOGroup || !OUGroup) throw new Error("Error obtaining user groups.");

    let groups: SPGroupListItem[] = [];
    groups.push({ type: 'OO', data: OOGroup });
    groups.push({ type: 'OU', data: OUGroup });

    let permissions = await this.getGroupPermissions(GEOGRAPHIES_LIST_NAME);
    let stages = await this.getStages(entity.ID);
    if (stages && stages.length) {
      for (const oppGeo of newGeos) {
        await this.setPermissions(permissions, groups, oppGeo.Id); // assign permissions to new entity geo items
        for (let index = 0; index < stages.length; index++) {
          let stage = stages[index];
          let stageFolders = await this.getStageFolders(stage.StageNameId, entity.ID, entity.BusinessUnitId);
          let mf = stageFolders.find(el => el.Title == FORECAST_MODELS_FOLDER_NAME);
  
          if (!mf) throw new Error("Could not find Models folder");
  
          let folder = await this.createFolder(`/${entity.BusinessUnitId}/${entity.ID}/${stage.StageNameId}/${mf.DepartmentID}/${oppGeo.Id}`);
          if(folder) {
            // department group and Stage Users Group
            const DUGroupName = `DU-${entity.ID}-${mf.DepartmentID}-${oppGeo.Id}`;
            let DUGroup = await this.createGroup(DUGroupName, 'Department ID ' + mf.DepartmentID + ' / Geography ID ' + oppGeo.Id);
            let SUGroup = await this.getGroup(`SU-${entity.ID}-${stage.StageNameId}`);
            if (DUGroup && SUGroup) {
              const permissions = await this.getGroupPermissions(FILES_FOLDER);
              let folderGroups: SPGroupListItem[] = [...groups, { type: 'DU', data: DUGroup }, { type: 'SU', data: SUGroup }];
              await this.setPermissions(permissions, folderGroups, folder.ServerRelativeUrl);
            } else {
              if (!DUGroup) throw new Error("Error creating geography group permissions.");
              else throw new Error("Error getting SU group.");
            }
            await this.createFolder(`/${entity.BusinessUnitId}/${entity.ID}/${stage.StageNameId}/${mf.DepartmentID}/${oppGeo.Id}/0`);
          }
        }
      } 
    } else {
      const folders = await this.createInternalFolders(entity, groups, newGeos);
      
      for (const oppGeo of newGeos) {
        await this.setPermissions(permissions, groups, oppGeo.Id); // assign permissions to new entity geo items
      }
      // add groups to folders
      // (department folders non needed)
      // const departmentPermissions = await this.getGroupPermissions(FILES_FOLDER);
      // await this.createFolderGroups(entity.ID, departmentPermissions, folders.rw.filter(el => el.DepartmentID), groups);
      const WIPPermissions = await this.getGroupPermissions(FOLDER_WIP);
      await this.createFolderGroups(entity.ID, WIPPermissions, folders.rw.filter(el => el.GeographyID), groups);
      const approvedPermissions = await this.getGroupPermissions(FOLDER_APPROVED);
      await this.createFolderGroups(entity.ID, approvedPermissions, folders.ro.filter(el => el.ServerRelativeUrl.includes(FOLDER_APPROVED)), groups);
      const archivedPermissions = await this.getGroupPermissions(FOLDER_ARCHIVED);
      await this.createFolderGroups(entity.ID, archivedPermissions, folders.ro.filter(el => el.ServerRelativeUrl.includes(FOLDER_ARCHIVED)), groups);
    }
  }

  /** Soft delete entity geographies. Delete DU geography groups related */
  private async deleteGeographies(entity: Opportunity, removeGeos: EntityGeography[]) {
    //removes groups
    let stages = await this.getStages(entity.ID);
    if (stages && stages.length) {
      // external
      for (const geo of removeGeos) {
        for (const stage of stages) {
          let stageFolders = await this.getStageFolders(stage.StageNameId, entity.ID, entity.BusinessUnitId);
          let modelFolders = stageFolders.filter(el => el.containsModels);
          if (modelFolders.length < 1) continue;

          for (const mf of modelFolders) {
            const DUGroupId = await this.getGroupId(`DU-${entity.ID}-${mf.DepartmentID}-${geo.Id}`);
            if (DUGroupId) await this.deleteGroup(DUGroupId);
          }
        }
      }
    } else {
      // internal
      for (const geo of removeGeos) {
        const DUGroupId = await this.getGroupId(`DU-${entity.ID}-0-${geo.Id}`);
        if (DUGroupId) await this.deleteGroup(DUGroupId);
      }
    }

    // soft delete entity geographies
    for (let i = 0; i < removeGeos.length; i++) {
      await this.updateItem(removeGeos[i].ID, GEOGRAPHIES_LIST, {
        Removed: "true"
      });

      // Power BI RLS access 
      const geoCountriesList = await this.getCountriesOfEntityGeography(removeGeos[i].ID);
      await this.removePowerBI_RLS(entity.ID, geoCountriesList);
    }
  }

  /** Restore previously soft deleted entity geographies and create DU groups related */
  private async restoreGeographies(entity: Opportunity, restoreGeos: EntityGeography[]) {
    let OOGroup = await this.getGroup(`OO-${entity.ID}`);
    let OUGroup = await this.getGroup(`OU-${entity.ID}`);
    if (!OOGroup || !OUGroup) throw new Error("Error obtaining user groups.");

    let groups: SPGroupListItem[] = [];
    groups.push({ type: 'OO', data: OOGroup });
    groups.push({ type: 'OU', data: OUGroup });

    let stages = await this.getStages(entity.ID);
    if (stages && stages.length) {
      // external
      for (const geo of restoreGeos) {
        for (const stage of stages) {
          let stageFolders = await this.getStageFolders(stage.StageNameId, entity.ID, entity.BusinessUnitId);
          let modelFolders = stageFolders.filter(el => el.containsModels);
          if (modelFolders.length < 1) continue;

          // not needed because SU group is never removed
          // let SUGroup = await this.createGroup(`SU-${entity.ID}-${stage.StageNameId}`);
          // if (!SUGroup) throw new Error('Error obtaining user group (SU)');

          const permissions = await this.getGroupPermissions(FILES_FOLDER);
          for (const mf of modelFolders) {
            const folder = await this.getFolderByUrl(this.getBaseFilesFolder() + `/${entity.BusinessUnitId}/${entity.ID}/${stage.StageNameId}/${mf.DepartmentID}/${geo.Id}`);
            const DUGroupName = `DU-${entity.ID}-${mf.DepartmentID}-${geo.Id}`;
            let DUGroup = await this.createGroup(DUGroupName, 'Department ID ' + mf.DepartmentID + ' / Geography ID ' + geo.Id);
            if (folder && DUGroup) {
              groups.push({ type: 'DU', data: DUGroup });
              await this.createFolderGroups(entity.ID, permissions, [folder], groups);
            }
          }
        }
      }
    } else {
      // internal
      const folders = await this.createInternalFolders(entity, groups, restoreGeos);

      const WIPPermissions = await this.getGroupPermissions(FOLDER_WIP);
      await this.createFolderGroups(entity.ID, WIPPermissions, folders.rw.filter(el => el.GeographyID), groups);
      const approvedPermissions = await this.getGroupPermissions(FOLDER_APPROVED);
      await this.createFolderGroups(entity.ID, approvedPermissions, folders.ro.filter(el => el.ServerRelativeUrl.includes(FOLDER_APPROVED)), groups);
      const archivedPermissions = await this.getGroupPermissions(FOLDER_ARCHIVED);
      await this.createFolderGroups(entity.ID, archivedPermissions, folders.ro.filter(el => el.ServerRelativeUrl.includes(FOLDER_ARCHIVED)), groups);   
    }

    // restore entity geographies
    for (let i = 0; i < restoreGeos.length; i++) {
      await this.updateItem(restoreGeos[i].ID, GEOGRAPHIES_LIST, {
        Removed: "false"
      });
    }
  }

  /** Returns the entire list of countries related to Entity Geography */
  async getCountriesOfEntityGeography(geoId: number): Promise<Country[]> {
    const countryExpandOptions = '$select=*,Country/ID,Country/Title&$expand=Country';
    const entityGeography: EntityGeography = await this.getOneItemById(geoId, GEOGRAPHIES_LIST, countryExpandOptions);
    if (entityGeography.CountryId && entityGeography.Country) {
      return [entityGeography.Country];
    }
    else if (entityGeography.GeographyId) {
      const masterGeography = await this.getOneItemById(entityGeography.GeographyId, MASTER_GEOGRAPHIES_LIST, countryExpandOptions);
      return masterGeography.Country;
    }
    return [];
  }

  async getReports(): Promise<PBIReport[]>{
    return await this.getAllItems(MASTER_POWER_BI,'$orderby=SortOrder');
  }

  async getReport(id:number): Promise<PBIReport>{
    return await this.getOneItemById(id,MASTER_POWER_BI);
  }

  /** TODEL ? */
  async createBrand(b: BrandInput, geographies: number[], countries: number[]): Promise<Brand|undefined> {
    const owner = await this.getUserInfo(b.EntityOwnerId);
    if (!owner.LoginName) throw new Error("Could not obtain owner's information.");
    if(this.app) b.AppTypeId = this.app.ID;
    let brand = await this.createItem(ENTITIES_LIST, b);
    const BUGroup = await this.createGroup('OU-'+brand.ID);
    const BOGroup = await this.createGroup('OO-'+brand.ID);
    if (!BUGroup || ! BOGroup) throw new Error("Error creating permission groups. Please contact the domain administrator.");

    await this.addUserToGroup(owner, BOGroup.Id);
    await this.addUserToGroup(owner, BUGroup.Id);

    //create geographies
    await this.createGeographies(brand.ID,geographies, countries);

    //create models folders
    let folders = await this.createInternalFolders(brand, []); // TODO pass groups

    let permissions = await this.getGroupPermissions(FOLDER_WIP);
    let groups: SPGroupListItem[] = [];
    groups.push({ type: 'OU', data: BUGroup });
    groups.push({ type: 'OO', data: BOGroup });

    for (const f of folders.rw) {
      let folderGroups = [...groups]; // copy default groups
      let GUGroup;
      if (f.GeographyID) {
        GUGroup = await this.createGroup(
          `OU-${brand.ID}-${f.GeographyID}`, 
          'Geography ID ' + f.GeographyID);
        if (GUGroup) {
          folderGroups.push( { type: 'GU', data: GUGroup} );
          await this.addUserToGroup(owner, GUGroup.Id);
        }
      } else if(f.DepartmentID) {
        let DUGroup = await this.createGroup(`DU-${brand.ID}-${f.DepartmentID}`, 'Department ID ' + f.DepartmentID);
        if (DUGroup) {
          folderGroups.push({ type: 'DU', data: DUGroup });
          await this.addUserToGroup(owner, DUGroup.Id);
        }
      }

      await this.setPermissions(permissions, folderGroups, f.ServerRelativeUrl);
    }

    permissions = (await this.getGroupPermissions()).filter(el => el.ListFilter === 'List');
    for (const f of folders.ro) {
      let folderGroups = [...groups]; // copy default groups
      let GUGroup;
      if (f.GeographyID) {
        GUGroup = await this.createGroup(
          `OU-${brand.ID}-${f.GeographyID}`, 
          'Geography ID ' + f.GeographyID);
        let DUGroup = await this.createGroup(
            `DU-${brand.ID}-0-${f.GeographyID}`, 
            'Geography ID ' + f.GeographyID);
        if (GUGroup && DUGroup) {
          folderGroups.push( { type: 'GU', data: GUGroup} );
          folderGroups.push( { type: 'DU', data: DUGroup} );
          await this.addUserToGroup(owner, GUGroup.Id);
          await this.addUserToGroup(owner, DUGroup.Id);
        }
      } 

      await this.setPermissions(permissions, folderGroups, f.ServerRelativeUrl);
    }
    
    return brand; 
  }

  private async createInternalFolders(entity: Opportunity, groups: SPGroupListItem[], geographies?: EntityGeography[]): Promise<{rw: SystemFolder[], ro: SystemFolder[]}> {
    let ReadWriteNames = [FOLDER_WIP, FOLDER_DOCUMENTS];
    let ReadOnlyNames = [FOLDER_APPROVED, FOLDER_ARCHIVED];

    const OUGroup = groups.find(el => el.type == "OU");
    if (!OUGroup) throw new Error("Error creating group permissions for internal folders.");
    
    if(!geographies) {
      geographies = await this.getEntityGeographies(entity.ID);
    }

    let rwFolders: SystemFolder[] = [];
    for (const mf of ReadWriteNames) {
      const mfFolder = await this.createFolder(`${mf}`, true);
      if(mfFolder) {
        const BUFolder = await this.createFolder(`${mf}/${entity.BusinessUnitId}`, true);
        if(BUFolder) {
          const folder = await this.createFolder(`${mf}/${entity.BusinessUnitId}/${entity.ID}`, true);
          if (folder) {
            await this.addRolePermissionToFolder(folder.ServerRelativeUrl, OUGroup.data.Id, 'ListRead');
            const emptyStageFolder = await this.createFolder(`${mf}/${entity.BusinessUnitId}/${entity.ID}/0`, true);
            if(emptyStageFolder) {
              await this.addRolePermissionToFolder(emptyStageFolder.ServerRelativeUrl, OUGroup.data.Id, 'ListRead');
              if(mf != FOLDER_DOCUMENTS) {
                const forecastFolder = await this.createFolder(`${mf}/${entity.BusinessUnitId}/${entity.ID}/0/0`, true);
                if(forecastFolder) {
                  rwFolders = rwFolders.concat(await this.createEntityGeographyFolders(entity, geographies, mf));
                }
              } else {
                rwFolders = rwFolders.concat(await this.createDepartmentFolders(entity, mf));
              } 
            }
          }
        }
      }
    }

    let roFolders: SystemFolder[] = [];
    for (const mf of ReadOnlyNames) {
      const mfFolder = await this.createFolder(`${mf}`, true);
      if(mfFolder) {
        const BUFolder = await this.createFolder(`${mf}/${entity.BusinessUnitId}`, true);
        if(BUFolder) {
          const folder = await this.createFolder(`${mf}/${entity.BusinessUnitId}/${entity.ID}`, true);
          if (folder) {
            await this.addRolePermissionToFolder(folder.ServerRelativeUrl, OUGroup.data.Id, 'ListRead');
            const emptyStageFolder = await this.createFolder(`${mf}/${entity.BusinessUnitId}/${entity.ID}/0`, true);
            if(emptyStageFolder) {
              await this.addRolePermissionToFolder(emptyStageFolder.ServerRelativeUrl, OUGroup.data.Id, 'ListRead');
              const forecastFolder = await this.createFolder(`${mf}/${entity.BusinessUnitId}/${entity.ID}/0/0`, true);
              if(forecastFolder) {  
                roFolders = roFolders.concat(await this.createEntityGeographyFolders(entity, geographies, mf));
              }
            }
          }
        }
      }
    }
    return {
      rw: rwFolders,
      ro: roFolders
    };
  }

  private async createEntityGeographyFolders(entity: Opportunity | Brand, geographies: EntityGeography[], mf: string, departmentId: number = 0, cycleId: number = 0): Promise<SystemFolder[]> {
    let folders: SystemFolder[] = [];
    let basePath = `${mf}/${entity.BusinessUnitId}/${entity.ID}/0/${departmentId}`;
    for (const geo of geographies) {
      let geoFolder = await this.createFolder(`${basePath}/${geo.ID}`, true);
      if (geoFolder) {
        geoFolder.GeographyID = geo.ID;
        geoFolder.DepartmentID = departmentId;
        folders.push(geoFolder);
        await this.createFolder(`${basePath}/${geo.ID}/${cycleId}`, true);
      }
    }
    
    return folders;
  }
  
  private async createDepartmentFolders(entity: Brand | Opportunity, mf: string): Promise<SystemFolder[]> {
    let folders: SystemFolder[] = [];
    let basePath = `${mf}/${entity.BusinessUnitId}/${entity.ID}/0`;
    let departmentFolders = await this.getInternalDepartments();
    for(const dept of departmentFolders) {
      let folder = await this.createFolder(`${basePath}/${dept.DepartmentID}`, true);
      if (folder) {
        folder.DepartmentID = dept.DepartmentID;
        folders.push(folder);
        folder = await this.createFolder(`${basePath}/${dept.DepartmentID}/0`, true);
        if(folder) {
          folder = await this.createFolder(`${basePath}/${dept.DepartmentID}/0/0`, true);
        }
      }
    }
    return folders;
  }

  async getAllEntities(appId: number) {
    let countCond = `$filter=AppTypeId eq ${appId}`;
    let max = await this.countItems(ENTITIES_LIST, countCond);

    let cond = countCond+"&$select=*,Indication/Title,Indication/TherapyArea,EntityOwner/Title,ForecastCycle/Title,BusinessUnit/Title&$expand=EntityOwner,ForecastCycle,BusinessUnit,Indication&$skiptoken=Paged=TRUE&$top="+max;
    
    let results = await this.getAllItems(ENTITIES_LIST, cond);
    
    return results;

  }

  async getBrand(id: number): Promise<Brand> {
    let cond = "&$select=*,Indication/Title,Indication/ID,Indication/TherapyArea,EntityOwner/Title,ForecastCycle/Title,BusinessUnit/Title&$expand=EntityOwner,ForecastCycle,BusinessUnit,Indication";
   
    let results = await this.getOneItem(ENTITIES_LIST, "$filter=Id eq "+id+cond);
    
    return results;
  }

  async getBrandFields() {
    return [
      { value: 'Title', label: 'Brand Name' },
      //{ value: 'FCDueDate', label: 'Forecast Cycle Due Date' },
      { value: 'BusinessUnit.Title', label: 'Business Unit' },
      { value: 'Indication.Title', label: 'Indication Name' },
    ];
  }


  async getBrandModelsCount(brand: Brand) {
    return await this.getBrandFolderFilesCount(brand, FOLDER_WIP);
  }

  async getBrandApprovedModelsCount(brand: Brand) {
    return await this.getBrandFolderFilesCount(brand, FOLDER_APPROVED);
  }

  async getBrandFolderFilesCount(brand: Brand, folder: string) {
    let currentFolder = folder+'/'+brand.BusinessUnitId+'/'+brand.ID+'/0/0';
    const geoFolders = await this.getSubfolders(currentFolder);
    let currentFiles = [];
    for (const geofolder of geoFolders) {
      let folder = currentFolder + '/' + geofolder.Name+'/0';
      currentFiles.push(...await this.readEntityFolderFiles(folder, true));
    }
    return currentFiles.length;
  }
/*
  //return all geographies for now
  async getBrandAccessibleGeographiesList(brand: Brand): Promise<SelectInputList[]> {
    const geographiesList = await this.getBrandGeographies(brand.ID);

    const geoFoldersWithAccess = await this.getSubfolders(`${FOLDER_WIP}/${brand.BusinessUnitId}/${brand.ID}/${FORECAST_MODELS_FOLDER_NAME}`);
    return geographiesList.filter(mf => geoFoldersWithAccess.some((gf: any) => +gf.Name === mf.Id))
      .map(t => { return { value: t.Id, label: t.Title } });
  }
*/
  async getEntityAccessibleGeographiesList(entity: Opportunity | Brand): Promise<SelectInputList[]> {
    const geographiesList = await this.getEntityGeographies(entity.ID);

    const geoFoldersWithAccess = await this.getSubfolders(`${FOLDER_WIP}/${entity.BusinessUnitId}/${entity.ID}/0/0`, true);
    return geographiesList.filter(mf => geoFoldersWithAccess.some((gf: any) => +gf.Name === mf.Id))
      .map(t => { return { value: t.Id, label: t.Title } });
  }

  async getBusinessUnitsList(): Promise<SelectInputList[]> {
    let cache = this.masterBusinessUnits;
    if (cache && cache.length) {
      return cache;
    }
    let max = await this.countItems(BUSINESS_UNIT_LIST);
    let cond = "$skiptoken=Paged=TRUE&$top="+max;
    let results = await this.getAllItems(BUSINESS_UNIT_LIST, cond);
    this.masterBusinessUnits = results.map(el => { return {value: el.ID, label: el.Title }});
    return this.masterBusinessUnits;
  }

  async getForecastCycles(): Promise<SelectInputList[]> {
    let cache = this.masterForecastCycles;
    if (cache && cache.length) {
      return cache;
    }
    let max = await this.countItems(FORECAST_CYCLES_LIST);
    let cond = "$skiptoken=Paged=TRUE&$top="+max;
    let results = await this.getAllItems(FORECAST_CYCLES_LIST, cond);
    this.masterForecastCycles = results.map(el => { return {value: el.ID, label: el.Title }});
    return this.masterForecastCycles;
  }
/*
  async getBrandGeographies(brandId: number, all?: boolean) {
    let filter = `$filter=BrandId eq ${brandId}`;
    if (!all) {
      filter += ' and Removed ne 1';
    }
    return await this.getAllItems(
      GEOGRAPHIES_LIST, filter,
    );
  }*/

  async getEntityGeographies(entityId: number, all?: boolean) {
    let filter = `$filter=EntityNameId eq ${entityId}`;
    if (!all) {
      filter += ' and Removed ne 1';
    }
    return await this.getAllItems(
      GEOGRAPHIES_LIST, filter,
    );
  }

  async updateBrand(brandId: number, brandData: BrandInput): Promise<boolean> {
    const oppBeforeChanges: Brand = await this.getOneItemById(brandId, ENTITIES_LIST);
    const success = await this.updateItem(brandId, ENTITIES_LIST, brandData);

    if (success && oppBeforeChanges.EntityOwnerId !== brandData.EntityOwnerId) { // owner changed
      return this.changeEntityOwnerPermissions(brandId, oppBeforeChanges.EntityOwnerId, brandData.EntityOwnerId);
    }

    return success;
  }
  
  async getEntityFileInfo(folder: string, file: NPPFile): Promise<NPPFile> {
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
    
    return await this.query(
      `lists/getbytitle('${rootFolder}')` + `/items(${file.ListItemAllFields?.ID})`,
      select,
      'all'
    ).toPromise();
  }

  clearFileName(name: string): string {
    return name.replace(/[~#%&*{}:<>?+|"'/\\]/g, "");
  }

  async getEntityForecastCycles(entity: Brand | Opportunity) {
    let filter = `$filter=EntityNameId eq ${entity.ID}`;
    
    return await this.getAllItems(
      OPPORTUNITY_FORECAST_CYCLE_LIST, filter,
    ); 
  }

  async createEntityForecastCycle(entity: Opportunity, values: any) {
    const geographies = await this.getEntityGeographies(entity.ID); // 1 = stage id would be dynamic in the future
    let archivedBasePath = `${FOLDER_ARCHIVED}/${entity.BusinessUnitId}/${entity.ID}/0/0`;
    let approvedBasePath = `${FOLDER_APPROVED}/${entity.BusinessUnitId}/${entity.ID}/0/0`;
    let workInProgressBasePath = `${FOLDER_WIP}/${entity.BusinessUnitId}/${entity.ID}/0/0`;

    let cycle = await this.createItem(OPPORTUNITY_FORECAST_CYCLE_LIST, {
      EntityNameId: entity.ID,
      ForecastCycleTypeId: entity.ForecastCycleId,
      Year: entity.Year+"",
      Title: entity.ForecastCycle?.Title + ' ' + entity.Year,
      ForecastCycleDescriptor: entity.ForecastCycleDescriptor
    });

    for (const geo of geographies) {
      let geoFolder = `${archivedBasePath}/${geo.ID}/${cycle.ID}`;
      const cycleFolder = await this.createFolder(geoFolder, true);
      if(cycleFolder) {
        await this.moveAllFolderFiles(`${approvedBasePath}/${geo.ID}/0`, geoFolder);
      }else {
        throw new Error("Could not create Forecast Cycle folder");
      }
    }

    let changes = {
      ForecastCycleId: values.ForecastCycle,
      ForecastCycleDescriptor: values.ForecastCycleDescriptor,
      Year: values.Year
    };

    await this.updateItem(entity.ID, OPPORTUNITIES_LIST, changes);

    await this.setAllEntityModelsStatusInFolder(entity, workInProgressBasePath, "In Progress");

    return changes;

  }


  async setAllEntityModelsStatusInFolder(entity: Opportunity | Brand, folder: string, status: string) {
    
    const geographies = await this.getEntityGeographies(entity.ID); // 1 = stage id would be dynamic in the future
    
    let arrFolder = folder.split("/");
    let rootFolder = arrFolder[0];

    for(let i=0; i<geographies.length; i++) {
      let geo = geographies[i];
      let files = await this.readEntityFolderFiles(folder+"/"+geo.ID+"/0", true);
      for(let j=0; files && j<files.length; j++) {
        let model = files[j];
        await this.setEntityApprovalStatus(rootFolder, model, entity, "In Progress");
      }
    }
    
  }

  async moveAllFolderFiles(origin: string, dest: string, moveCSVs: boolean = true) {
    let files = await this.readEntityFolderFiles(origin);
    for(let i=0;i<files.length; i++){
      let model = files[i];
      let path = await this.moveFile(model.ServerRelativeUrl, dest);
      if(moveCSVs) {
        await this.moveCSV(model, path);
      }
    }
  }

  async copyAllFolderFiles(origin: string, dest: string, copyCSVs: boolean = true) {
    let files = await this.readEntityFolderFiles(origin);
    for(let i=0;i<files.length; i++){
      let model = files[i];
      let path = await this.copyFile(model.ServerRelativeUrl, dest, model.Name);
      if(copyCSVs) {
        let arrUrl = model.ServerRelativeUrl.split("/"); // server relative url base for path
        await this.copyCSV(model, "/"+arrUrl[1]+"/"+arrUrl[2]+"/"+path);
      }
    }
  }

  /** Updates a read only field fieldname of the list's element with the value */
  private async updateReadOnlyField(list: string, elementId: number, fieldname: string, value: string) {

    await this.http.post(
      this.licensing.getSharepointApiUri() + `lists/getByTitle('${list}')/items(${elementId})/validateUpdateListItem`,
      JSON.stringify({
        "formValues": [
          {
            "__metadata": { "type": "SP.ListItemFormUpdateValue" },
            "FieldName": fieldname,
            "FieldValue": "[{'Key':'" + value + "'}]"
          }
        ],
        "bNewDocumentUpdate": false
      }),
      {
        headers: new HttpHeaders({
          "Accept": "application/json; odata=verbose",
          "Content-Type": "application/json; odata=verbose"
        })
      }).toPromise();
  }

  async addComment(file: NPPFile, str: string) {
    let comments = file.ListItemAllFields?.Comments?.replace(/""/g, '"');
    let parsedComments = [];
    let commentsStr = "";
    if(comments) {
      try {
        parsedComments = JSON.parse(comments);
      } catch(e) {

      }
      let currentUser = await this.getCurrentUserInfo();
      let newComment = {
        text: str,
        email: currentUser.Email,
        name: currentUser.Title?.indexOf("@") == -1 ? currentUser.Title : currentUser.Email,
        userId: currentUser.Id,
        createdAt: new Date().toISOString()
      }
      parsedComments.push(newComment);
      commentsStr = JSON.stringify(parsedComments)
      if(file.ListItemAllFields) file.ListItemAllFields.Comments = commentsStr;
    }
    return commentsStr;   
  }

  async copyFile(originServerRelativeUrl: string, destinationFolder: string, newFileName: string): Promise<any> {
    const originUrl = `getfilebyserverrelativeurl('${originServerRelativeUrl}')/`;
    let path = destinationFolder + this.clearFileName(newFileName);
    let destinationUrl = `copyTo('${path}')`;
    try {
      const r = await this.http.post(
        this.licensing.getSharepointApiUri() + originUrl + destinationUrl,
        null
      ).toPromise();
      return path;
    }
    catch (e) {
      return false;
    }
  }

  async moveFile(originServerRelativeUrl: string, destinationFolder: string, newFilename: string = ''): Promise<any> {
    let arrUrl = originServerRelativeUrl.split("/");
    let fileName = arrUrl[arrUrl.length - 1];
    const originUrl = `getfilebyserverrelativeurl('${originServerRelativeUrl}')/`;
    let path = destinationFolder + "/" + (newFilename ? newFilename : fileName);
    let destinationUrl = `moveTo('${path}')`;
    const r = await this.http.post(
      this.licensing.getSharepointApiUri() + originUrl + destinationUrl,
      null
    ).toPromise();

    return "/"+arrUrl[1]+"/"+arrUrl[2]+"/"+path;
  }

  async setBrandApprovalStatus(rootFolder: string, file: NPPFile, brand: Brand | null, status: string, comments: string | null = null) {
    if(file.ListItemAllFields) {
      const statusId = await this.getApprovalStatusId(status);
      if (!statusId) return false;
      /*TODO use something like this to ensure unique name
      while (await this.sharepoint.existsFile(fileName, destinationFolder) && ++attemps < 11) {
        fileName = baseFileName + '-copy-' + attemps + '.' + extension;
      }*/
      let data = { ApprovalStatusId: statusId };
      if (comments) Object.assign(data, { Comments: comments });
  
      await this.updateItem(file.ListItemAllFields.ID, `lists/getbytitle('${rootFolder}')`, data);
      let res;
      if(status === "Approved" && brand) {
        let arrFolder = file.ServerRelativeUrl.split("/");
        await this.removeOldAcceptedModel(brand, file);
        res = await this.copyFile(file.ServerRelativeUrl, '/'+arrFolder[1]+'/'+arrFolder[2]+'/'+FOLDER_APPROVED+'/'+brand.BusinessUnitId+'/'+brand.ID+'/0/0/'+arrFolder[arrFolder.length - 3]+'/0/', file.Name);
        return res;
      };
      
      return true;
    } else {
      throw new Error("Missing file metadata.");
    }
  }

  async updateFileFields(path: string, fields: any) {
    this.http.post(
      this.licensing.getSharepointApiUri() + `GetFileByServerRelativeUrl('${path}')/ListItemAllFields`,
      fields,
      {
        headers: new HttpHeaders({
          'If-Match': '*',
          'X-HTTP-Method': "MERGE"
        }),
      }
    ).toPromise();
  }


  getPowerBICSVRootPathFromModelPath(path: string) {
    let mappings: any = {}
    mappings[FOLDER_DOCUMENTS] =  FOLDER_POWER_BI_DOCUMENTS,
    mappings[FOLDER_WIP] =  FOLDER_POWER_BI_WIP,
    mappings[FOLDER_APPROVED] =  FOLDER_POWER_BI_APPROVED,
    mappings[FOLDER_ARCHIVED] =  FOLDER_POWER_BI_ARCHIVED
    

    for (const [key, value] of Object.entries(mappings)) {
      if(path.indexOf(key) !== -1) {
        return value;
      }
    }

    return false;

  }

  async getModelCSVFiles(file: NPPFile) {
    let powerBiLibrary = this.getPowerBICSVRootPathFromModelPath(file.ServerRelativeUrl);
    let files: NPPFile[] = []

    if (powerBiLibrary && file.ListItemAllFields) {
      
      const result = await this.query(
        `GetFolderByServerRelativeUrl('${powerBiLibrary}')/Files`,
        '$expand=ListItemAllFields&$filter=ListItemAllFields/ForecastId eq '+file.ListItemAllFields.ID,
      ).toPromise();
  
      if (result.value) {
        files = result.value;
      }   
    }

    return files;
  }

  async deleteRelatedCSV(url: string) {
    let metadata: NPPFileMetadata = (await this.http.get(
    this.licensing.getSharepointApiUri() + `GetFileByServerRelativeUrl('${url}')/ListItemAllFields`).toPromise()) as NPPFileMetadata;
    let csvFiles = await this.getModelCSVFiles({ ServerRelativeUrl: url, ListItemAllFields: metadata } as NPPFile);
    for(let i = 0; i < csvFiles.length; i++) {
      this.deleteFile(csvFiles[i].ServerRelativeUrl, false);
    } 
  }

  async copyCSV(file: NPPFile, path: string) {
    if (file.ListItemAllFields) {
      let arrFolder = file.ServerRelativeUrl.split("/");
      let destLibrary = this.getPowerBICSVRootPathFromModelPath(path);
  
      let csvFiles = await this.getModelCSVFiles(file);
      let destModel: NPPFileMetadata = (await this.http.get(
        this.licensing.getSharepointApiUri() + `GetFileByServerRelativeUrl('${path}')/ListItemAllFields`).toPromise()) as NPPFileMetadata;
  
      for(let i = 0; i < csvFiles.length; i++) {
        let tmpFile = csvFiles[i];
        let newFileName = tmpFile.Name.replace('_'+file.ListItemAllFields.ID+'.', '_'+destModel.ID+'.');
        let newPath = '/'+arrFolder[1]+'/'+arrFolder[2]+'/'+destLibrary+'/';
        await this.copyFile(tmpFile.ServerRelativeUrl, newPath, newFileName);
        await this.updateFileFields(newPath+newFileName, {ForecastId: destModel.ID});
      } 
    }
  }

  async moveCSV(file: NPPFile, path: string) {
    if (file.ListItemAllFields) {
      let arrFolder = file.ServerRelativeUrl.split("/");
      let destLibrary = this.getPowerBICSVRootPathFromModelPath(path);
  
      let csvFiles = await this.getModelCSVFiles(file);
      let destModel: NPPFileMetadata = (await this.http.get(
        this.licensing.getSharepointApiUri() + `GetFileByServerRelativeUrl('${path}')/ListItemAllFields`).toPromise()) as NPPFileMetadata;
  
      for(let i = 0; i < csvFiles.length; i++) {
        let tmpFile = csvFiles[i];
        let newFileName = tmpFile.Name.replace('_'+file.ListItemAllFields.ID+'.', '_'+destModel.ID+'.');
        let newPath = destLibrary+'';
        await this.moveFile(tmpFile.ServerRelativeUrl, newPath, newFileName);
        await this.updateFileFields("/"+arrFolder[1]+"/"+arrFolder[2]+"/"+newPath+"/"+newFileName, {ForecastId: destModel.ID});
      } 
    }
  }

  async setEntityApprovalStatus(rootFolder: string, file: NPPFile, entity: Brand | Opportunity | null, status: string, comments: string | null = null) {
    if(file.ListItemAllFields) {
      const statusId = await this.getApprovalStatusId(status);
      if (!statusId) return false;
      /*TODO use something like this to ensure unique name
      while (await this.sharepoint.existsFile(fileName, destinationFolder) && ++attemps < 11) {
        fileName = baseFileName + '-copy-' + attemps + '.' + extension;
      }*/
      let data = { ApprovalStatusId: statusId };
      if (comments) Object.assign(data, { Comments: comments });
  
      await this.updateItem(file.ListItemAllFields.ID, `lists/getbytitle('${rootFolder}')`, data);
      let res;
      if(status === "Approved" && entity && file.ServerRelativeUrl.indexOf(FILES_FOLDER) == -1) {
        let arrFolder = file.ServerRelativeUrl.split("/");
        await this.removeNPPOldAcceptedModel(entity, file);
        res = await this.copyFile(file.ServerRelativeUrl, '/'+arrFolder[1]+'/'+arrFolder[2]+'/'+FOLDER_APPROVED+'/'+entity.BusinessUnitId+'/'+entity.ID+'/0/0/'+arrFolder[arrFolder.length - 3]+'/0/', file.Name);

        if (res) {
          await this.updateFileFields(res, {OriginalModelId: file.ListItemAllFields.ID});
          await this.copyCSV(file, res);
        }
        return res;
      };
      
      return true;
    } else {
      throw new Error("Missing file metadata.");
    }
  }

  async removeOldAcceptedModel(brand: Brand, file: NPPFile) {
    if(file.ListItemAllFields && file.ListItemAllFields.ModelScenarioId) {
      let arrFolder = file.ServerRelativeUrl.split("/");
      let path = '/'+arrFolder[1]+'/'+arrFolder[2]+'/'+FOLDER_APPROVED+'/'+brand.BusinessUnitId+'/'+brand.ID+'/0/0/'+arrFolder[arrFolder.length - 3]+'/0/';
      let scenarios = file.ListItemAllFields.ModelScenarioId;

      let model = await this.getFileByScenarios(path, scenarios);
      if(model) {
        await this.deleteFile(model.ServerRelativeUrl);
      }
    }
  }

  async removeNPPOldAcceptedModel(entity: Opportunity | Brand, file: NPPFile) {
    if(file.ListItemAllFields && file.ListItemAllFields.ModelScenarioId) {
      let arrFolder = file.ServerRelativeUrl.split("/");
      let path = '/'+arrFolder[1]+'/'+arrFolder[2]+'/'+FOLDER_APPROVED+'/'+entity.BusinessUnitId+'/'+entity.ID+'/0/0/'+arrFolder[arrFolder.length - 3]+'/0/';
      let scenarios = file.ListItemAllFields.ModelScenarioId;

      let model = await this.getFileByScenarios(path, scenarios);
      if(model) {
        await this.deleteFile(model.ServerRelativeUrl);
      }
    }
  }

  async getFileByScenarios(path: string, scenarios: number[]) {
    let files = await this.readEntityFolderFiles(path,false);
    for(let i=0;i<files.length; i++){
      let model = files[i];
      let sameScenario = this.sameScenarios(model, scenarios);
      if(sameScenario) {
        return model;
      }
    }
    return null;
  }

  sameScenarios(model: NPPFile, scenarios: number[]) {
    if(model.ListItemAllFields && model.ListItemAllFields.ModelScenarioId) {
      
      let sameScenario = model.ListItemAllFields.ModelScenarioId.length === scenarios.length;
      
      for(let j=0; sameScenario && j < model.ListItemAllFields.ModelScenarioId.length ; j++) {
        let scenarioId = model.ListItemAllFields.ModelScenarioId[j];
        sameScenario = sameScenario && (scenarios.indexOf(scenarioId) != -1);
      }
      
      return sameScenario;

    } else return false;
  }

  /** Copy files of one external opportunity to an internal one */
  async copyFilesExternalToInternal(extOppId: number, intOppId: number) {
    const externalEntity = await this.getOpportunity(extOppId);
    const internalEntity = await this.getOpportunity(intOppId);

    // copy models
    // [TODO] search for last stage number (now 3, but could change?)
    const externalModelsFolder = FILES_FOLDER + `/${externalEntity.BusinessUnitId}/${externalEntity.ID}/3/0`;
    const internalModelsFolder = FOLDER_WIP + `/${internalEntity.BusinessUnitId}/${internalEntity.ID}/0/0`;
    const externalGeographies = await this.getEntityGeographies(externalEntity.ID);
    const internalGeographies = await this.getEntityGeographies(internalEntity.ID);
    for (const extGeo of externalGeographies) {
      const intGeo = internalGeographies.find((g: EntityGeography) => {
        if (g.EntityGeographyType == 'Geography') return extGeo.GeographyId === g.GeographyId;
        else if (g.EntityGeographyType == 'Country') return extGeo.CountryId === g.CountryId;
        else return false;
      });

      if (intGeo) {
        await this.copyAllFolderFiles(`${externalModelsFolder}/${extGeo.Id}/0/`, `${internalModelsFolder}/${intGeo.Id}/0/`);
      }
    }
  }

} 
