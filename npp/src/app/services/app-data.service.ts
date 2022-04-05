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
import { GraphService } from './graph.service';
import { InlineNppDisambiguationService } from './inline-npp-disambiguation.service';
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
  Title: string;
  EntityOwnerId: number;
  IndicationId: number;
  BusinessUnitId: number;
  ForecastCycleId: number;
  FCDueDate?: Date;
  Year: number;
  AppTypeId?: number;
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

  constructor(private readonly sharepoint: SharepointService, private readonly msgraph: GraphService,
    private readonly licensing: LicensingService,
    private readonly toastr: ToastrService) { }

  async canConnectAndAccessData(): Promise<boolean> {
    try {
      const currentUser = await this.getCurrentUserInfo();
      const userInfo = await this.getUserInfo(currentUser.Id);
      return true;
    } catch (e) {
      return false;
    }
  }

  async getOpportunity(id: number, expand = true): Promise<Opportunity> {
    let options = "$filter=Id eq " + id;
    if (expand) {
      options += "&$select=*,ClinicalTrialPhase/Title,ForecastCycle/Title,BusinessUnit/Title,OpportunityType/Title,Indication/TherapyArea,Indication/ID,Indication/Title,Author/FirstName,Author/LastName,Author/ID,Author/EMail,EntityOwner/ID,EntityOwner/Title,EntityOwner/FirstName,EntityOwner/EMail,EntityOwner/LastName&$expand=OpportunityType,Indication,Author,EntityOwner,BusinessUnit,ClinicalTrialPhase,ForecastCycle";
    }
    return await this.sharepoint.getOneItem(SPLists.ENTITIES_LIST_NAME, options);
  }

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

  /** --- STAGES --- **/

  /** ok */
  async createStage(data: StageInput): Promise<Stage | null> {
    if (!data.Title && data.StageNameId) {
      // get from master list
      const masterStage = await this.sharepoint.getOneItemById(data.StageNameId, SPLists.MASTER_STAGES_LIST_NAME);
      Object.assign(data, { Title: masterStage.Title });
    }
    return await this.sharepoint.createItem(SPLists.ENTITY_STAGES_LIST_NAME, data);
  }

  /** ok */
  async updateStage(id: number, data: StageInput): Promise<boolean> {
    return await this.sharepoint.updateItem(id, SPLists.ENTITY_STAGES_LIST_NAME, data);
  }

  /** ok */
  async getAllStages(): Promise<Stage[]> {
    return await this.sharepoint.getAllItems(SPLists.ENTITY_STAGES_LIST_NAME);
  }

  /** ok */
  async getEntityStages(entityId: number): Promise<Stage[]> {
    return await this.sharepoint.getAllItems(SPLists.ENTITY_STAGES_LIST_NAME, "$filter=EntityNameId eq " + entityId);
  }

  /** ok */
  async getEntityStage(id: number): Promise<Stage> {
    return await this.sharepoint.getOneItemById(id, SPLists.ENTITY_STAGES_LIST_NAME);
  }

  /** ok */
  async getFirstStage(entity: Opportunity) {
    const stageType = await this.getStageType(entity.OpportunityTypeId);
    const firstMasterStage = await this.getMasterStage(stageType, 1);
    return await this.sharepoint.getOneItem(
      SPLists.ENTITY_STAGES_LIST_NAME,
      `$filter=EntityNameId eq ${entity.ID} and StageNameId eq ${firstMasterStage.ID}`
    );
  }

  

  

  /** read app config values from sharepoint */ /* TOCHECK unused ? */
  public async getAppConfig() {
    return await this.sharepoint.getAllItems(SPLists.APP_CONFIG_LIST_NAME);
  }

  /** TOCHECK unused? */
  public async getApp(appId: string) {
    return await this.sharepoint.getAllItems(SPLists.MASTER_APPS_LIST_NAME, "$select=*&$filter=Title eq '"+appId+"'");
  }

  /** ok */
  async getUserInfo(userId: number): Promise<User> {
    return await this.sharepoint.query(`siteusers/getbyid('${userId}')`).toPromise();
  }

  /** ok */
  async getUsers(): Promise<User[]> {
    const result = await this.sharepoint.query('siteusers').toPromise();
    if (result.value) {
      return result.value;
    }
    return [];
  }

  /** ok */
  async getUserGroups(userId: number): Promise<SPGroup[]> {
    const user = await this.sharepoint.query(`siteusers/getbyid('${userId}')?$expand=groups`).toPromise();
    if (user.Groups.length > 0) {
      return user.Groups;
    }
    return [];
  }

  /** ok */
  /** Adds a user to a group */
  async addUserToGroup(user: User, groupId: number): Promise<boolean> {
    return user.LoginName ? await this.sharepoint.addUserToSharepointGroup(user.LoginName, groupId) : false;
  }

  /** ok */
  async removeUserFromGroupId(userId: number, groupId: number): Promise<boolean> {
    return await this.sharepoint.removeUserFromSharepointGroup(userId, groupId);
  }

  /** ok */
  async removeUserFromGroupName(userId: number, groupName: string): Promise<boolean> {
    return await this.sharepoint.removeUserFromSharepointGroup(userId, groupName);
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

  async removeUserSeat(user: User) {
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

  /** ---- MASTER INFO ---- */
  /** ok */
  async getMasterApprovalStatuses(): Promise<MasterApprovalStatus[]> {
    if (this.masterApprovalStatusList.length < 1) {
      this.masterApprovalStatusList = await this.sharepoint.getAllItems(SPLists.MASTER_APPROVAL_STATUS_LIST_NAME);
    }
    return this.masterApprovalStatusList;
  }

  /** ok */
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





  /** unused ? */
  async setApprovalStatus(fileId: number, status: string, comments: string | null = null, folder: string = FILES_FOLDER): Promise<boolean> {
    const statusId = await this.getMasterApprovalStatusId(status);
    if (!statusId) return false;

    let data = { ApprovalStatusId: statusId };
    if (comments) Object.assign(data, { Comments: comments });

    return await this.sharepoint.updateItem(fileId, `lists/getbytitle('${folder}')`, data);
  }


  /** TOCHECK getbrand o get Entity? */
  async getBrand(id: number): Promise<Opportunity> {
    let cond = "&$select=*,Indication/Title,Indication/ID,Indication/TherapyArea,EntityOwner/Title,ForecastCycle/Title,BusinessUnit/Title&$expand=EntityOwner,ForecastCycle,BusinessUnit,Indication";
   
    let results = await this.sharepoint.getOneItem(SPLists.ENTITIES_LIST_NAME, "$filter=Id eq "+id+cond);
    
    return results;
  }

  /** TOCHECK on ha d'anar? */
  async getBrandFields() {
    return [
      { value: 'Title', label: 'Brand Name' },
      //{ value: 'FCDueDate', label: 'Forecast Cycle Due Date' },
      { value: 'BusinessUnit.Title', label: 'Business Unit' },
      { value: 'Indication.Title', label: 'Indication Name' },
    ];
  }

  async getBrandModelsCount(brand: Opportunity) {
    return await this.getBrandFolderFilesCount(brand, FOLDER_WIP);
  }

  async getBrandApprovedModelsCount(brand: Opportunity) {
    return await this.getBrandFolderFilesCount(brand, FOLDER_APPROVED);
  }

  async removeOldAcceptedModel(brand: Opportunity, file: NPPFile) {
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

  async removeNPPOldAcceptedModel(entity: Opportunity, file: NPPFile) {
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

  async setEntityApprovalStatus(rootFolder: string, file: NPPFile, entity: Opportunity | null, status: string, comments: string | null = null) {
    if(file.ListItemAllFields) {
      const statusId = await this.getMasterApprovalStatusId(status);
      if (!statusId) return false;
      /*TODO use something like this to ensure unique name
      while (await this.sharepoint.existsFile(fileName, destinationFolder) && ++attemps < 11) {
        fileName = baseFileName + '-copy-' + attemps + '.' + extension;
      }*/
      let data = { ApprovalStatusId: statusId };
      if (comments) Object.assign(data, { Comments: comments });
  
      await this.sharepoint.updateItem(file.ListItemAllFields.ID, `lists/getbytitle('${rootFolder}')`, data);
      let res;
      if(status === "Approved" && entity && file.ServerRelativeUrl.indexOf(FILES_FOLDER) == -1) {
        let arrFolder = file.ServerRelativeUrl.split("/");
        await this.removeNPPOldAcceptedModel(entity, file);
        res = await this.copyFile(file.ServerRelativeUrl, '/'+arrFolder[1]+'/'+arrFolder[2]+'/'+FOLDER_APPROVED+'/'+entity.BusinessUnitId+'/'+entity.ID+'/0/0/'+arrFolder[arrFolder.length - 3]+'/0/', file.Name);

        if (res) {
          await this.sharepoint.updateFileFields(res, {OriginalModelId: file.ListItemAllFields.ID});
          await this.copyCSV(file, res);
        }
        return res;
      };
      
      return true;
    } else {
      throw new Error("Missing file metadata.");
    }
  }

  /** TOCHECK similud amb setentityapprovalstatus */
  async setBrandApprovalStatus(rootFolder: string, file: NPPFile, brand: Opportunity | null, status: string, comments: string | null = null) {
    if(file.ListItemAllFields) {
      const statusId = await this.getMasterApprovalStatusId(status);
      if (!statusId) return false;
      /*TODO use something like this to ensure unique name
      while (await this.sharepoint.existsFile(fileName, destinationFolder) && ++attemps < 11) {
        fileName = baseFileName + '-copy-' + attemps + '.' + extension;
      }*/
      let data = { ApprovalStatusId: statusId };
      if (comments) Object.assign(data, { Comments: comments });
  
      await this.sharepoint.updateItem(file.ListItemAllFields.ID, `lists/getbytitle('${rootFolder}')`, data);
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

  async copyCSV(file: NPPFile, path: string) {
    if (file.ListItemAllFields) {
      let arrFolder = file.ServerRelativeUrl.split("/");
      let destLibrary = this.getPowerBICSVRootPathFromModelPath(path);
  
      let csvFiles = await this.getModelCSVFiles(file);
      let destModel = await this.sharepoint.readFileMetadata(path);
  
      for(let i = 0; i < csvFiles.length; i++) {
        let tmpFile = csvFiles[i];
        let newFileName = tmpFile.Name.replace('_'+file.ListItemAllFields.ID+'.', '_'+destModel.ID+'.');
        let newPath = '/'+arrFolder[1]+'/'+arrFolder[2]+'/'+destLibrary+'/';
        await this.copyFile(tmpFile.ServerRelativeUrl, newPath, newFileName);
        await this.sharepoint.updateFileFields(newPath+newFileName, {ForecastId: destModel.ID});
      } 
    }
  }

  async moveCSV(file: NPPFile, path: string) {
    if (file.ListItemAllFields) {
      let arrFolder = file.ServerRelativeUrl.split("/");
      let destLibrary = this.getPowerBICSVRootPathFromModelPath(path);
  
      let csvFiles = await this.getModelCSVFiles(file);
      let destModel = await this.sharepoint.readFileMetadata(path);
  
      for(let i = 0; i < csvFiles.length; i++) {
        let tmpFile = csvFiles[i];
        let newFileName = tmpFile.Name.replace('_'+file.ListItemAllFields.ID+'.', '_'+destModel.ID+'.');
        let newPath = destLibrary+'';
        await this.moveFile(tmpFile.ServerRelativeUrl, newPath, newFileName);
        await this.sharepoint.updateFileFields("/"+arrFolder[1]+"/"+arrFolder[2]+"/"+newPath+"/"+newFileName, {ForecastId: destModel.ID});
      } 
    }
  }

  /** TOCHECK on ha d'anar? */
  async setActionDueDate(actionId: number, newDate: string) {
    return await this.sharepoint.updateItem(actionId, SPLists.ENTITY_ACTIONS_LIST_NAME, { ActionDueDate: newDate });
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
    
    return await this.sharepoint.query(
      `lists/getbytitle('${rootFolder}')` + `/items(${file.ListItemAllFields?.ID})`,
      select,
      'all'
    ).toPromise();
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

  /***************************** OK **********************************/

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

  /** ok */
  async getEntityGeographies(entityId: number, all?: boolean) {
    let filter = `$filter=EntityNameId eq ${entityId}`;
    if (!all) {
      filter += ' and Removed ne 1';
    }
    return await this.sharepoint.getAllItems(
       SPLists.GEOGRAPHIES_LIST_NAME, filter,
    );
  }
  
  /** ok */
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

  private clearFileName(name: string): string {
    return name.replace(/[~#%&*{}:<>?+|"'/\\]/g, "");
  }

  async getEntityForecastCycles(entity: Opportunity) {
    let filter = `$filter=EntityNameId eq ${entity.ID}`;
    
    return await this.sharepoint.getAllItems(
      SPLists.OPPORTUNITY_FORECAST_CYCLE_LIST_NAME, filter,
    ); 
  }

  async createEntityForecastCycle(entity: Opportunity, values: any) {
    const geographies = await this.getEntityGeographies(entity.ID); // 1 = stage id would be dynamic in the future
    let archivedBasePath = `${FOLDER_ARCHIVED}/${entity.BusinessUnitId}/${entity.ID}/0/0`;
    let approvedBasePath = `${FOLDER_APPROVED}/${entity.BusinessUnitId}/${entity.ID}/0/0`;
    let workInProgressBasePath = `${FOLDER_WIP}/${entity.BusinessUnitId}/${entity.ID}/0/0`;

    let cycle = await this.sharepoint.createItem(SPLists.OPPORTUNITY_FORECAST_CYCLE_LIST_NAME, {
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

    await this.sharepoint.updateItem(entity.ID, SPLists.ENTITIES_LIST_NAME, changes);

    await this.setAllEntityModelsStatusInFolder(entity, workInProgressBasePath, "In Progress");

    return changes;

  }

  async cloneEntityForecastModel(originFile: NPPFile, newFilename: string, newScenarios: number[], authorId: number, comments = ''): Promise<boolean> {

    const destinationFolder = originFile.ServerRelativeUrl.replace('/' + originFile.Name, '/');

    let success = await this.sharepoint.cloneFile(originFile.ServerRelativeUrl, destinationFolder, newFilename);
    if (!success) return false;

    let newFileInfo = await this.sharepoint.query(
      `GetFolderByServerRelativeUrl('${destinationFolder}')/Files`,
      `$expand=ListItemAllFields&$filter=Name eq '${this.clearFileName(newFilename)}'`,
    ).toPromise();

    if (newFileInfo.value[0].ListItemAllFields && originFile.ListItemAllFields) {
      const newData:any = {
        ModelScenarioId: newScenarios,
        Comments: comments ? comments : null,
        ApprovalStatusId: await this.getMasterApprovalStatusId("In Progress")
      }
      
      let arrFolder = destinationFolder.split("/");
      let rootFolder = arrFolder[3];
      
      success = await this.sharepoint.updateItem(newFileInfo.value[0].ListItemAllFields.ID, `lists/getbytitle('${rootFolder}')`, newData);
      if(success && authorId) {
        const user = await this.getUserInfo(authorId);
        if (user.LoginName)
          await this.sharepoint.updateReadOnlyField(rootFolder, newFileInfo.value[0].ListItemAllFields.ID, 'Editor', user.LoginName);
      }
    }

    return success;
  }

  async cloneForecastModel(originFile: NPPFile, newFilename: string, newScenarios: number[], comments = ''): Promise<boolean> {

    const destinationFolder = originFile.ServerRelativeUrl.replace('/' + originFile.Name, '/');

    let success = await this.sharepoint.cloneFile(originFile.ServerRelativeUrl, destinationFolder, newFilename);
    if (!success) return false;

    let newFileInfo = await this.sharepoint.query(
      `GetFolderByServerRelativeUrl('${destinationFolder}')/Files`,
      `$expand=ListItemAllFields&$filter=Name eq '${newFilename}'`,
    ).toPromise();

    if (newFileInfo.value[0].ListItemAllFields && originFile.ListItemAllFields) {
      const newData = {
        ModelScenarioId: newScenarios,
        Comments: comments ? comments : null,
        ApprovalStatusId: await this.getMasterApprovalStatusId("In Progress")
      }
      success = await this.sharepoint.updateItem(newFileInfo.value[0].ListItemAllFields.ID, `lists/getbytitle('${FILES_FOLDER}')`, newData);
    }

    return success;
  }


  async setAllEntityModelsStatusInFolder(entity: Opportunity, folder: string, status: string) {
    
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

  /** ----- USERS ----- **/

  /** ok */
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

  /** ok */
  removeCurrentUserInfo() {
    localStorage.removeItem('sharepointAccount');
  }

  /** ok */
  async getSeats(email: string) {
    return await this.licensing.getSeats(email);
  }

  /** ok */
  async addseattouser(email: string) {
    await this.licensing.addSeat(email);
  }

  /** ok */
  async removeseattouser(email: string) {
    await this.licensing.removeSeat(email);
  }

  /** --- NOTIFICATIONS --- */

  /** ok */
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

  /** ok */
  async updateNotification(notificationId: number, data: any): Promise<boolean> {
    return await this.sharepoint.updateItem(notificationId, SPLists.NOTIFICATIONS_LIST_NAME, data);
  }

  /** ok */
  async notificationsCount(userId: number, conditions = ''): Promise<number> {
    conditions = `$filter=TargetUserId eq '${userId}'` + ( conditions ? ' and ' + conditions : '');
    // item count de sharepoint ho retorna tot sense condicions => getAllItems + length
    return (await this.sharepoint.getAllItems(SPLists.NOTIFICATIONS_LIST_NAME, '$select=Id&' + conditions)).length;
  }

  /** ok */
  async createNotification(userId: number, text: string): Promise<NPPNotification> {
    return await this.sharepoint.createItem(SPLists.NOTIFICATIONS_LIST_NAME, {
      Title: text,
      TargetUserId: userId
    });
  }

  /** ---- Power BI ---- **/

  /** ok */
  async getReports(): Promise<PBIReport[]>{
    return await this.sharepoint.getAllItems(SPLists.MASTER_POWER_BI_LIST_NAME,'$orderby=SortOrder');
  }

  /** ok */
  async getReport(id:number): Promise<PBIReport>{
    return await this.sharepoint.getOneItemById(id, SPLists.MASTER_POWER_BI_LIST_NAME);
  }

  /** ok */
  async getReportByName(reportName:string): Promise<PBIReport>{
    let filter = `$filter=Title eq '${reportName}'`;
    let select = `$select=ID,name,GroupId,pageName,Title`;
    return await this.sharepoint.getOneItem(SPLists.MASTER_POWER_BI_LIST_NAME,`${select}&${filter}`)
  }

  /** ok */
  async getComponents(report: PBIReport): Promise<PBIRefreshComponent[]> {
    let select = `$select=Title,ComponentType,GroupId`
    let filter = `$filter=ReportTypeId eq'${report.ID}'`;
    let order = '$orderby=ComponentOrder';
    let reportComponents: PBIRefreshComponent[];
    return reportComponents = (await this.sharepoint.getAllItems(SPLists.MASTER_POWER_BI_COMPONENTS_LIST_NAME, `${select}&${filter}&${order}`)).map(t => { return { ComponentType: t.ComponentType, GroupId: t.GroupId, ComponentName: t.Title } })
  }

  /** ---- Files ----- **/

  /** ok */
  async readFile(fileUri: string): Promise<any> {
    return await this.sharepoint.readFile(fileUri);
  }

  /** ok */
  async deleteFile(fileUri: string, checkCSV: boolean = true): Promise<boolean> {
    //First check if it has related CSV files to remove
    if (checkCSV) {
      await this.deleteRelatedCSV(fileUri);
    }
    //then remove
    return await this.sharepoint.deleteFile(fileUri);
  }

  /** ok */
  async renameFile(fileUri: string, newName: string): Promise<boolean> {
    return await this.sharepoint.renameFile(fileUri, newName);
  }

  /** ok */
  async copyFile(originServerRelativeUrl: string, destinationFolder: string, newFileName: string): Promise<any> {
    return await this.sharepoint.copyFile(originServerRelativeUrl, destinationFolder, this.clearFileName(newFileName));
  }

  /** ok */
  async moveFile(originServerRelativeUrl: string, destinationFolder: string, newFilename: string = ''): Promise<any> {
    return await this.sharepoint.moveFile(originServerRelativeUrl, destinationFolder, newFilename);
  }

  /** ok */
  async existsFile(filename: string, folder: string): Promise<boolean> {
    return await this.sharepoint.existsFile(filename, folder);
  }

  /** TOCHECK move ? */
  async uploadFile(fileData: string, folder: string, fileName: string, metadata?: any): Promise<any> {
    let uploaded: any = await this.sharepoint.uploadFileQuery(fileData, folder, this.clearFileName(fileName));

    if (metadata && uploaded.ListItemAllFields?.ID/* && uploaded.ServerRelativeUrl*/) {

      // GetFileByServerRelativeUrl('/Folder Name/{file_name}')/CheckOut()
      // GetFileByServerRelativeUrl('/Folder Name/{file_name}')/CheckIn(comment='Comment',checkintype=0)

      await this.sharepoint.updateItem(uploaded.ListItemAllFields.ID, `lists/getbytitle('${FILES_FOLDER}')`, metadata);
    }
    return uploaded;
  }

  /** TOCHECK move ? */
  async uploadInternalFile(fileData: string, folder: string, fileName: string, metadata?: any): Promise<any> {
    if(metadata) {
      let scenarios = metadata.ModelScenarioId;
      if(scenarios) {
        let file = await this.getFileByScenarios(folder, scenarios);
        if(file) this.deleteFile(file?.ServerRelativeUrl);
      }
    }
    
    let uploaded: any = await this.sharepoint.uploadFileQuery(fileData, folder, this.clearFileName(fileName));

    if (metadata && uploaded.ListItemAllFields?.ID/* && uploaded.ServerRelativeUrl*/) {

      // GetFileByServerRelativeUrl('/Folder Name/{file_name}')/CheckOut()
      // GetFileByServerRelativeUrl('/Folder Name/{file_name}')/CheckIn(comment='Comment',checkintype=0)
      let arrFolder = folder.split("/");
      let rootFolder = arrFolder[0];
      if(!metadata.Comments) {
        metadata.Comments = " ";
      }
      await this.sharepoint.updateItem(uploaded.ListItemAllFields.ID, `lists/getbytitle('${rootFolder}')`, metadata);
    }
    return uploaded;
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

  async deleteRelatedCSV(url: string) {
    let metadata = await this.sharepoint.readFileMetadata(url);
    let csvFiles = await this.getModelCSVFiles({ ServerRelativeUrl: url, ListItemAllFields: metadata } as NPPFile);
    for(let i = 0; i < csvFiles.length; i++) {
      this.deleteFile(csvFiles[i].ServerRelativeUrl, false);
    } 
  }

  async getModelCSVFiles(file: NPPFile) {
    let powerBiLibrary = this.getPowerBICSVRootPathFromModelPath(file.ServerRelativeUrl);
    let files: NPPFile[] = []

    if (powerBiLibrary && file.ListItemAllFields) {
      
      const result = await this.sharepoint.query(
        `GetFolderByServerRelativeUrl('${powerBiLibrary}')/Files`,
        '$expand=ListItemAllFields&$filter=ListItemAllFields/ForecastId eq '+file.ListItemAllFields.ID,
      ).toPromise();
  
      if (result.value) {
        files = result.value;
      }   
    }

    return files;
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

  /** Impossible to expand ListItemAllFields/Author in one query using Sharepoint REST API */

  async readEntityFolderFiles(folder: string, expandProperties = false): Promise<NPPFile[]> {
    let files: NPPFile[] = []
    const result = await this.sharepoint.query(
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

  async addScenarioSufixToFilename(originFilename: string, scenarioId: number): Promise<string | false> {
    const scenarios = await this.getScenariosList();
    const extension = originFilename.split('.').pop();
    if (!extension) return false;

    const baseFileName = originFilename.substring(0, originFilename.length - (extension.length + 1));
    return baseFileName
      + '-' + scenarios.find(el => el.value === scenarioId)?.label.replace(/ /g, '').toLocaleLowerCase()
      + '.' + extension;
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

  async getBrandFolderFilesCount(brand: Opportunity, folder: string) {
    let currentFolder = folder+'/'+brand.BusinessUnitId+'/'+brand.ID+'/0/0';
    const geoFolders = await this.getSubfolders(currentFolder);
    let currentFiles = [];
    for (const geofolder of geoFolders) {
      let folder = currentFolder + '/' + geofolder.Name+'/0';
      currentFiles.push(...await this.readEntityFolderFiles(folder, true));
    }
    return currentFiles.length;
  }

  /** TOCHECK set type */
  async getSubfolders(folder: string, isAbsolutePath: boolean = false): Promise<any> {
    let basePath = FILES_FOLDER;
    if (isAbsolutePath) basePath = '';
    return await this.sharepoint.getPathSubfolders(basePath + folder);
  }

  /** Copy files of one external opportunity to an internal one */
  async copyFilesExternalToInternal(extOppId: number, intOppId: number) {
    const externalEntity = await this.getOpportunity(extOppId);
    const internalEntity = await this.getOpportunity(intOppId);

    // copy models
    // [TODO] search for last stage number (now 3, but could change?)
    const externalModelsFolder =  FILES_FOLDER + `/${externalEntity.BusinessUnitId}/${externalEntity.ID}/3/0`;
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
  
    
  
    
  
    
    
  
    
    
  
    
  
    
  
    
  
    
  
    
  
    
  
    
  
    
  
    

    getAppType(): AppType {
      return this.app;
    }
  
}
