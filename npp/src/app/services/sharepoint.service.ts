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
  ID: number;
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

const OPPORTUNITIES_LIST = "lists/getbytitle('Opportunities')";
const OPPORTUNITY_STAGES_LIST = "lists/getbytitle('Opportunity Stages')";
const OPPORTUNITY_ACTIONS_LIST = "lists/getbytitle('Opportunity Action List')";
const MASTER_OPPORTUNITY_TYPES_LIST = "lists/getbytitle('Master Opportunity Type List')";
const MASTER_THERAPY_AREAS_LIST = "lists/getbytitle('Master Therapy Areas')";
const MASTER_STAGES_LIST = "lists/getbytitle('Master Stage List')";
const MASTER_FOLDER_LIST = "lists/getByTitle('Master Folder List')";
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

  async createOpportunity(op: OpportunityInput, st: StageInput): Promise<Opportunity> {
    let opportunity = await this.createItem(OPPORTUNITIES_LIST, { OpportunityStatus: "Processing", ...op });
    let stageType = await this.getStageType(op.OpportunityTypeId);
    let masterStage = await this.getOneItem(MASTER_STAGES_LIST, `$select=ID&$filter=(StageType eq '${stageType}') and (StageNumber eq 1)`);
    let stage = await this.createItem(OPPORTUNITY_STAGES_LIST, { ...st, OpportunityNameId: opportunity.ID, StageNameId: masterStage.ID });

    console.log(`CREATED OPPORTUNITY (ID ${opportunity.ID}) AND FIRST STAGE (ID ${stage.ID})`);
    console.log('[TOOD] Use Endpoint to create groups and actions');

    let result = await this.http.get(
      `https://demoazurefunction20210820114014.azurewebsites.net/api/Connect?stageID=${stage.ID}&OppID=${opportunity.ID}&siteUrl=${this.licensing.getSharepointUri()}`
    ).toPromise();
    
    console.log('result', result);
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

}
