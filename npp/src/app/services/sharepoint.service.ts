import { NumberSymbol } from '@angular/common';
import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { AccountInfo, AuthorizationUrlRequest } from '@azure/msal-browser';
import { catchError, filter } from 'rxjs/operators';
import { ErrorService } from './error.service';
import { LicensingService } from './licensing.service';
import { TeamsService } from './teams.service';

export interface OpportunityTest {
  title: string;
  moleculeName: string;
  opportunityOwner: UserTest;
  projectStart: Date;
  projectEnd: Date;
  opportunityType: string;
  opportunityStatus: string;
  indicationName: string;
  Id: number;
  therapyArea: string;
  updated: Date;
  users?: User[];
  progress: number;
}

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
  OpportunityStatus: string;
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
  // StageUsers: string;
  StageReview: Date;
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
  FirstName: string;
  LastName: string;
  profilePicUrl: string;
}

export interface UserTest {
  id: number;
  name: string;
  email?: string;
  profilePic?: string;
}

export interface ActionTest {
  id: number,
  gateId: number;
  opportunityId: number;
  title: string;
  actionName: string;
  dueDate: Date;
  completed: boolean;
  timestamp: Date;
  targetUserId: Number;
  targetUser: UserTest;
  status?: string;
}

export interface GateTest {
  id: number;
  title: string;
  opportunityId: number;
  name: string;
  reviewedAt: Date;
  createdAt: Date;
  actions: Action[];
  folders?: NPPFolder[];
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

export interface NPPFileTest {
  id: number;
  parentId: number;
  name: string;
  updatedAt: Date;
  description: string;
  stageId: number;
  opportunityId: number;
  country: string[];
  modelScenario: string[];
  modelApprovalComments: string;
  approvalStatus: string;
  user: UserTest;
}

export interface NPPFile {
  ID: number;
  ApprovalStatusId: number;
  OpportunityNameId: number;
  StageNameId: number;
  ModelApprovalComments: string;
  AuthorId: number;
  TargetUserId: number;
  TargetUser?: User;
}

export interface NPPFolder {
  id: number;
  name: string;
  containsModels?: boolean;
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

const OPPORTUNITIES_LIST = "lists/getbytitle('Opportunities')";
const OPPORTUNITY_STAGES_LIST = "lists/getbytitle('Opportunity Stages')";
const OPPORTUNITY_ACTIONS_LIST = "lists/getbytitle('Opportunity Action List')";
const MASTER_OPPORTUNITY_TYPES_LIST = "lists/getbytitle('Master Opportunity Type List')";
const MASTER_THERAPY_AREAS_LIST = "lists/getbytitle('Master Therapy Areas')";
const MASTER_STAGES_LIST = "lists/getbytitle('Master Stage List')";
const USER_INFO_LIST = "lists/getByTitle('User Information List')";

@Injectable({
  providedIn: 'root'
})
export class SharepointService {

  folders: NPPFolder[] = [{
    id: 1,
    name: 'Finance'
  },{
    id: 2,
    name: 'Commercial'
  }, {
    id: 3,
    name: 'Technical'
  }, {
    id: 4,
    name: 'Regulatory'
  }, {
    id: 5,
    name: 'Other'
  },{
    id: 6,
    name: 'Forecast Models',
    containsModels: true
  }];

  files: NPPFileTest[] = [{
    id: 1,
    parentId: 1,
    name: 'test.pdf',
    updatedAt: new Date(),
    description: 'test description',
    stageId: 1,
    opportunityId: 1,
    country: [],
    modelScenario: [],
    modelApprovalComments: '',
    approvalStatus: '',
    user: {
      id: 1,
      name: "David Baker",
      profilePic: "/assets/profile.png"
    }
  },{
    id: 2,
    parentId: 1,
    name: 'test2.pdf',
    updatedAt: new Date(),
    description: 'Another test description',
    stageId: 1,
    opportunityId: 1,
    country: [],
    modelScenario: [],
    modelApprovalComments: '',
    approvalStatus: '',
    user: {
      id: 1,
      name: "David Baker",
      profilePic: "/assets/profile.png"
    }
  },{
    id: 3,
    parentId: 1,
    name: 'test3.pdf',
    updatedAt: new Date(),
    description: 'Yet another test description',
    stageId: 1,
    opportunityId: 1,
    country: [],
    modelScenario: [],
    modelApprovalComments: '',
    approvalStatus: '',
    user: {
      id: 1,
      name: "David Baker",
      profilePic: "/assets/profile.png"
    }
  },{
    id: 4,
    parentId: 6,
    name: 'test_model',
    updatedAt: new Date(),
    description: 'Yet another test description',
    stageId: 1,
    opportunityId: 1,
    country: ['UK', 'Spain', 'Belgium'],
    modelScenario: ['Upside', 'Downside'],
    modelApprovalComments: 'Lorem Ipsum Dolor amet and all that',
    approvalStatus: 'In Progress',
    user: {
      id: 1,
      name: "David Baker",
      profilePic: "/assets/profile.png"
    }
  },{
    id: 5,
    parentId: 6,
    name: 'test_model3',
    updatedAt: new Date(),
    description: 'Yet another test description',
    stageId: 1,
    opportunityId: 1,
    country: ['UK', 'Spain', 'Belgium'],
    modelScenario: ['Upside', 'Downside'],
    modelApprovalComments: 'Some test random comment',
    approvalStatus: 'In Progress',
    user: {
      id: 1,
      name: "David Baker",
      profilePic: "/assets/profile.png"
    }
  }];

  opportunities: OpportunityTest[] =  [{
    title: "Acquisition of Nucala for COPD",
    moleculeName: "Nucala",
    opportunityOwner: {
      id: 1,
      name: "David Baker",
      profilePic: "/assets/profile.png"
    },
    projectStart: new Date("5/1/2021"),
    projectEnd: new Date("11/1/2021"),
    opportunityType: "Acquisition",
    opportunityStatus: "Active",
    indicationName: "Chronic Obstructive Pulmonary Disease (COPD)",
    Id: 67,
    therapyArea: "Respiratory",
    updated: new Date("5/25/2021 3:04 PM"),
    progress: 79
  },{
    title: "Acquisition of Tezepelumab (Asthma)",
    moleculeName: "Tezepelumab",
    opportunityOwner: {
      id: 1,
      name: "David Baker",
      profilePic: "/assets/profile.png"
    },
    projectStart: new Date("5/1/2021"),
    projectEnd: new Date("1/1/2024"),
    opportunityType: "Acquisition",
    opportunityStatus: "Active",
    indicationName: "Asthma",
    Id: 68,
    therapyArea: "Respiratory",
    updated: new Date("5/25/2021 3:55 PM"),
    progress: 45
  },{
    title: "Development of Concizumab",
    moleculeName: "Concizumab",
    opportunityOwner: {
      id: 1,
      name: "David Baker",
      profilePic: "/assets/profile.png"
    },
    projectStart: new Date("5/1/2021"),
    projectEnd: new Date("5/1/2025"),
    opportunityType: "Product Development",
    opportunityStatus: "Archived",
    indicationName: "Hemophilia",
    Id: 69,
    therapyArea: "Haematology",
    updated: new Date("5/25/2021 4:14 PM"),
    progress: 12
  }];

  constructor(private teams: TeamsService, private http: HttpClient, private error: ErrorService, private licensing: LicensingService) { }

  async query(partial: string, conditions: string = '', count: number | 'all' = 'all'): Promise<any> {
    //TODO implement usage of count
    try {
      let endpoint = this.licensing.getSharepointUri() + partial;
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

  async updateItem(id: number, list: string, data: any): Promise<any> {
    try {
      return await this.http.post(
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
      return {};
    }
  }


  async uploadFileQuery(fileData: FormData, folder: string, filename: string) {
    try {
      let url = `GetFolderByServerRelativeUrl('${folder}')/Files/add(url='${filename}',overwrite=true)?$expand=ListItemAllFields`;
      return await this.http.post(
        this.licensing.getSharepointUri() + url, 
        fileData,
      ).toPromise();
    } catch (e) {
      if(e.status == 401) {
        // await this.teams.refreshToken(true);
      }
      return {};
    }
  }

  async uploadFile(file: any, folder: string, metadata?: any): Promise<any> {
    const formData: FormData = new FormData();
    formData.append('file', file);

    let uploaded: any = await this.uploadFileQuery(formData, folder, file.name);

    if (metadata && uploaded.ListItemAllFields?.ID/* && uploaded.ServerRelativeUrl*/) {

      // GetFileByServerRelativeUrl('/Folder Name/{file_name}')/CheckOut()
      // GetFileByServerRelativeUrl('/Folder Name/{file_name}')/CheckIn(comment='Comment',checkintype=0)

      await this.updateItem(uploaded.ListItemAllFields.ID, `lists/getbytitle('${folder}')`, metadata);
    }
    return uploaded;
  }

  /** UNUSED ¿TODEL? */
  async buildDefaultHeaders(): Promise<any> {
    if (!this.teams.token) {
      await this.teams.refreshToken();
    }
    let headersObject = new HttpHeaders({
      'Accept':'application/json;odata=verbose',
      'Authorization': 'Bearer ' + this.teams.token
    });
    return headersObject;
  }

  async createOpportunity(op: OpportunityInput, st: StageInput): Promise<Opportunity> {
    let opportunity = await this.createItem(OPPORTUNITIES_LIST, op);
    let stageType = await this.getStageType(op.OpportunityTypeId);
    let masterStage = await this.getOneItem(MASTER_STAGES_LIST, `$select=ID&$filter=(StageType eq '${stageType}') and (StageNumber eq 1)`);
    let stage = await this.createItem(OPPORTUNITY_STAGES_LIST, { ...st, OpportunityNameId: opportunity.ID, StageNameId: masterStage.ID });

    console.log(`CREATED OPPORTUNITY (ID ${opportunity.ID}) AND FIRST STAGE (ID ${stage.ID})`);
    console.log('[TOOD] Call Endpoint to create groups and actions');
    
    return opportunity;
  }

  async getOpportunities(): Promise<Opportunity[]> {
    return await this.getAllItems(OPPORTUNITIES_LIST, "$select=*,OpportunityType/Title,Indication/TherapyArea,Indication/Title,Author/FirstName,Author/LastName,Author/ID,Author/EMail&$expand=OpportunityType,Indication,Author");
  }

  async getIndications(therapy?: string): Promise<Indication[]> {
    let max = await this.countItems(MASTER_THERAPY_AREAS_LIST);
    let cond = "/items?$skiptoken=Paged=TRUE&$top="+max;
    if (therapy) {
      cond += `&$filter=TherapyArea eq '${therapy}'`;
    }
    return await this.getAllItems(MASTER_THERAPY_AREAS_LIST, cond);
  }

  async getIndicationsList(therapy?: string): Promise<SelectInputList[]> {
    let indications = await this.getIndications(therapy);

    if (therapy) {
      return indications.map(el => { return {value: el.ID, label: el.Title}})
    }
    return indications.map(el => { return {value: el.ID, label: el.Title, group: el.TherapyArea}})
  }

  async getTherapiesList(): Promise<SelectInputList[]> {
    let count = await this.countItems(MASTER_THERAPY_AREAS_LIST);
    let indications: Indication[] = await this.getAllItems(MASTER_THERAPY_AREAS_LIST, "$orderby=TherapyArea asc&$skiptoken=Paged=TRUE&$top="+count);

    return indications
      .map(v => v.TherapyArea)
      .filter((value, index, self) => self.indexOf(value) === index)
      .map(v => { return { label: v, value: v }});
  }

  async getStages(opportunityId: number): Promise<Stage[]> {
    return await this.getAllItems(OPPORTUNITY_STAGES_LIST, "$filter=OpportunityNameId eq "+opportunityId);
  }

  async getActions(opportunityId: number, stageId?: number): Promise<Action[]> {
    let filterConditions = `(OpportunityNameId eq ${opportunityId})`;
    if (stageId) filterConditions += ` and (StageNameId eq ${stageId})`;
    return await this.getAllItems(OPPORTUNITY_ACTIONS_LIST, `$select=*,TargetUser/ID,TargetUser/FirstName,TargetUser/LastName&$filter=${filterConditions}&$orderby=StageNameId%20asc&$expand=TargetUser`);
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
    return await this.getAllItems(MASTER_OPPORTUNITY_TYPES_LIST);
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

  async getFiles(id: number) {
    return this.files.filter(f => f.parentId == id);
  }

  async getUserProfilePic(userId: number): Promise<string> {
    let queryObj = await this.getOneItem(USER_INFO_LIST, `$filter=Id eq ${userId}&$select=Picture`);
    return queryObj.Picture.Url;
  }

  async getStageType(OpportunityTypeId: number): Promise<string> {
    let result = await this.getOneItem(MASTER_OPPORTUNITY_TYPES_LIST, "$filter=Id eq "+OpportunityTypeId+"&$select=StageType");
    if (result == null) {
      return '';
    }
    return result.StageType;
  }
}
