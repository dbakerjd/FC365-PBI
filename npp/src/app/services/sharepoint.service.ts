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

export interface Gate {
  ID: number;
  Title: string;
  OpportunityNameId: number;
  StageNameId: number;
  StageReview: Date;
  Created: Date;
  actions?: Action[];
  folders?: NPPFolder[];
}

export interface GateInput {
  Title: string;
  OpportunityNameId: number;
  StageNameId: number;
  StageReview: Date;
  actions?: Action[];
  folders?: NPPFolder[];
}

export interface NPPFile {
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

  files: NPPFile[] = [{
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


  /*
"Registration changes (MA owner)","Gate 2","Acquisition of Nucala for COPD","Registration changes (MA owner)","4/5/2021","Sí","7/6/2021 9:04 AM","Marc Torruella Altadill"
"QA Audit","Gate 2","Acquisition of Nucala for COPD","QA Audit","5/5/2021","Sí","7/6/2021 9:04 AM","Marc Torruella Altadill"
"Contract signing","Gate 2","Acquisition of Nucala for COPD","Contract signing","6/4/2021","Sí","7/6/2021 9:04 AM","Marc Torruella Altadill"
"Registration changes (MA owner)","Gate 2","Acquisition of Nucala for COPD","Registration changes (MA owner)","7/4/2021","Sí","7/6/2021 9:04 AM","Marc Torruella Altadill"
"Commercial terms negotiations","Gate 1","Acquisition of Tezepelumab (Asthma)","Commercial terms negotiations","2/4/2021","No","6/8/2021 4:55 PM","Marc Torruella Altadill"
"Innovation board","Gate 1","Acquisition of Tezepelumab (Asthma)","Innovation board","3/6/2021","Sí","6/8/2021 4:55 PM","Marc Torruella Altadill"
"SMT Approval","Gate 1","Acquisition of Tezepelumab (Asthma)","SMT Approval","4/5/2021","No",,
"DD/Contract approving process","Gate 1","Acquisition of Tezepelumab (Asthma)","DD/Contract approving process","5/5/2021","No",,
"Commercial terms negotiations","Gate 1","Acquisition of Tezepelumab (Asthma)","Commercial terms negotiations","6/4/2021","No",,
"Innovation board","Gate 1","Acquisition of Tezepelumab (Asthma)","Innovation board","7/4/2021","No",,
"SMT Approval","Gate 1","Acquisition of Tezepelumab (Asthma)","SMT Approval","8/3/2021","No",,
"DD/Contract approving process","Gate 1","Acquisition of Tezepelumab (Asthma)","DD/Contract approving process","9/2/2021","No",,
"Initiation and Prototyping (incl API sourcing and decision making)","Phase 1","Development of Concizumab","Initiation and Prototyping (incl API sourcing and decision making)","2/4/2021","Sí","5/25/2021 3:59 PM","David Baker"
"Formulation optimisation","Phase 1","Development of Concizumab","Formulation optimisation","3/6/2021","Sí","5/25/2021 3:59 PM","David Baker"
"Pre-Clinical study (with Report)","Phase 1","Development of Concizumab","Pre-Clinical study (with Report)","4/5/2021","Sí","5/25/2021 3:59 PM","David Baker"
"Pilot BE (incl CTA and supplies)","Phase 1","Development of Concizumab","Pilot BE (incl CTA and supplies)","5/5/2021","Sí","5/25/2021 3:59 PM","David Baker"
"Final Business case","Phase 1","Development of Concizumab","Final Business case","6/4/2021","Sí","5/25/2021 3:59 PM","David Baker"
"Tech Transfer","Phase 2","Development of Concizumab","Tech Transfer","2/4/2021","Sí","5/25/2021 4:02 PM","David Baker"
"Stability (Regulatory batches)","Phase 2","Development of Concizumab","Stability (Regulatory batches)","3/6/2021","Sí","5/25/2021 4:02 PM","David Baker"
"Pivotal BE study (incl CTA, supplies and CSR)","Phase 2","Development of Concizumab","Pivotal BE study (incl CTA, supplies and CSR)","4/5/2021","Sí","5/25/2021 4:02 PM","David Baker"
"Phase III Clinical study (incl CTA, supplies and CSR)","Phase 3","Development of Concizumab","Phase III Clinical study (incl CTA, supplies and CSR)","2/4/2021","No",,
"Market Authorisation Submission-Approval","Phase 3","Development of Concizumab","Market Authorisation Submission-Approval","3/6/2021","No",,
"Patent expiry","Phase 3","Development of Concizumab","Patent expiry","4/5/2021","No",,
"Launch activities (including pricing/reimbursement)","Phase 3","Development of Concizumab","Launch activities (including pricing/reimbursement)","5/5/2021","No",,

  */
  constructor(private teams: TeamsService, private http: HttpClient, private error: ErrorService, private licensing: LicensingService) { }

  async query(url: string): Promise<any> {
    try {
      let lists = await this.http.get(this.licensing.siteUrl + url, { headers: await this.buildDefaultHeaders() }).toPromise();
      return lists;
    } catch (e) {
      if(e.status == 401) {
        await this.teams.refreshToken(true); 
        return await this.http.get(this.licensing.siteUrl + url, { headers: await this.buildDefaultHeaders() }).toPromise();
      }
      return {};
    }
  }

  async createHttpGate(url: string): Promise<any> {
    let newGate: GateInput = {
      Title: 'New posted gate 3',
      OpportunityNameId: 1,
      StageNameId: 3,
      StageReview: new Date('2021/01/19 12:00:00'),
    }
    try {
      let result = await this.http.post(
        this.licensing.siteUrl + url, 
        newGate,
        { 
          headers: await this.buildDefaultHeaders()
        }
      ).toPromise();
      // .pipe(
      // catchError()
    // );.pipe(
      console.log('result', result);
      return result;
    } catch (e) {
      if(e.status == 401) {
        await this.teams.refreshToken(true); 
        return await this.http.get(this.licensing.siteUrl + url, { headers: await this.buildDefaultHeaders() }).toPromise();
      }
      return {};
    }
  }

  async create(url: string, data: any): Promise<any> {
    try {
      return await this.http.post(
        this.licensing.siteUrl + url, 
        data,
        { headers: await this.buildDefaultHeaders() }
      ).toPromise();
    } catch (e) {
      if(e.status == 401) {
        await this.teams.refreshToken(true);
      }
      return {};
    }
  }

  async buildDefaultHeaders(): Promise<any> {
    if (!this.teams.token) {
      await this.teams.refreshToken();
    }
    let headersObject = new HttpHeaders({
      'Accept':'application/json;odata=verbose',
      'Authorization': 'Bearer ' + this.teams.token
    });
    console.log('headers', headersObject);
    return headersObject;
  }

  async createGate() {
    this.createHttpGate("lists/getbytitle('Opportunity Stages')/items");
  }

  async createOpportunity(op: OpportunityInput): Promise<Opportunity> {
    return await this.create("lists/getbytitle('Opportunities')/items", op);
  }

  async getOpportunities(): Promise<Opportunity[]> {
    let queryObj = await this.query("lists/getbytitle('Opportunities')/items?$select=*,OpportunityType/Title,Indication/TherapyArea,Indication/Title,Author/FirstName,Author/LastName,Author/ID,Author/EMail&$expand=OpportunityType,Indication,Author");
    console.log('qObj', queryObj);
    return queryObj.d.results;
  }

  async getIndications(therapy?: string): Promise<Indication[]> {
    let max = await this.query("lists/getbytitle('Master Therapy Areas')/ItemCount");
    let queryUrl = "lists/getbytitle('Master Therapy Areas')/items?$skiptoken=Paged=TRUE&$top="+max.d.ItemCount;
    if (therapy) {
      queryUrl += `&$filter=TherapyArea eq '${therapy}'`;
    }
    let queryObj = await this.query(queryUrl);
    return queryObj.d.results;
  }

  async getIndicationsList(therapy?: string): Promise<SelectInputList[]> {
    let indications = await this.getIndications(therapy);

    if (therapy) {
      return indications.map(el => { return {value: el.ID, label: el.Title}})
    }
    return indications.map(el => { return {value: el.ID, label: el.Title, group: el.TherapyArea}})
  }

  async getTherapiesList(): Promise<SelectInputList[]> {
    let count = await this.query("lists/getbytitle('Master Therapy Areas')/ItemCount");
    let queryObj = await this.query("lists/getbytitle('Master Therapy Areas')/items?$orderby=TherapyArea asc&$skiptoken=Paged=TRUE&$top="+count.d.ItemCount);
    let indications: Indication[] = queryObj.d.results;

    return indications
      .map(v => v.TherapyArea)
      .filter((value, index, self) => self.indexOf(value) === index)
      .map(v => { return { label: v, value: v }});
  }

  async getGates(opportunityId: number): Promise<Gate[]> {
    let queryObj = await this.query("lists/getbytitle('Opportunity Stages')/items?$filter=OpportunityNameId eq "+opportunityId);
    console.log('qObjGates', queryObj);
    return queryObj.d.results;
    // return this.gates.filter(el => el.opportunityId == opportunityId);
  }

  async getActions(opportunityId: number, stageId?: number): Promise<Action[]> {
    let filterConditions = `(OpportunityNameId eq ${opportunityId})`;
    if (stageId) filterConditions += ` and (StageNameId eq ${stageId})`;
    let queryObj = await this.query(`lists/getbytitle('Opportunity Action List')/items?$select=*,TargetUser/ID,TargetUser/FirstName,TargetUser/LastName&$filter=${filterConditions}&$orderby=StageNameId%20asc&$expand=TargetUser`);
    console.log('qObjActions', queryObj);
    return queryObj.d.results;
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
  }

  async getOpportunityTypes(): Promise<OpportunityType[]> {
    let queryObj = await this.query("lists/getbytitle('Master Opportunity Type List')/items");
    console.log('qObjOpTypes', queryObj);
    return queryObj.d.results;
  }

  async getOpportunityTypesList() {
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
    let queryObj = await this.query("lists/getbytitle('Opportunities')/items?$filter=Id eq "+id+"&$select=*,OpportunityType/Title,Indication/TherapyArea,Indication/Title,Author/FirstName,Author/LastName,Author/ID,Author/EMail,OpportunityOwner/ID,OpportunityOwner/FirstName,OpportunityOwner/EMail,OpportunityOwner/LastName&$expand=OpportunityType,Indication,Author,OpportunityOwner");
    console.log('objSingleOpportunity', queryObj);
    return queryObj.d.results[0];
  }

  async getFiles(id: number) {
    return this.files.filter(f => f.parentId == id);
  }

  async getUserProfilePic(userId: number): Promise<string> {
    let queryObj = await this.query(`lists/getByTitle('User Information List')/items?$filter=Id eq ${userId}&$select=Picture`);
    return queryObj.d.results[0].Picture.Url;
  }
}
