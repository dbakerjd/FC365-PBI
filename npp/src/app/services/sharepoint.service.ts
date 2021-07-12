import { Injectable } from '@angular/core';

export interface Opportunity {
  title: string;
  moleculeName: string;
  opportunityOwner: string;
  projectStart: Date;
  projectEnd: Date;
  opportunityType: string;
  opportunityStatus: string;
  indicationName: string;
  Id: number;
  therapyArea: string;
  updated: Date;
}

@Injectable({
  providedIn: 'root'
})
export class SharepointService {

  opportunities: Opportunity[] = [];

  constructor() { }

  async getOpportunities(): Promise<Opportunity[]> {
    return [{
      title: "Acquisition of Nucala for COPD",
      moleculeName: "Nucala",
      opportunityOwner: "Kristian Barker",
      projectStart: new Date("5/1/2021"),
      projectEnd: new Date("11/1/2021"),
      opportunityType: "Acquisition",
      opportunityStatus: "Active",
      indicationName: "Chronic Obstructive Pulmonary Disease (COPD)",
      Id: 67,
      therapyArea: "",
      updated: new Date("5/25/2021 3:04 PM")
    },{
      title: "Acquisition of Tezepelumab (Asthma)",
      moleculeName: "Tezepelumab",
      opportunityOwner: "Kristian Barker",
      projectStart: new Date("5/1/2021"),
      projectEnd: new Date("1/1/2024"),
      opportunityType: "Acquisition",
      opportunityStatus: "Active",
      indicationName: "Asthma",
      Id: 68,
      therapyArea: "",
      updated: new Date("5/25/2021 3:55 PM")
    },{
      title: "Development of Concizumab",
      moleculeName: "Concizumab",
      opportunityOwner: "David James",
      projectStart: new Date("5/1/2021"),
      projectEnd: new Date("5/1/2025"),
      opportunityType: "Product Development",
      opportunityStatus: "Active",
      indicationName: "Hemophilia",
      Id: 69,
      therapyArea: "",
      updated: new Date("5/25/2021 4:14 PM")
    }];
  }
}
