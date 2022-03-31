import { AppType, Country, MasterGeography } from "src/app/services/sharepoint.service";
import { NPPFolder } from "./file-system";
import { User } from "./user";

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
    ForecastCycle?: ForecastCycle;
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

export interface OpportunityType {
    ID: number;
    Title: string;
    StageType: string;
    IsInternal: boolean;
}

export interface ClinicalTrialPhase {
    ID: number;
    Title: string;
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

  export interface Indication {
    ID: number;
    Title: string;
    TherapyArea: string;
  }
  
  export interface ForecastCycle {
    ID: number;
    Title: string;
    ForecastCycleDescriptor: string;
    SortOrder: number;
  }

  export interface EntityForecastCycle {
    ID: number;
    Title: string;
    EntityId: number;
    Entity?: Opportunity;
    ForecastCycleTypeId: number;
    ForecastCycleType?: ForecastCycle;
    ForecastCycleDescriptor: string;
    Year: string;
  }

  export interface BusinessUnit {
    ID: number;
    Title: string;
    BUOwnerID: number;
    BUOwner?: User;
    SortOrder: number;
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