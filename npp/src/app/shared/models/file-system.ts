import { EntityGeography, Indication } from "./entity";
import { User } from "./user";

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
    ApprovalDate?: Date;
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