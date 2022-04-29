export type OpportunityInput = {
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

export type StageInput = {
    StageUsersId?: number[];
    StageUsersMails?: string[];
    StageReview?: Date;
    Title?: string;
    EntityNameId?: number;
    StageNameId?: number;
}

export type BrandInput = {
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

type EntityGeographyType = 'Geography' | 'Country';

export type EntityGeographyInput = {
  Title: string;
  EntityNameId: number;
  GeographyId?: number;
  CountryId?: number;
  EntityGeographyType: EntityGeographyType 
}
