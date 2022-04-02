import { Injectable } from '@angular/core';
import { EntityGeography, Indication, Opportunity, OpportunityType, Stage } from '@shared/models/entity';
import { SharepointService } from './sharepoint.service';
import { InlineNppDisambiguationService } from './inline-npp-disambiguation.service';
import { ENTITIES_LIST_NAME, ENTITY_STAGES_LIST_NAME, GEOGRAPHIES_LIST_NAME, MASTER_STAGES_LIST_NAME } from '@shared/sharepoint/list-names';
import { AppDataService } from './app-data.service';
import { SystemFolder } from '@shared/models/file-system';

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
  StageReview: Date;
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
  AppTypeId: number;
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
export class EntitiesService {

  constructor(
    private readonly appService: InlineNppDisambiguationService,
    private readonly appData: AppDataService
  ) { }

  



  

  
}
