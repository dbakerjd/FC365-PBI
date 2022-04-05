import { Injectable } from '@angular/core';
import { EntityGeography, Indication, Opportunity, OpportunityType, Stage } from '@shared/models/entity';
import { SharepointService } from './sharepoint.service';
import { InlineNppDisambiguationService } from './inline-npp-disambiguation.service';
import { ENTITIES_LIST_NAME, ENTITY_STAGES_LIST_NAME, GEOGRAPHIES_LIST_NAME, MASTER_STAGES_LIST_NAME } from '@shared/sharepoint/list-names';
import { AppDataService } from './app-data.service';
import { SystemFolder } from '@shared/models/file-system';
import { PermissionsService } from './permissions.service';

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
    private readonly appData: AppDataService,
    private readonly permissions: PermissionsService
  ) { }

  async createOpportunity(opp: OpportunityInput, st: StageInput, stageStartNumber: number = 1):
    Promise<{ opportunity: Opportunity, stage: Stage | null } | false> {
    opp.AppTypeId = this.appData.getAppType().ID;
    if(!opp.AppTypeId) throw new Error("Could not create an Entity (no AppType assigned)");
    
    // clean fields according type
    const isInternal = await this.isInternalOpportunity(opp.OpportunityTypeId);
    if (isInternal) {
      opp.ProjectStartDate = opp.ProjectEndDate = undefined;
    } else {
      opp.Year = undefined;
    }

    const opportunity = await this.appData.createEntity(opp);
    if (!opportunity) return false;

    // get master stage info
    let stage = null;

    if(!isInternal) {
      const opportunityType = await this.appData.getOpportunityType(opp.OpportunityTypeId);
      const stageType = opportunityType?.StageType;
      if(!stageType) throw new Error("Could not determine Opportunity Type");
      const masterStage = await this.appData.getMasterStage(stageType, stageStartNumber);
  
      stage = await this.appData.createStage(
        { ...st, Title: masterStage.Title, EntityNameId: opportunity.ID, StageNameId: masterStage.ID }
      );
      if (!stage) this.appData.deleteOpportunity(opportunity.ID);
    }

    return { opportunity, stage };
  }

  async createBrand(b: BrandInput, geographies: number[], countries: number[]): Promise<Opportunity|undefined> {
    const owner = await this.appData.getUserInfo(b.EntityOwnerId);
    if (!owner.LoginName) throw new Error("Could not obtain owner's information.");
    b.AppTypeId = this.appData.getAppType().ID;
    if(!b.AppTypeId) throw new Error("Could not create an Entity (no AppType assigned)");
    let brand: Opportunity = await this.appData.createEntity(b);

    if (brand) {
      await this.permissions.createGeographies(brand.ID, geographies, countries);
      await this.permissions.initializeOpportunity(brand, null);
    }
    
    return brand; 
  }

  /*
  async updateBrand(brandId: number, brandData: BrandInput): Promise<boolean> {
    const oppBeforeChanges: Opportunity = await this.sharepoint.getOneItemById(brandId, SPLists.ENTITIES_LIST_NAME);
    const success = await this.sharepoint.updateItem(brandId, SPLists.ENTITIES_LIST_NAME, brandData);

    if (success && oppBeforeChanges.EntityOwnerId !== brandData.EntityOwnerId) { // owner changed
      return this.permissions.changeEntityOwnerPermissions(brandId, oppBeforeChanges.EntityOwnerId, brandData.EntityOwnerId);
    }

    return success;
  }
  */

  /** TOCHECK igual que update brand ? */
  async updateOpportunity(oppId: number, oppData: OpportunityInput): Promise<boolean> {
    // const oppBeforeChanges: Opportunity = await this.sharepoint.getOneItemById(oppId, SPLists.ENTITIES_LIST_NAME);
    const oppBeforeChanges = await this.appData.getOpportunity(oppId, false);
    // const success = await this.sharepoint.updateItem(oppId, SPLists.ENTITIES_LIST_NAME, oppData);
    const success = await this.appData.updateEntity(oppId, oppData);

    if (success && oppBeforeChanges.EntityOwnerId !== oppData.EntityOwnerId) { // owner changed
      return this.permissions.changeEntityOwnerPermissions(oppId, oppBeforeChanges.EntityOwnerId, oppData.EntityOwnerId);
    }

    return success;
  }

  async isInternalOpportunity(oppTypeId: number): Promise<boolean> {
    const oppType = await this.appData.getOpportunityType(oppTypeId);
    if (oppType?.IsInternal) {
      return oppType.IsInternal;
    }
    return false;
  }
  
  /** TOCHECK move to upper service? */
  async updateStageSettings(stageId: number, data: any): Promise<boolean> {
    const currentStage = await this.appData.getEntityStage(stageId);
    let success = await this.appData.updateStage(stageId, data);
    // let success = await this.sharepoint.updateItem(stageId, SPLists.ENTITY_STAGES_LIST_NAME, data);

    return success && await this.permissions.changeStageUsersPermissions(
      currentStage.EntityNameId,
      currentStage.StageNameId,
      currentStage.StageUsersId,
      data.StageUsersId
    );
  }

  
}
