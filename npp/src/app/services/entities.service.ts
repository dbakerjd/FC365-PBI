import { Injectable } from '@angular/core';
import { Opportunity, Stage } from '@shared/models/entity';
import { AppDataService } from './app/app-data.service';
import { PermissionsService } from './permissions.service';
import { FOLDER_APPROVED, FOLDER_ARCHIVED, FOLDER_WIP } from '@shared/sharepoint/folders';
import { FilesService } from './files.service';
import { BrandInput, OpportunityInput, StageInput } from '@shared/models/inputs';

@Injectable({
  providedIn: 'root'
})
export class EntitiesService {

  constructor(
    private readonly appData: AppDataService,
    private readonly permissions: PermissionsService,
    private readonly files: FilesService
  ) { }

  async getAll(expand = true, onlyActive = false): Promise<Opportunity[]> {
    return await this.appData.getAllOpportunities(expand, onlyActive);
  }

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
    if (!b.EntityOwnerId) throw new Error("Invalid data for creating brand");
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

  /** Update the entity with new entity data. Returns true in success */
  async updateEntity(entityId: number, entityData: OpportunityInput | BrandInput): Promise<boolean> {
    const oppBeforeChanges = await this.appData.getEntity(entityId, false);
    const success = await this.appData.updateEntity(entityId, entityData);

    if (success && entityData.EntityOwnerId && oppBeforeChanges.EntityOwnerId !== entityData.EntityOwnerId) { // owner changed
      return this.permissions.changeEntityOwnerPermissions(entityId, oppBeforeChanges.EntityOwnerId, entityData.EntityOwnerId);
    }

    return success;
  }
  
  /** Update the entity stage with new data. Returns true in success */
  async updateStageSettings(stageId: number, data: any): Promise<boolean> {
    const currentStage = await this.appData.getEntityStage(stageId);
    let success = await this.appData.updateStage(stageId, data);

    return success && await this.permissions.changeStageUsersPermissions(
      currentStage.EntityNameId,
      currentStage.StageNameId,
      currentStage.StageUsersId,
      data.StageUsersId
    );
  }

  async createEntityForecastCycle(entity: Opportunity, values: any) {
    const geographies = await this.appData.getEntityGeographies(entity.ID); // 1 = stage id would be dynamic in the future
    let archivedBasePath = `${FOLDER_ARCHIVED}/${entity.BusinessUnitId}/${entity.ID}/0/0`;
    let approvedBasePath = `${FOLDER_APPROVED}/${entity.BusinessUnitId}/${entity.ID}/0/0`;
    let workInProgressBasePath = `${FOLDER_WIP}/${entity.BusinessUnitId}/${entity.ID}/0/0`;

    let cycle = await this.appData.createEntityForecastCycle(entity);

    for (const geo of geographies) {
      let geoFolder = `${archivedBasePath}/${geo.ID}/${cycle.ID}`;
      const cycleFolder = await this.appData.createFolder(geoFolder, true);
      if(cycleFolder) {
        await this.files.moveAllFolderFiles(`${approvedBasePath}/${geo.ID}/0`, geoFolder);
      }else {
        throw new Error("Could not create Forecast Cycle folder");
      }
    }

    let changes = {
      ForecastCycleId: values.ForecastCycle,
      ForecastCycleDescriptor: values.ForecastCycleDescriptor,
      Year: values.Year
    };
    await this.appData.updateEntity(entity.ID, changes);

    await this.files.restartModelsInFolder(entity, workInProgressBasePath);

    return changes;
  }

  async getBrandModelsCount(brand: Opportunity) {
    return await this.files.getBrandFolderFilesCount(brand, FOLDER_WIP);
  }

  async getBrandApprovedModelsCount(brand: Opportunity) {
    return await this.files.getBrandFolderFilesCount(brand, FOLDER_APPROVED);
  }

  async isInternalOpportunity(oppTypeId: number): Promise<boolean> {
    const oppType = await this.appData.getOpportunityType(oppTypeId);
    if (oppType?.IsInternal) {
      return oppType.IsInternal;
    }
    return false;
  }

  async archiveEntity(entityId: number) {
    return await this.appData.setOpportunityStatus(entityId, "Archive");
  }

  async activeEntity(entityId: number) {
    return await this.appData.setOpportunityStatus(entityId, "Active");
  }

  async approveEntity(entityId: number) {
    return await this.appData.setOpportunityStatus(entityId, "Approved");
  }

  async getProgress(entity: Opportunity) {
    console.log('entity', entity);
    if (entity.OpportunityTypeId && await this.isInternalOpportunity(entity.OpportunityTypeId)) {
      return -1; // progress no applies
    }
    let actions = await this.appData.getActions(entity.ID);
    if (actions.length) {
      let gates: {'total': number; 'completed': number}[] = [];
      let currentGate = 0;
      let gateIndex = 0;
      for(let act of actions) {
        if (act.StageNameId == currentGate) {
          gates[gateIndex-1]['total']++;
          if (act.Complete) gates[gateIndex-1]['completed']++;
        } else {
          currentGate = act.StageNameId;
          if (act.Complete) gates[gateIndex] = {'total': 1, 'completed': 1};
          else gates[gateIndex] = {'total': 1, 'completed': 0};
          gateIndex++;
        }
      }

      let gatesMedium = gates.map(function(x) { return x.completed / x.total; });
      return Math.round((gatesMedium.reduce((a, b) => a + b, 0) / gatesMedium.length) * 10000) / 100;
    }
    return 0;
  }

}
