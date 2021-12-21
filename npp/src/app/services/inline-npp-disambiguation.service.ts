import { Injectable } from '@angular/core';
import { environment } from 'src/environments/environment';
import { Brand, NPPFile, NPPFileMetadata, Opportunity, SharepointService } from './sharepoint.service';

@Injectable({
  providedIn: 'root'
})
export class InlineNppDisambiguationService {
  
  isInline: boolean = false;
  
  constructor(private readonly sharepoint: SharepointService) { 
    this.isInline = environment.isInlineApp;
  }

  getEntity(id: number) {
    if(this.isInline) {
      return this.sharepoint.getBrand(id);
    } else {
      return this.sharepoint.getOpportunity(id);
    }
  }

  getOwnerId(entity: Brand | Opportunity) {
    if(this.isInline) {
      return (entity as Brand).BrandOwnerId;
    } else {
      return (entity as Opportunity).OpportunityOwnerId;
    }
  }

  getOwner(entity: Brand | Opportunity) {
    if(this.isInline) {
      return (entity as Brand).BrandOwner;
    } else {
      return (entity as Opportunity).OpportunityOwner;
    }
  }

  getForecastCycles(entity: Brand | Opportunity) {
    if(this.isInline) {
      return this.sharepoint.getBrandForecastCycles(entity as Brand);
    } else {
      return this.sharepoint.getOpportunityForecastCycles(entity as Opportunity);
    }
  }

  readFolderFiles(folder: string, expandProperties: boolean) {
    if(this.isInline) {
      return this.sharepoint.readBrandFolderFiles(folder, expandProperties);
    } else {
      return this.sharepoint.readOpportunityFolderFiles(folder, expandProperties);
    }
  }

  getAccessibleGeographiesList(entity: Brand | Opportunity) {
    if(this.isInline) {
      return this.sharepoint.getOpportunityAccessibleGeographiesList(entity as Opportunity);
    } else {
      return this.sharepoint.getBrandAccessibleGeographiesList(entity as Brand);
    }
  }
  
  getEntityGeographies(entityId: number) {
    if(this.isInline) {
      return this.sharepoint.getOpportunityGeographies(entityId);
    } else {
      return this.sharepoint.getBrandGeographies(entityId);
    }
  }

  getFileByScenarios(fileFolder: string, scenario: number[]) {
    if(this.isInline) {
      return this.sharepoint.getFileByScenarios(fileFolder, scenario);
    } else {
      return this.sharepoint.getNPPFileByScenarios(fileFolder, scenario);
    }
  }

  async uploadFile(fileData: string, folder: string, fileName: string, metadata?: NPPFileMetadata) {
    if(this.isInline) {
      return this.sharepoint.uploadInlineFile(fileData, folder, fileName, metadata);
    } else {
      return this.sharepoint.uploadNPPFile(fileData, folder, fileName, metadata);
    }
  }

  async updateEntityGeographyUsers(entityId: number, geoId: number, currentUsersList: number[], newUsersList: number[]) {
    if(this.isInline) {
      return this.sharepoint.updateBrandGeographyUsers(entityId, geoId, currentUsersList, newUsersList);
    } else {
      return this.sharepoint.updateOpportunityGeographyUsers(entityId, geoId, currentUsersList, newUsersList);
    }
  }

  async setEntityApprovalStatus(rootFolder: string, file: NPPFile, entity: Brand | Opportunity | null, status: string, comments: string | null = null) {
    if(this.isInline) {
      return this.sharepoint.setBrandApprovalStatus(rootFolder, file, entity as Brand, "Approved", comments);
    } else {
      return this.sharepoint.setOpportunityApprovalStatus(rootFolder, file, entity as Opportunity, "Approved", comments);
    }
  }

  async createForecastCycle(entity: Brand | Opportunity, values: any) {
    if(this.isInline) {
      return this.sharepoint.createForecastCycle(entity as Brand, values);
    } else {
      return this.sharepoint.createOpportunityForecastCycle(entity as Opportunity, values);
    }
  }

}
