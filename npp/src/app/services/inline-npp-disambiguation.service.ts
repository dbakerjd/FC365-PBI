import { Injectable } from '@angular/core';
import { environment } from 'src/environments/environment';
import { ErrorService } from './error.service';
import { AppType, Brand, NPPFile, NPPFileMetadata, Opportunity, SharepointService } from './sharepoint.service';
import { TeamsService } from './teams.service';

@Injectable({
  providedIn: 'root'
})
export class InlineNppDisambiguationService {
  
  isInline: boolean = false;
  app: AppType | undefined;

  constructor(private readonly sharepoint: SharepointService, private readonly teams: TeamsService, private readonly error: ErrorService) { 
    this.isInline = environment.isInlineApp;

    if (this.teams.initialized) {
      this.setApp();
    } else {
      this.teams.statusSubject.subscribe(async (msg) => {
        if(msg == 'initialized') {
          this.setApp();
        }
      })
    }
    
  }

  async setApp() {
    let appTitle = 'NPP';
    if(this.isInline) {
      appTitle = 'Inline';
    }

    let apps = await this.sharepoint.getApp(appTitle);
    this.app = (apps && apps.length) ? apps[0] : undefined;

    if(!this.app) {
      this.error.handleError(new Error("Could not find ID for app: "+appTitle));
    } else {
      this.sharepoint.app = this.app;
    }
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
      return (entity as Opportunity).EntityOwnerId;
    }
  }

  getOwner(entity: Brand | Opportunity) {
    if(this.isInline) {
      return (entity as Brand).BrandOwner;
    } else {
      return (entity as Opportunity).EntityOwner;
    }
  }

  getForecastCycles(entity: Brand | Opportunity) {
    return this.sharepoint.getEntityForecastCycles(entity);
  }

  readFolderFiles(folder: string, expandProperties: boolean) {
    if(this.isInline) {
      return this.sharepoint.readBrandFolderFiles(folder, expandProperties);
    } else {
      return this.sharepoint.readOpportunityFolderFiles(folder, expandProperties);
    }
  }

  getAccessibleGeographiesList(entity: Brand | Opportunity) {
    return this.sharepoint.getEntityAccessibleGeographiesList(entity as Opportunity);
  }
  
  getEntityGeographies(entityId: number) {
    return this.sharepoint.getEntityGeographies(entityId);
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
      return this.sharepoint.updateEntityGeographyUsers(entityId, geoId, currentUsersList, newUsersList);
    }
  }

  async setEntityApprovalStatus(rootFolder: string, file: NPPFile, entity: Brand | Opportunity | null, status: string, comments: string | null = null) {
    return this.sharepoint.setEntityApprovalStatus(rootFolder, file, entity, status, comments);
  }

  async createForecastCycle(entity: Brand | Opportunity, values: any) {
    return this.sharepoint.createEntityForecastCycle(entity, values);    
  }

  getGroupName(name: string):string {
    if(this.isInline) {
      name = name.replace("EU-", "BU-");
      name = name.replace("EO-", "BO-");
    } else {
      name = name.replace("EU-", "OU-");
      name = name.replace("EO-", "OO-");
    }
    return name;
  }

}
