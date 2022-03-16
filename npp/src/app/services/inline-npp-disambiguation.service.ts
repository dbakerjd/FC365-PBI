import { Injectable } from '@angular/core';
import { Subject } from 'rxjs';
import { environment } from 'src/environments/environment';
import { ErrorService } from './error.service';
import { AppType, NPPFile, NPPFileMetadata, Opportunity, SharepointService } from './sharepoint.service';
import { TeamsService } from './teams.service';

@Injectable({
  providedIn: 'root'
})
export class InlineNppDisambiguationService {
  
  isInline: boolean = false;
  isReady: boolean = false;
  readySubscriptions: Subject<boolean> = new Subject<boolean>();
  app: AppType | undefined;
  config: { Title: string; Value: string, ConfigType: string }[] = [];

  constructor(private readonly sharepoint: SharepointService, private readonly teams: TeamsService, private readonly error: ErrorService) { 
    this.isInline = environment.isInlineApp;
    this.isReady = false;

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

    this.config = await this.sharepoint.getAppConfig();
    console.log('config', this.config);
    let apps = await this.sharepoint.getApp(appTitle);
    this.app = (apps && apps.length) ? apps[0] : undefined;

    if(!this.app) {
      this.error.handleError(new Error("Could not find ID for app: "+appTitle));
    } else {
      this.sharepoint.app = this.app;
      this.isReady = true;
      this.readySubscriptions.next(true);
    }
  }

  getConfigValue(name: string): any {
    const item = this.config.find(el => el.Title === name);
    if (item) {
      switch (item.ConfigType) {
        case 'Boolean': return +item.Value != 0 ? true : false; 
        case 'String': return item.Value;
        case 'Number': return +item.Value;
        default: return item.Value;
      }
    }
    return undefined;
  }

  getEntity(id: number) {
    if(this.isInline) {
      return this.sharepoint.getBrand(id);
    } else {
      return this.sharepoint.getOpportunity(id);
    }
  }

  async getEntities() {
    if(this.app) {
      return this.sharepoint.getAllEntities(this.app.ID);
    } else {
      this.error.toastr.error("Tried to get Entities but the app was not ready yet.")
      return [];
    }
    
  }

  getOwnerId(entity: Opportunity) {
    return entity.EntityOwnerId;
  }

  getOwner(entity: Opportunity) {
    return entity.EntityOwner;
  }

  getForecastCycles(entity: Opportunity) {
    return this.sharepoint.getEntityForecastCycles(entity);
  }

  readFolderFiles(folder: string, expandProperties: boolean) {
    return this.sharepoint.readEntityFolderFiles(folder, expandProperties);
  }

  getAccessibleGeographiesList(entity: Opportunity) {
    return this.sharepoint.getEntityAccessibleGeographiesList(entity as Opportunity);
  }
  
  getEntityGeographies(entityId: number) {
    return this.sharepoint.getEntityGeographies(entityId);
  }

  getFileByScenarios(fileFolder: string, scenario: number[]) {
    return this.sharepoint.getFileByScenarios(fileFolder, scenario);
  }

  async uploadFile(fileData: string, folder: string, fileName: string, metadata?: NPPFileMetadata) {
    return this.sharepoint.uploadInternalFile(fileData, folder, fileName, metadata);
  }

  async setEntityApprovalStatus(rootFolder: string, file: NPPFile, entity: Opportunity | null, status: string, comments: string | null = null) {
    return this.sharepoint.setEntityApprovalStatus(rootFolder, file, entity, status, comments);
  }

  async createForecastCycle(entity: Opportunity, values: any) {
    return this.sharepoint.createEntityForecastCycle(entity, values);    
  }

}
