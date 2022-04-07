import { Injectable } from '@angular/core';
import { AppType } from '@shared/models/app-config';
import { Subject } from 'rxjs';
import { environment } from 'src/environments/environment';
import { Opportunity } from '../shared/models/entity';
import { NPPFile, NPPFileMetadata } from '../shared/models/file-system';
import { AppDataService } from './app-data.service';
import { ErrorService } from './error.service';
import { SharepointService } from './sharepoint.service';
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

  constructor(private readonly sharepoint: SharepointService, private readonly teams: TeamsService, private readonly error: ErrorService,
    private readonly appData: AppDataService) { 
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

    this.config = await this.appData.getAppConfig();
    console.log('config', this.config);
    let apps = await this.appData.getApp(appTitle);
    this.app = (apps && apps.length) ? apps[0] : undefined;

    if(!this.app) {
      this.error.handleError(new Error("Could not find ID for app: "+appTitle));
    } else {
      this.appData.app = this.app;
      this.isReady = true;
      this.readySubscriptions.next(true);
    }
  }

  getAppType() {
    return this.app;
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

  // getEntity(id: number) {
  //   if(this.isInline) {
  //     return this.appData.getBrand(id);
  //   } else {
  //     return this.appData.getOpportunity(id);
  //   }
  // }

  async getEntities() {
    if(this.app) {
      return this.appData.getAllEntities(this.app.ID);
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
    return this.appData.getEntityForecastCycles(entity);
  }

  readFolderFiles(folder: string, expandProperties: boolean) {
    return this.appData.getFolderFiles(folder, expandProperties);
  }

  getAccessibleGeographiesList(entity: Opportunity) {
    return this.appData.getEntityAccessibleGeographiesList(entity as Opportunity);
  }
  
  getEntityGeographies(entityId: number) {
    return this.appData.getEntityGeographies(entityId);
  }

  // getFileByScenarios(fileFolder: string, scenario: number[]) {
  //   return this.appData.getFileByScenarios(fileFolder, scenario);
  // }

  // async uploadFile(fileData: string, folder: string, fileName: string, metadata?: NPPFileMetadata) {
  //   return this.appData.uploadInternalFile(fileData, folder, fileName, metadata);
  // }

  // async setEntityApprovalStatus(rootFolder: string, file: NPPFile, entity: Opportunity | null, status: string, comments: string | null = null) {
  //   return this.appData.setEntityApprovalStatus(rootFolder, file, entity, status, comments);
  // }

  // async createForecastCycle(entity: Opportunity, values: any) {
  //   return this.appData.createEntityForecastCycle(entity, values);    
  // }

}
