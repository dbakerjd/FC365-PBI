import { Injectable } from '@angular/core';
import { Subject } from 'rxjs';
import { environment } from 'src/environments/environment';
import { ErrorService } from './error.service';
import { AppType, Brand, NPPFile, NPPFileMetadata, Opportunity, SharepointService } from './sharepoint.service';
import { TeamsService } from './teams.service';

@Injectable({
  providedIn: 'root'
})
export class InlineNppDisambiguationService {
  
  isInline: boolean = false;
  isReady: boolean = false;
  readySubscriptions: Subject<boolean> = new Subject<boolean>();
  app: AppType | undefined;

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

  getOwnerId(entity: Brand | Opportunity) {
    return entity.EntityOwnerId;
  }

  getOwner(entity: Brand | Opportunity) {
    return entity.EntityOwner;
  }

  getForecastCycles(entity: Brand | Opportunity) {
    return this.sharepoint.getEntityForecastCycles(entity);
  }

  readFolderFiles(folder: string, expandProperties: boolean) {
    return this.sharepoint.readEntityFolderFiles(folder, expandProperties);
  }

  getAccessibleGeographiesList(entity: Brand | Opportunity) {
    return this.sharepoint.getEntityAccessibleGeographiesList(entity as Opportunity);
  }
  
  getEntityGeographies(entityId: number) {
    return this.sharepoint.getEntityGeographies(entityId);
  }

  getFileByScenarios(fileFolder: string, scenario: number[]) {
    return this.sharepoint.getFileByScenarios(fileFolder, scenario);
  }

  async uploadFile(fileData: string, folder: string, fileName: string, metadata?: NPPFileMetadata) {
    if(this.isInline) {
      return this.sharepoint.uploadInlineFile(fileData, folder, fileName, metadata);
    } else {
      return this.sharepoint.uploadNPPFile(fileData, folder, fileName, metadata);
    }
  }

  async updateEntityGeographyUsers(entityId: number, geoId: number, currentUsersList: number[], newUsersList: number[]) {
    return this.sharepoint.updateEntityGeographyUsers(entityId, geoId, currentUsersList, newUsersList);
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
