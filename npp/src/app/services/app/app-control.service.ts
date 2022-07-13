import { Injectable } from '@angular/core';
import { AppType } from '@shared/models/app-config';
import { Subject } from 'rxjs';
import { environment } from 'src/environments/environment';
import { AppDataService } from '@services/app/app-data.service';
import { ErrorService } from './error.service';
import { TeamsService } from '@services/microsoft-data/teams.service';
import { Router } from '@angular/router';

@Injectable({
  providedIn: 'root'
})
export class AppControlService {
  
  isInline: boolean = false;
  isReady: boolean = false;
  readySubscriptions: Subject<boolean> = new Subject<boolean>();
  app: AppType | undefined;
  config: { Title: string; Value: string, ConfigType: string }[] = [];

  constructor(
    private readonly teams: TeamsService, 
    private readonly error: ErrorService,
    private readonly appData: AppDataService,
    private router: Router
  ) { 
    this.isInline = environment.isInlineApp;
    this.isReady = false;

    if (this.teams.initialized) {
      this.setApp();
    } else {
      this.teams.statusSubject.subscribe(async (msg) => {
        if(msg == 'loggedIn') {
          // check if we are allowed to connect to the license sharepoint
          if (await this.appData.canConnectAndAccessData()) {
            this.setApp();
          } else {
            this.router.navigate(['splash/non-access']); 
          }
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

  /** Check for the value of an app config value by name */
  getAppConfigValue(name: string): any {
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

  /**
   * Check if the user is assigned at least to one entity information
   * 
   * @param userId User Id. If not present, the current user
   * @returns boolean
   */
   async userHasAccessToEntities(userId?: number) {
    const user = userId ? await this.appData.getUserInfo(userId) : await this.appData.getCurrentUserInfo();
    const currentUserGroups = await this.appData.getUserGroups(user.Id);
    return !!currentUserGroups.find(g => g.Title.startsWith('OU')) || !!user.IsSiteAdmin;
  }

}
