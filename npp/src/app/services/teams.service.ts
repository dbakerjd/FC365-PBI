import { Inject, Injectable } from '@angular/core';
import { MsalGuardConfiguration, MsalService, MSAL_GUARD_CONFIG } from '@azure/msal-angular';
import { PopupRequest } from '@azure/msal-browser';
import * as microsoftTeams from "@microsoft/teams-js";
import { ErrorService } from './error.service';

@Injectable({
  providedIn: 'root'
})
export class TeamsService {
  public user: any = false;
  public token: any = false;
  public context: any = false;

  constructor( @Inject(MSAL_GUARD_CONFIG) private msalGuardConfig: MsalGuardConfiguration,
               private errorService: ErrorService, private authService: MsalService) { 

    microsoftTeams.initialize();
    
    microsoftTeams.getContext((context) => {
      this.context = context;
      console.log(context);
    });
  }

  async getActiveAccount() {
    if(this.authService.instance.getAllAccounts().length == 0) {
      let response;
      if (this.msalGuardConfig.authRequest){
        response = await this.authService.loginPopup({...this.msalGuardConfig.authRequest} as PopupRequest).toPromise();
        
      } else {
        response = await this.authService.loginPopup().toPromise();
      }
      this.authService.instance.setActiveAccount(response.account);    
    } 

    let activeAccount = this.authService.instance.getActiveAccount();
    if (!activeAccount && this.authService.instance.getAllAccounts().length > 0) {
      let accounts = this.authService.instance.getAllAccounts();
      this.authService.instance.setActiveAccount(accounts[0]);
      activeAccount = this.authService.instance.getActiveAccount();
    }

    return activeAccount;
    
  }



}
