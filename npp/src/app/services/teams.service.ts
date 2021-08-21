import { Inject, Injectable } from '@angular/core';
import { MsalGuardConfiguration, MsalService, MSAL_GUARD_CONFIG } from '@azure/msal-angular';
import { PopupRequest, AccountInfo } from '@azure/msal-browser';
import * as microsoftTeams from "@microsoft/teams-js";
import { ErrorService } from './error.service';

@Injectable({
  providedIn: 'root'
})
export class TeamsService {
  public account: any = false;
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

  setToken(token: string) {
    this.token = token;
    this.setStorageToken(token);
  }

  setStorageToken(token: string) {
    localStorage.setItem('teamsAccessToken', token);
  }

  getStorageToken() {
    this.token = localStorage.getItem('teamsAccessToken');
  }

  async refreshToken(force = false) {
    this.authService.logoutRedirect();
    if (this.getStorageToken() == null || force) {
      console.log('no token found in storage');
      /*
      let activeAccount = this.authService.instance.getActiveAccount();
      if (activeAccount) {
        let newToken = await this.authService.instance.acquireTokenSilent({scopes: ["Sites.FullControl", "user.read"], account: activeAccount['name'] as AccountInfo | undefined}).then(function(accessTokenResponse) {
          return accessTokenResponse;
        });
        // this.setToken(newToken.accessToken);
        console.log('setted new token', this.token);
      } else {
        this.authService.loginRedirect();
      }
      */
    }
  }

  async logout() {
    console.log('logout redirect');
    // this.authService.logoutRedirect();
    // this.authService.logout();
  }

  async loginAgain() {
    // this.authService.logoutRedirect();
    /*
    this.token = null;
    localStorage.removeItem('teamsAccessToken');
    localStorage.removeItem('teamsAccount');
    this.getActiveAccount();
    */
  }

  getActiveAccount() {
    let activeAccount = this.authService.instance.getActiveAccount();

    if(!activeAccount) {
      activeAccount = this.getStorageAccount();
      if(activeAccount) this.authService.instance.setActiveAccount(activeAccount);
      activeAccount = this.authService.instance.getActiveAccount();
    }

    if (!activeAccount) {
      let accounts = this.authService.instance.getAllAccounts();
      if (!accounts || accounts.length == 0) {
        if(this.msalGuardConfig.authRequest){
          this.authService.loginRedirect({...this.msalGuardConfig.authRequest} as PopupRequest);
        } else {
          this.authService.loginRedirect();
        }
        return false;
      } else {
        this.setActiveAccount(accounts[0]);
        activeAccount = this.authService.instance.getActiveAccount();
      }
    }

    return activeAccount;  
  }

  setActiveAccount(account: any) {
    this.authService.instance.setActiveAccount(account);
    this.account = account;
    this.setStorageAccount(account);
  }

  setStorageAccount(account: any) {
    localStorage.setItem('teamsAccount', JSON.stringify(account));
  }

  getStorageAccount() {
    this.getStorageToken();
    let account = localStorage.getItem('teamsAccount');
    if(account) {
      return JSON.parse(account);
    } else {
      return false;
    }
  }


}
