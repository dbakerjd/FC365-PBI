import { Inject, Injectable } from '@angular/core';
import { MsalGuardConfiguration, MsalService, MSAL_GUARD_CONFIG } from '@azure/msal-angular';
import { PopupRequest } from '@azure/msal-browser';
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

  loginAgain() {
    this.token = null;
    localStorage.removeItem('teamsAccessToken');
    localStorage.removeItem('teamsAccount');
    this.getActiveAccount();
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
