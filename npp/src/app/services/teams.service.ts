import { Inject, Injectable } from '@angular/core';
import { Router } from '@angular/router';
import { LogLevel, PopupRequest, AccountInfo, RedirectRequest, AuthenticationResult, PublicClientApplication, BrowserCacheLocation, InteractionType } from '@azure/msal-browser';
import * as microsoftTeams from "@microsoft/teams-js";
import { Subject } from 'rxjs';
import { environment } from 'src/environments/environment';
import { ErrorService } from './error.service';
import { LicensingService } from './licensing.service';

const isIE = window.navigator.userAgent.indexOf("MSIE ") > -1 || window.navigator.userAgent.indexOf("Trident/") > -1; // Remove this line to use Angular Universal

export function loggerCallback(logLevel: LogLevel, message: string) {
  console.log(message);
}


@Injectable({
  providedIn: 'root'
})
export class TeamsService {
  public account: any = false;
  public user: any = false;
  public token: any = false;
  public context: any = false;
  public currentlyLoginIn = false;
  public authObj: string = '';
  public hackyConsole: string = '';
  public statusSubject = new Subject<string>();

  //David's
  //'e504af88-0105-426f-bd33-9990e49c8122'
  //'https://janddconsulting.sharepoint.com/.default'
  //Beta's
  //'17534ca2-f4f8-43c0-8612-72bdd29a9ee8'
  //'https://betasoftwaresl.sharepoint.com/.default'

  public msalInstance = new PublicClientApplication({
    auth: {
      // clientId: '6226576d-37e9-49eb-b201-ec1eeb0029b6', // Prod enviroment. Uncomment to use. 
      clientId: 'e504af88-0105-426f-bd33-9990e49c8122', // PPE testing environment
      authority: 'https://login.microsoftonline.com/common', // Prod environment. Uncomment to use.
      //authority: 'https://login.windows-ppe.net/common', // PPE testing environment.
      redirectUri: environment.ssoRedirectUrl,
      postLogoutRedirectUri: environment.ssoRedirectUrl
    },
    cache: {
      cacheLocation: BrowserCacheLocation.LocalStorage,
      storeAuthStateInCookie: isIE, // set to true for IE 11. Remove this line to use Angular Universal
    },
    system: {
      loggerOptions: {
        loggerCallback,
        logLevel: LogLevel.Info,
        piiLoggingEnabled: false
      }
    }
  });

  constructor( private router: Router, private errorService: ErrorService, private licensing: LicensingService) { 

    microsoftTeams.initialize(() => {
      this.statusSubject.next("initialized");
    });
    
    microsoftTeams.getContext((context) => {
      this.context = context;
      this.validateLicense();
    });

    this.msalInstance.handleRedirectPromise().then((tokenResponse) => {
      if(tokenResponse) {
        console.log(tokenResponse);
        microsoftTeams.authentication.notifySuccess(JSON.stringify(tokenResponse));
      } else {
        console.log("empty tokenResponse"); 
      }
      // Check if the tokenResponse is null
      // If the tokenResponse !== null, then you are coming back from a successful authentication redirect. 
      // If the tokenResponse === null, you are not coming back from an auth redirect.
    }).catch((error) => {
        // handle error, either in the library or coming back from the server
        this.errorService.handleError(error);
    });
  }

  getResourceMap() {
    /*if(!this.licensing.license) {
      console.log("Trying to get resources without an active license, failing silently.");
      return;
    }*/

    const protectedResourceMap = new Map<string, Array<string>>();
    protectedResourceMap.set('janddconsulting.sharepoint.com', ['https://janddconsulting.sharepoint.com/.default']);
    //protectedResourceMap.set('betasoftwaresl.sharepoint.com', ['https://betasoftwaresl.sharepoint.com/.default']);
    // protectedResourceMap.set('https://betasoftwaresl.sharepoint.com', ['AllSites.FullControl', 'AllSites.Manage', 'Sites.Search.All']);
    //https://nppprovisioning20210831.azurewebsites.net/api/NewOpportunity?StageID=2&OppID=1&siteUrl=https://janddconsulting.sharepoint.com/sites/NPPDemoV15
    protectedResourceMap.set('nppprovisioning20210831.azurewebsites.net', ['api://b431132e-d7ea-4206-a0a9-5403adf64155/.default']);
  
    return {
      interactionType: InteractionType.Redirect,
      protectedResourceMap
    };

  }

  getResourceByDomain(domain: string) {
    let map = this.getResourceMap();
    return map?.protectedResourceMap.get(domain);
  }

  getMSALGuardConfig() {
    /*if(!this.licensing.license) {
      console.log("Trying to get guard config without an active license, failing silently.");
      return;
    }*/

    return { 
      interactionType: InteractionType.Redirect,
      authRequest: {
<<<<<<< HEAD
        //scopes: ['https://betasoftwaresl.sharepoint.com/.default']
        scopes: ['https://janddconsulting.sharepoint.com/.default']
=======
        scopes: ['https://betasoftwaresl.sharepoint.com/.default', 'api://b431132e-d7ea-4206-a0a9-5403adf64155/.default']
>>>>>>> feature/teams-login
      },
      loginFailedRoute: '/auth-end'
    };
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
    return this.token;
  }

  async checkAndSetActiveAccount(){
    /**
     * If no active account set but there are accounts signed in, sets first account to active account
     * To use active account set here, subscribe to inProgress$ first in your component
     */
    let activeAccount = this.msalInstance.getActiveAccount();
    if (!activeAccount && this.msalInstance.getAllAccounts().length > 0) {
      let accounts = this.msalInstance.getAllAccounts();
      this.msalInstance.setActiveAccount(accounts[0]);
    } else if(!activeAccount) {
      await this.login();
    } 
  }

  async login() {
    if(!this.currentlyLoginIn) {
      this.currentlyLoginIn = true;
      microsoftTeams.authentication.authenticate({
        url: window.location.origin + "/auth-start",
        width: 600,
        height: 535,
        successCallback: async (result) => {
          try {
            this.currentlyLoginIn = false;
            const payload = JSON.parse(result ? result : '')  as AuthenticationResult;
            this.authObj = JSON.stringify(payload);
            this.setActiveAccount(payload.account);
            this.setToken(payload.accessToken);
          } catch(e: any) {
            this.hackyConsole += "*************ERROR************* -> "+e+"      -      ";
            this.errorService.handleError(e);
          }
          
        },
        failureCallback: (error) => {
          this.hackyConsole += "got error called: "+error+"     -     ";
            this.currentlyLoginIn = false;
            this.errorService.handleError(error ? new Error(error) : new Error("Something went wrong trying to log in"));
        }
      });
    }
  }

  async validateLicense() {
    this.licensing.validateLicense(this.context);
  }

  async logout() {
    localStorage.removeItem('teamsAccount');
    this.msalInstance.logoutRedirect();
  }

  async getActiveAccount() {
    return await this.msalInstance.getActiveAccount();
  }
  
  setActiveAccount(account: AccountInfo | null) {
    if (account) {
      this.msalInstance.setActiveAccount(account);
      this.account = account;
    }
  }

}
