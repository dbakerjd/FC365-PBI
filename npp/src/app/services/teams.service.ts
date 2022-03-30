import { Inject, Injectable } from '@angular/core';
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

export interface LicenseContext {
  entityId: string;
  teamSiteDomain: string;
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
  public initialized = false;
  private userLoggedIn = false;
  private isLoadedInsideTeams = false;
  private _hasAttemptedConnection = false; // control attempted connection to teams

  //David's
  //'e504af88-0105-426f-bd33-9990e49c8122'
  //'https://janddconsulting.sharepoint.com/.default'
  //Beta's
  //'17534ca2-f4f8-43c0-8612-72bdd29a9ee8'
  //'https://betasoftwaresl.sharepoint.com/.default'

  public msalInstance = new PublicClientApplication({
    auth: {
      //clientId: '17534ca2-f4f8-43c0-8612-72bdd29a9ee8', // Prod enviroment. Uncomment to use. 
      clientId: '6c76f4df-ba13-4aca-8e16-d7c0bb9d9a51',
      //clientId: 'e504af88-0105-426f-bd33-9990e49c8122', // PPE testing environment
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

  constructor( private errorService: ErrorService, private licensing: LicensingService) { 

    this.context = this.getEnvironmentContext();
    this.startTeams();

    while (!this._hasAttemptedConnection); // wait for start teams attempt
    setTimeout(() => {
      if (!this.isLoadedInTeams()) {
        this.initialized = true;
        this.statusSubject.next("initialized");
      }
      this.validateLicense();
    }, 500);

   this.msalInstance.handleRedirectPromise().then((tokenResponse) => {
    if(tokenResponse) {
      if (this.isLoadedInTeams()) {
        microsoftTeams.authentication.notifySuccess(JSON.stringify(tokenResponse));
      }
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

    errorService.subject.subscribe(msg => {
      if(msg == 'unauthorized') {
        this.login();
      }
    })
  }

  setLoggedIn(isLoggedIn = true) {
    if (isLoggedIn) {
      this.statusSubject.next("loggedIn");
    }
    this.userLoggedIn = isLoggedIn;
  }

  isLoggedIn(): boolean {
    return this.userLoggedIn;
  }

  getEnvironmentContext(): LicenseContext | null {
    if (environment.licensingInfo) {
      if (environment.licensingInfo.entityId && environment.licensingInfo.teamSiteDomain) {
        return {
          entityId: environment.licensingInfo.entityId,
          teamSiteDomain: environment.licensingInfo.teamSiteDomain
        };
      } else {
        this.errorService.handleError({ message: 'Bad licensing info in app environment' });
      }
    }
    return null;
  }

  startTeams() {
    this._hasAttemptedConnection = false;
    microsoftTeams.initialize(() => {
      this.initialized = true;
      microsoftTeams.getContext((context) => {
        if (context) this.isLoadedInsideTeams = true;
        this.context = context;
        this.statusSubject.next("initialized");
      });
    });
    this._hasAttemptedConnection = true;
  }

  isLoadedInTeams() {
    return this._hasAttemptedConnection && this.isLoadedInsideTeams;
  }

  getResourceMap() {
    if(!this.licensing.license) {
      this.errorService.toastr.error("Trying to get resources without an active license");
      return;
    }

    const protectedResourceMap = new Map<string, Array<string>>();
    //protectedResourceMap.set('janddconsulting.sharepoint.com', ['https://janddconsulting.sharepoint.com/.default']);
    //protectedResourceMap.set('betasoftwaresl.sharepoint.com', ['https://betasoftwaresl.sharepoint.com/.default']);
    
    let sharepointUri = this.licensing.getSharepointDomain();
    if(sharepointUri) {
      protectedResourceMap.set(sharepointUri, ['https://'+sharepointUri+'/.default']);
    }
    protectedResourceMap.set('graph.microsoft.com', ['User.Read', 'GroupMember.ReadWrite.All']);
    protectedResourceMap.set('api.powerbi.com', ['https://analysis.windows.net/powerbi/api/.default']);
    protectedResourceMap.set(environment.functionAppDomain,['https://janddconsulting.onmicrosoft.com/FC365-Test-NPP/user_impersonation']);
    
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
    if(!this.licensing.license) {
      this.errorService.toastr.error("Trying to get resources without an active license 2");
      return;
    }
    //'api://b431132e-d7ea-4206-a0a9-5403adf64155/.default'
    //let scopes = ['api://b431132e-d7ea-4206-a0a9-5403adf64155/.default', 'User.Read', 'https://analysis.windows.net/powerbi/api/.default'];
    let scopes = [];
    let sharepointUri = this.licensing.getSharepointDomain();
    if(sharepointUri) {
      scopes.push('https://'+sharepointUri+'/.default');
      // scopes.push('https://graph.microsoft.com/.default');
    }

    return { 
      interactionType: InteractionType.Redirect,
      authRequest: {
        scopes
      },
      loginFailedRoute: '/auth-end'
    };
  }

  /** could be deleted */
  setToken(token: string) {
    this.token = token;
    this.setStorageToken(token);
  }

  /** could be deleted */
  setStorageToken(token: string) {
    localStorage.setItem('teamsAccessToken', token);
  }

  /** unused */
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
    if (activeAccount) this.setLoggedIn(true);
    else if (!activeAccount && this.msalInstance.getAllAccounts().length > 0) {
      let accounts = this.msalInstance.getAllAccounts();
      this.msalInstance.setActiveAccount(accounts[0]);
      this.setLoggedIn(true);
    } else if(!activeAccount) {
      await this.login();
    } 
  }

  async login() {
    if(!this.currentlyLoginIn) {
      this.currentlyLoginIn = true;
      localStorage.removeItem('teamsAccount');
      localStorage.removeItem("teamsAccessToken");
      let sharepointUrl = this.licensing.getSharepointApiUri();
      let accountStorageKey = sharepointUrl + '-sharepointAccount';
      localStorage.removeItem(accountStorageKey);
      if (!this.isLoadedInTeams()) {
        this.msalInstance.loginRedirect();
        this.currentlyLoginIn = false;
        // this.setToken() // not necessary / token in storage unused
      } else {
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
            // token in teams service is unused, could be deleted
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
  }

  async validateLicense() {
    await this.licensing.validateLicense(this.context);
    this.statusSubject.next('license');
  }

  async logout() {
    localStorage.removeItem('teamsAccount');
    localStorage.removeItem("teamsAccessToken");
    let sharepointUrl = this.licensing.getSharepointApiUri();
    let accountStorageKey = sharepointUrl + '-sharepointAccount';
    localStorage.removeItem(accountStorageKey);
    this.msalInstance.logoutRedirect();
  }

  async getActiveAccount() {
    return await this.msalInstance.getActiveAccount();
  }
  
  setActiveAccount(account: AccountInfo | null) {
    if (account) {
      this.msalInstance.setActiveAccount(account);
      this.account = account;
      this.statusSubject.next('loggedIn');
    }
  }

}
