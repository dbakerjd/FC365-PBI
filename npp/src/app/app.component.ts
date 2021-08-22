import { Component, Inject, OnDestroy, OnInit } from '@angular/core';
import { MsalBroadcastService, MsalGuardConfiguration, MsalService, MSAL_GUARD_CONFIG } from '@azure/msal-angular';
import { AuthenticationResult, EventMessage, EventType, InteractionStatus, RedirectRequest } from '@azure/msal-browser';
import { Subject } from 'rxjs';
import { filter, takeUntil } from 'rxjs/operators';
import { LicensingService } from './services/licensing.service';
import { TeamsService } from './services/teams.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit, OnDestroy {
  isIframe = false;
  loginDisplay = false;
  private readonly _destroying$ = new Subject<void>();
  
  constructor(
    @Inject(MSAL_GUARD_CONFIG) private msalGuardConfig: MsalGuardConfiguration,
    private readonly teams: TeamsService, 
    private authService: MsalService, 
    private msalBroadcastService: MsalBroadcastService, 
    private licensing: LicensingService
  ) {

  }
  ngOnInit(): void {
    this.isIframe = window !== window.parent && !window.opener; // Remove this line to use Angular Universal
    this.setLoginDisplay();

    this.authService.handleRedirectObservable().subscribe();

    this.msalBroadcastService.msalSubject$
      .pipe(
        filter((msg: EventMessage) => msg.eventType === EventType.LOGIN_SUCCESS),
      )
      .subscribe((result: EventMessage) => {
        console.log('result login', result);
        const payload = result.payload as AuthenticationResult;
        this.authService.instance.setActiveAccount(payload.account);
      });

    this.msalBroadcastService.inProgress$
      .pipe(
        filter((status: InteractionStatus) => status === InteractionStatus.None),
        takeUntil(this._destroying$)
      )
      .subscribe(() => {
        console.log('IN PROGRESS');
        this.setLoginDisplay();
        this.checkAndSetActiveAccount();
      })

    this.checkAndSetActiveAccount();

      /*
    this.authService.handleRedirectObservable().subscribe();

    this.msalBroadcastService.msalSubject$
      .pipe(
        filter((msg: EventMessage) => msg.eventType === EventType.LOGIN_SUCCESS),
        takeUntil(this._destroying$)
      )
      .subscribe((result: EventMessage) => {
        console.log('login success', result);
        const payload = result.payload as AuthenticationResult;
        console.log('payload', payload);
        if (payload.account) {
          this.teams.setActiveAccount(payload.account);
        }
        // this.teams.setToken(payload.accessToken);
    });
    */

    /*
    this.msalBroadcastService.msalSubject$
        .pipe(
            filter((msg: EventMessage) => msg.eventType === EventType.ACQUIRE_TOKEN_SUCCESS),
            takeUntil(this._destroying$)
        )
        .subscribe((result: EventMessage) => {
            const payload = result.payload as AuthenticationResult;
            if (payload.account) {
              this.teams.setActiveAccount(payload.account);
              // this.teams.setToken(payload.accessToken);
            }
        });
        */

    // this.teams.getActiveAccount();
    // console.log("l", this.licensing.hasJplusDLicense( this.teams.token));
    // this.teams.logout();
    // let activeAccount = this.authService.instance.getActiveAccount();
    // if (!activeAccount) {
    //   this.authService.loginRedirect();
    // }

  }

  setLoginDisplay() {
    this.loginDisplay = this.authService.instance.getAllAccounts().length > 0;
  }

  checkAndSetActiveAccount(){
    /**
     * If no active account set but there are accounts signed in, sets first account to active account
     * To use active account set here, subscribe to inProgress$ first in your component
     * Note: Basic usage demonstrated. Your app may require more complicated account selection logic
     */
    let activeAccount = this.authService.instance.getActiveAccount();

    if (!activeAccount && this.authService.instance.getAllAccounts().length > 0) {
      let accounts = this.authService.instance.getAllAccounts();
      this.authService.instance.setActiveAccount(accounts[0]);
    } else if(!activeAccount) {
      this.loginRedirect();
    }
  }

  loginRedirect() {
    if (this.msalGuardConfig.authRequest){
      this.authService.loginRedirect({...this.msalGuardConfig.authRequest} as RedirectRequest);
    } else {
      this.authService.loginRedirect();
    }
  }

  logout() {
    this.teams.logout();
  }

  ngOnDestroy(): void {
    this._destroying$.next();
    this._destroying$.complete();
  }
}
