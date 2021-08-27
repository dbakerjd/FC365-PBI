import { Component, OnDestroy, OnInit } from '@angular/core';
import { MsalBroadcastService, MsalService } from '@azure/msal-angular';
import { AuthenticationResult, EventMessage, EventType, InteractionStatus } from '@azure/msal-browser';
import { Subject } from 'rxjs';
import { filter, takeUntil } from 'rxjs/operators';
import { LicensingService } from './services/licensing.service';
import { SharepointService } from './services/sharepoint.service';
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
    private readonly teams: TeamsService, 
    private readonly sharepoint: SharepointService, 
    private authService: MsalService, 
    private msalBroadcastService: MsalBroadcastService, 
    private licensing: LicensingService
  ) {

  }
  ngOnInit(): void {
    this.isIframe = window !== window.parent && !window.opener; // Remove this line to use Angular Universal
    this.setLoginDisplay();

    this.msalBroadcastService.msalSubject$
      .pipe(
        filter((msg: EventMessage) => msg.eventType === EventType.LOGIN_SUCCESS),
      )
      .subscribe(async (result: EventMessage) => {
        console.log('result login', result);
        const payload = result.payload as AuthenticationResult;
        await this.licensing.setJDLicense(payload.accessToken);
        console.log('isvalid', this.licensing.isValidJDLicense());
        console.log('sharepoint uri', this.licensing.getSharepointUri());
        if (this.licensing.isValidJDLicense()) {
          this.teams.setActiveAccount(payload.account);
          // this.teams.setToken(payload.accessToken);
        } else {
          console.log('NO VALID LICENSE');
        }
        
      });

    this.msalBroadcastService.inProgress$
      .pipe(
        filter((status: InteractionStatus) => status === InteractionStatus.None),
        takeUntil(this._destroying$)
      )
      .subscribe(() => {
        this.setLoginDisplay();
        this.teams.checkAndSetActiveAccount();
      })

    this.teams.checkAndSetActiveAccount();

  }

  setLoginDisplay() {
    this.loginDisplay = this.authService.instance.getAllAccounts().length > 0;
  }

  logout() {
    this.sharepoint.removeCurrentUserInfo(); // clean local storage
    this.teams.logout();
  }

  ngOnDestroy(): void {
    this._destroying$.next();
    this._destroying$.complete();
  }
}
