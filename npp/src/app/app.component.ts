import { Component } from '@angular/core';
import { MsalBroadcastService, MsalService } from '@azure/msal-angular';
import { AuthenticationResult, EventMessage, EventType } from '@azure/msal-browser';
import { Subject } from 'rxjs';
import { filter, takeUntil } from 'rxjs/operators';
import { LicensingService } from './services/licensing.service';
import { TeamsService } from './services/teams.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  private readonly _destroying$ = new Subject<void>();
  
  constructor(
    private readonly teams: TeamsService, 
    private authService: MsalService, 
    private msalBroadcastService: MsalBroadcastService, 
    private licensing: LicensingService
  ) {

  }
  ngOnInit(): void {
    this.authService.handleRedirectObservable().subscribe();

    this.msalBroadcastService.msalSubject$
      .pipe(
        filter((msg: EventMessage) => msg.eventType === EventType.LOGIN_SUCCESS),
        takeUntil(this._destroying$)
      )
      .subscribe((result: EventMessage) => {
        console.log(result);
        const payload = result.payload as AuthenticationResult;
        console.log('payload', payload);
        console.log("l", this.licensing.hasJplusDLicense(payload.accessToken));
        this.teams.setActiveAccount(payload.account);
        this.teams.setToken(payload.accessToken);
    });

    this.msalBroadcastService.msalSubject$
        .pipe(
            filter((msg: EventMessage) => msg.eventType === EventType.ACQUIRE_TOKEN_SUCCESS),
            takeUntil(this._destroying$)
        )
        .subscribe((result: EventMessage) => {
            const payload = result.payload as AuthenticationResult;
            this.teams.setActiveAccount(payload.account);
            this.teams.setToken(payload.accessToken);
        });

    this.teams.getActiveAccount();
    console.log("l", this.licensing.hasJplusDLicense( this.teams.token));

  }

  ngOnDestroy(): void {
    this._destroying$.next();
    this._destroying$.complete();
  }
}
