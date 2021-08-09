import { Component } from '@angular/core';
import { MsalBroadcastService, MsalService } from '@azure/msal-angular';
import { AuthenticationResult, EventMessage, EventType } from '@azure/msal-browser';
import { filter } from 'rxjs/operators';
import { TeamsService } from './services/teams.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  constructor(private readonly teams: TeamsService, private authService: MsalService, private msalBroadcastService: MsalBroadcastService, ) {

  }
  ngOnInit(): void {
    this.authService.handleRedirectObservable().subscribe();

    this.msalBroadcastService.msalSubject$
      .pipe(
        filter((msg: EventMessage) => msg.eventType === EventType.LOGIN_SUCCESS),
      )
      .subscribe((result: EventMessage) => {
        console.log(result);
        const payload = result.payload as AuthenticationResult;
        this.teams.setActiveAccount(payload.account);
        this.teams.setToken(payload.accessToken);
    });

    this.teams.getActiveAccount();
  }
}
