import { Component, Inject, OnInit } from '@angular/core';
import { Router } from '@angular/router';
import { MsalBroadcastService, MsalService, MSAL_INSTANCE } from '@azure/msal-angular';
import { EventMessage, EventType, IPublicClientApplication } from '@azure/msal-browser';
import { filter } from 'rxjs/operators';
import * as microsoftTeams from "@microsoft/teams-js";

@Component({
  selector: 'app-auth-end',
  templateUrl: './auth-end.component.html',
  styleUrls: ['./auth-end.component.scss']
})
export class AuthEndComponent {
  
  constructor(@Inject(MSAL_INSTANCE) private msalInstance: IPublicClientApplication, public msalBroadcastService: MsalBroadcastService, public router: Router) { 
    msalInstance.handleRedirectPromise().then((tokenResponse) => {
      let res = JSON.stringify(tokenResponse);
      console.log(res);
      microsoftTeams.authentication.notifySuccess(res);

      setTimeout(() => {
        window.opener.jdTeamsHackMethod(res);
        window.close();
      }, 500);
    }).catch((error) => {
      console.log("Login failure");
      microsoftTeams.authentication.notifyFailure(JSON.stringify(error));
    });

    this.msalBroadcastService.msalSubject$
      .pipe(
        filter((msg: EventMessage) => msg.eventType === EventType.LOGIN_SUCCESS),
      )
      .subscribe(async (result: EventMessage) => {
        let res = JSON.stringify(result);
        console.log(res);
        microsoftTeams.authentication.notifySuccess(res);
        setTimeout(() => {
          window.opener.jdTeamsHackMethod(res);
          window.close();
        }, 500);
      });

      this.msalBroadcastService.msalSubject$
      .pipe(
        filter((msg: EventMessage) => msg.eventType === EventType.LOGIN_FAILURE),
      )
      .subscribe(async (result: EventMessage) => {
        console.log("Login failure");
        microsoftTeams.authentication.notifyFailure(JSON.stringify(result));
      });
  
      setTimeout(() => {
        console.log("Login failure TIMEOUT");
        microsoftTeams.authentication.notifyFailure("login timeout");
      }, 10000)
  }

  ngOnInit() {
    if(!window.opener) {
      this.router.navigate(['/']);
    }
  }
}
