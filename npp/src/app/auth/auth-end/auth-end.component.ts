import { Component, OnInit } from '@angular/core';
import { MsalBroadcastService } from '@azure/msal-angular';
import { EventMessage, EventType } from '@azure/msal-browser';
import { filter } from 'rxjs/operators';

@Component({
  selector: 'app-auth-end',
  templateUrl: './auth-end.component.html',
  styleUrls: ['./auth-end.component.scss']
})
export class AuthEndComponent implements OnInit {
  success = false;
  constructor(private msalBroadcastService: MsalBroadcastService) { }

  ngOnInit(): void {
    this.msalBroadcastService.msalSubject$
      .pipe(
        filter((msg: EventMessage) => msg.eventType === EventType.LOGIN_SUCCESS),
      )
      .subscribe(async (result: EventMessage) => {
        this.success = true;
        microsoftTeams.authentication.notifySuccess(JSON.stringify(result));
      });

      this.msalBroadcastService.msalSubject$
      .pipe(
        filter((msg: EventMessage) => msg.eventType === EventType.LOGIN_FAILURE),
      )
      .subscribe(async (result: EventMessage) => {
        this.success = false;
        microsoftTeams.authentication.notifyFailure(JSON.stringify(result));
      });
    setTimeout(() => {
      microsoftTeams.authentication.notifyFailure("login timeout");
    }, 10000)
  }

}
