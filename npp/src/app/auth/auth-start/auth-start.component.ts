import { Component, Inject, OnInit } from '@angular/core';
import { EventMessage, EventType, IPublicClientApplication, RedirectRequest } from '@azure/msal-browser';
import { ErrorService } from '@services/app/error.service';
import { TeamsService } from '@services/microsoft-data/teams.service';

@Component({
  selector: 'app-auth-start',
  templateUrl: './auth-start.component.html',
  styleUrls: ['./auth-start.component.scss']
})
export class AuthStartComponent implements OnInit {

  constructor(public teams: TeamsService, public error: ErrorService) {
    
  }

  ngOnInit(): void {
    let config = this.teams.getMSALGuardConfig();
    if (config && config.authRequest){
      this.teams.msalInstance.loginRedirect({...config.authRequest} as RedirectRequest);
    } else {
      this.teams.msalInstance.loginRedirect();
    }
  }

}
