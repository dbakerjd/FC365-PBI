import { Component, Inject } from '@angular/core';
import { Router } from '@angular/router';
import { EventMessage, EventType, IPublicClientApplication } from '@azure/msal-browser';
import { filter } from 'rxjs/operators';
import * as microsoftTeams from "@microsoft/teams-js";

@Component({
  selector: 'app-auth-end',
  templateUrl: './auth-end.component.html',
  styleUrls: ['./auth-end.component.scss']
})
export class AuthEndComponent {
  
  constructor(public router: Router) { 
   
  }

  ngOnInit() {
    //magic is handled directly on teamsService
    if(!window.opener) {
      this.router.navigate(['/']);
    }
  }
}
