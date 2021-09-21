import { Component, Inject, OnInit } from '@angular/core';
import { MsalBroadcastService, MsalGuardConfiguration, MsalService, MSAL_GUARD_CONFIG, MSAL_INSTANCE } from '@azure/msal-angular';
import { EventMessage, EventType, IPublicClientApplication, RedirectRequest } from '@azure/msal-browser';
import { filter } from 'rxjs/operators';

@Component({
  selector: 'app-auth-start',
  templateUrl: './auth-start.component.html',
  styleUrls: ['./auth-start.component.scss']
})
export class AuthStartComponent implements OnInit {

  constructor(@Inject(MSAL_INSTANCE) private msalInstance: IPublicClientApplication, public msalBroadcastService: MsalBroadcastService, @Inject(MSAL_GUARD_CONFIG) private msalGuardConfig: MsalGuardConfiguration, private authService: MsalService) {
    
  }

  ngOnInit(): void {
    if (this.msalGuardConfig.authRequest){
      this.authService.loginRedirect({...this.msalGuardConfig.authRequest} as RedirectRequest);
    } else {
      this.authService.loginRedirect();
    }
  }

}
