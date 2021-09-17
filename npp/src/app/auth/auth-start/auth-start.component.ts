import { Component, Inject, OnInit } from '@angular/core';
import { MsalGuardConfiguration, MsalService, MSAL_GUARD_CONFIG } from '@azure/msal-angular';
import { RedirectRequest } from '@azure/msal-browser';

@Component({
  selector: 'app-auth-start',
  templateUrl: './auth-start.component.html',
  styleUrls: ['./auth-start.component.scss']
})
export class AuthStartComponent implements OnInit {

  constructor(@Inject(MSAL_GUARD_CONFIG) private msalGuardConfig: MsalGuardConfiguration, private authService: MsalService) { }

  ngOnInit(): void {
    if (this.msalGuardConfig.authRequest){
      this.authService.loginRedirect({...this.msalGuardConfig.authRequest} as RedirectRequest);
    } else {
      this.authService.loginRedirect();
    }
  }

}
