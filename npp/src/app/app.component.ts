import { Component, OnDestroy, OnInit } from '@angular/core';
import { ActivatedRoute, Router } from '@angular/router';
import { Subject } from 'rxjs';
import { AppDataService } from './services/app/app-data.service';
import { LicensingService } from './services/jd-data/licensing.service';
import { TeamsService } from '@services/microsoft-data/teams.service';

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
    public teams: TeamsService, 
    private readonly appData: AppDataService, 
    public licensing: LicensingService,
    private router: Router, 
  ) {
    
  }
  async ngOnInit() {
    this.isIframe = window !== window.parent && !window.opener; // Remove this line to use Angular Universal
    
    this.teams.statusSubject.subscribe(async (msg) => {
      if(msg == 'license') {
        this.setLoginDisplay();
    
        if(window.location.href.indexOf("auth") == -1) {
          await this.teams.checkAndSetActiveAccount();
        }

      }
    })
    

  }

  setLoginDisplay() {
    this.loginDisplay = this.teams.msalInstance.getAllAccounts().length > 0;
  }

  logout() {
    this.appData.removeCurrentUserInfo(); // clean local storage
    this.teams.logout();
  }

  ngOnDestroy(): void {
    this._destroying$.next();
    this._destroying$.complete();
  }
}
