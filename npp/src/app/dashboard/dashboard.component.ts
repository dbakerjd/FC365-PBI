import { Component, OnInit } from '@angular/core';
import { TeamsService } from '../services/teams.service';
import { MsalBroadcastService, MsalService } from '@azure/msal-angular';
import { AuthenticationResult, EventMessage, EventType } from '@azure/msal-browser';
import { filter } from 'rxjs/operators';
import { Router } from '@angular/router';

@Component({
  selector: 'app-dashboard',
  templateUrl: './dashboard.component.html',
  styleUrls: ['./dashboard.component.scss']
})
export class DashboardComponent implements OnInit {
  account: any;
  items = [{
    src: 'assets/npp-summary.svg',
    text: 'NPP Summary',
    route: ['summary']
  }, {
    src: 'assets/opportunities.svg',
    text: 'Your Opportunities',
    route: ['opportunities']
  }, {
    src: 'assets/power-bi.svg',
    text: 'Power BI Report',
    route: ['power-bi']
  }];

  constructor(private readonly teams: TeamsService, private authService: MsalService, private msalBroadcastService: MsalBroadcastService, private router: Router) { }

  ngOnInit(): void {
    /*this.msalBroadcastService.msalSubject$
      .pipe(
        filter((msg: EventMessage) => msg.eventType === EventType.LOGIN_SUCCESS),
      )
      .subscribe((result: EventMessage) => {
        console.log(result);
        const payload = result.payload as AuthenticationResult;
        this.authService.instance.setActiveAccount(payload.account);
      });
    
    this.account = this.teams.getActiveAccount();*/
  }

  getUser() {
    return this.teams.user;
  }

  getContext()  {
    return this.teams.context;
  }

  getToken()  {
    return this.teams.token;
  }

  navigateTo(item: any) {
    this.router.navigate(item.route);
  }
}

