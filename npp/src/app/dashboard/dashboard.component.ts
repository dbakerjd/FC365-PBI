import { Component, OnInit } from '@angular/core';
import { TeamsService } from '../services/teams.service';
import { Router } from '@angular/router';
import { LicensingService } from '../services/licensing.service';
import { environment } from 'src/environments/environment';

@Component({
  selector: 'app-dashboard',
  templateUrl: './dashboard.component.html',
  styleUrls: ['./dashboard.component.scss']
})
export class DashboardComponent implements OnInit {
  account: any;
  version = environment.version;
  items = [{
    src: 'assets/npp-summary.svg',
    text: 'NPP Summary',
    route: ['summary']
  }, {
    src: 'assets/opportunities.svg',
    text: 'Your Opportunities',
    route: ['opportunities']
  }];

  powerBiItem = {
    src: 'assets/power-bi.svg',
    text: 'Analytics Report',
    route: ['power-bi']
  };

  constructor(private readonly teams: TeamsService, private router: Router, private licensing: LicensingService) { }

  ngOnInit(): void {
    if(this.licensing.license && this.licensing.license.HasPowerBi) {
      this.items.push(this.powerBiItem);
    }
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

