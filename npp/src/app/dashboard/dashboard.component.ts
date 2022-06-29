import { Component, OnInit } from '@angular/core';
import { TeamsService } from '@services/microsoft-data/teams.service';
import { Router } from '@angular/router';
import { LicensingService } from '../services/jd-data/licensing.service';
import { environment } from 'src/environments/environment';
import { animate, query, stagger, style, transition, trigger } from '@angular/animations';
import { Subject } from 'rxjs';
import { takeUntil } from 'rxjs/operators';
import { PermissionsService } from '@services/permissions.service';

@Component({
  selector: 'app-dashboard',
  templateUrl: './dashboard.component.html',
  styleUrls: ['./dashboard.component.scss'],
  animations: [
    trigger('listAnimation', [
      transition('* => *', [ 
        query(':enter', [
          style({ opacity: 0, marginTop: '1000px' }),
          stagger(200, [
            animate('2s', style({ opacity: 1, marginTop: '0px' }))
          ])
        ])
      ])
    ]),
  ],
})
export class DashboardComponent implements OnInit {
  account: any;
  version = environment.version;
  items: any[] = [];
  loadedApp = false;

  private readonly _destroying$ = new Subject<void>();

  constructor(
    private readonly teams: TeamsService, 
    private router: Router, 
    private licensing: LicensingService,
    private readonly permissions: PermissionsService
  ) { }

  async ngOnInit() {

    this.loadedApp = this.teams.isLoggedIn();
    if (this.loadedApp) {
      if (!await this.permissions.userHasAccessToEntities()) this.router.navigate(['power-bi']);
    } else {
      this.teams.statusSubject
      .pipe(takeUntil(this._destroying$))
      .subscribe(async (msg) => {
        if (msg == 'loggedIn') {
          if (!await this.permissions.userHasAccessToEntities()) {
            this.router.navigate(['power-bi']);
          } else {
            this.loadedApp = true;
          }
        }
      });
    }
    
    let NPPitems = [{
      src: 'assets/dashboard/npp-summary.svg',
      text: 'NPP Summary',
      description: 'An overview of all your active opportunities',
      route: ['summary']
    }, {
      src: 'assets/dashboard/opportunities.svg',
      text: 'Your Opportunities',
      description: 'See all the detail behind your opportunities and create additional opportunities',
      route: ['opportunities']
    }];

    let Inlineitems = [{
      src: 'assets/dashboard/npp-summary.svg',
      text: 'Inline Summary',
      description: 'An overview of all your active brands',
      route: ['brands-summary']
    }, {
      src: 'assets/dashboard/opportunities.svg',
      text: 'Your Brands',
      description: 'See all the detail behind your brands and create additional brands',
      route: ['brands']
    }];
  
    let powerBiItem = {
      src: 'assets/dashboard/analytics.svg',
      text: 'Analytics Reports',
      description: 'Explore your forecast outputs with powerful analytics and visual reports',
      route: ['power-bi']
    };

    if(environment.isInlineApp) {
      this.items = Inlineitems;
    } else {
      this.items = NPPitems;
    }
    
    if(this.licensing.license && this.licensing.license.HasPowerBi) {
      this.items.push(powerBiItem);
    }
  }

  getUser() {
    return this.teams.user;
  }

  getToken()  {
    return this.teams.token;
  }

  navigateTo(item: any) {
    this.router.navigate(item.route);
  }

  ngOnDestroy(): void {
    this._destroying$.next();
    this._destroying$.complete();
  }

}

