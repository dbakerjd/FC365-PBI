import { Component, OnInit } from '@angular/core';
import * as Highcharts from 'highcharts';
import { User } from '@shared/models/user';
import { Opportunity } from '@shared/models/entity';
import { NPPNotification } from '@shared/models/notification';
import { EntitiesService } from '@services/entities.service';
import { AppControlService } from '@services/app/app-control.service';
import { NotificationsService } from '@services/notifications.service';
import { ErrorService } from '@services/app/error.service';
import { PermissionsService } from '@services/permissions.service';
import { StringMapperService } from '@services/string-mapper.service';
import { Router } from '@angular/router';

@Component({
  selector: 'app-brand-summary',
  templateUrl: './brand-summary.component.html',
  styleUrls: ['./brand-summary.component.scss']
})
export class BrandSummaryComponent implements OnInit {

  loadingGraphics = true;
  loadingTable = true;
  notificationsList: NPPNotification[] = [];
  therapyAreasData: any = {};
  seatsTableOption: 'All Users' | 'Admin Only' | 'Off' = 'All Users';
  currentUser: User | undefined = undefined;
  currentTherapyArea: string = '';
  brandData: {
    Id: number,
    brandName: string,
    cycle: string,
    modelsCount: number,
    approvedModelsCount: number
  }[] = [];

  constructor(
    private notifications: NotificationsService,
    private readonly permissions: PermissionsService,
    private readonly entities: EntitiesService,
    private readonly appControl: AppControlService,
    private readonly error: ErrorService,
    private readonly stringMapper: StringMapperService,
    private router: Router
  ) { }

  async ngOnInit(): Promise<void> {
    if (this.appControl.isReady) {
      this.init();
    } else {
      this.appControl.readySubscriptions.subscribe(val => {
        this.init();
      });
    }
  }

  async init() {
    if (!await this.appControl.userHasAccessToEntities()) {
      this.router.navigate(['splash/reports']); return;
    }

    //@ts-ignore
    window.SummaryComponent = this;

    this.currentUser = await this.permissions.getCurrentUserInfo();
    this.notificationsList = await this.notifications.getNotifications();
    this.seatsTableOption = await this.appControl.getAppConfigValue('SeatsTable');
    this.therapyAreasData = { areas: {}, total: 0 };

    const brands = await this.entities.getAll();

    for (const el of brands) {
      this.populateTherapyAreasData(el);
    }

    this.renderTherapyAreasGraph();

    this.loadingGraphics = false;

    for (const el of brands) {
      this.brandData.push({
        Id: el.ID,
        brandName: el.Title,
        cycle: el.ForecastCycleDescriptor ? el.ForecastCycleDescriptor + " " + el.Year : el.Year.toString(),
        modelsCount: await this.entities.getModelsCount(el),
        approvedModelsCount: await this.entities.getApprovedModelsCount(el),
      });

    }
    this.loadingTable = false;

  }

  renderTherapyAreasGraph() {
    let optionsTherapyAreas = {
      credits: {
        enabled: false
      },
      chart: {
        plotBorderWidth: null,
        plotShadow: false,
        plotBorderColor: "#ff0000",
        backgroundColor: "#fff",
        type: 'pie'
      },
      title: {
        text: this.stringMapper.getString('Therapy Areas') + ': ' + this.therapyAreasData.total + ' brands',
        style: {
          "fontSize": "1.2rem",
          "color": "#000"
        }
      },
      tooltip: {
        pointFormat: '{series.name}: <b>{point.value} brands</b>'
      },
      accessibility: {
        point: {
          valueSuffix: '%'
        }
      },
      plotOptions: {
        pie: {
          allowPointSelect: true,
          cursor: 'pointer',
          dataLabels: {
            enabled: true,
            format: '<b>{point.name}</b>: {point.value} brands'
          }
        },
        series: {
          events: {
            click: function (event: any) {
              //@ts-ignore
              window.SummaryComponent.currentTherapyArea = event.point.name;
              //@ts-ignore
              window.SummaryComponent.renderIndicationsGraph();
            }
          }
        }
      },
      series: [{
        name: this.stringMapper.getString('Therapy Areas'),
        colorByPoint: true,
        data: Object.keys(this.therapyAreasData.areas).map(key => {
          if (!this.currentTherapyArea) this.currentTherapyArea = key;
          return {
            name: key,
            y: this.therapyAreasData.areas[key].count * 100 / this.therapyAreasData.total,
            value: this.therapyAreasData.areas[key].count,
            sliced: false
          }
        })
      }]
    };

    try {
      //@ts-ignore
      if (Object.keys(this.therapyAreasData.areas).length) Highcharts.chart('chartTherapyAreas', optionsTherapyAreas);
    } catch (e) {
      this.error.handleError(e);
    }
    
  }

  renderIndicationsGraph() {
    //@ts-ignore
    let self = window.SummaryComponent;
    let optionsIndications = {
      credits: {
        enabled: false
      },
      chart: {
        plotShadow: true,
        backgroundColor: "#ebebeb",
        type: 'pie'
      },
      title: {
        text: this.stringMapper.getString('Indications') + ' for ' + self.currentTherapyArea + ': ' + self.therapyAreasData.areas[self.currentTherapyArea].count + ' brands'
      },
      tooltip: {
        pointFormat: '{series.name}: <b>{point.value} brands</b>'
      },
      accessibility: {
        point: {
          valueSuffix: '%'
        }
      },
      plotOptions: {
        pie: {
          allowPointSelect: true,
          cursor: 'pointer',
          dataLabels: {
            enabled: true,
            format: '<b>{point.name}</b>: {point.value} brands'
          }
        }
      },
      series: [{
        name: this.stringMapper.getString('Indications') + ' for ' + self.currentTherapyArea,
        colorByPoint: true,
        data: Object.keys(self.therapyAreasData.areas[self.currentTherapyArea].indications).map(key => {
          return {
            name: key,
            y: self.therapyAreasData.areas[self.currentTherapyArea].indications[key] * 100 / self.therapyAreasData.areas[self.currentTherapyArea].count,
            value: self.therapyAreasData.areas[self.currentTherapyArea].indications[key],
            sliced: false
          }
        })
      }]
    };
    //@ts-ignore
    if (Object.keys(self.therapyAreasData.areas).length) Highcharts.chart('chartIndications', optionsIndications);
  }

  showSeatsTable() {
    switch(this.seatsTableOption) {
      case 'All Users':
        return true;
      case 'Admin Only':
        return this.currentUser?.IsSiteAdmin;
      case 'Off':
        return false;
      default:
        return false;
    }
  }

  private populateTherapyAreasData(b: Opportunity) {
    if (b.Indication && b.Indication.length) {
      for (let i = 0; i < b.Indication.length; i++) {
        this.therapyAreasData.total += 1;
        let indication = b.Indication[i];
        if (this.therapyAreasData.areas[indication.TherapyArea]) {
          this.therapyAreasData.areas[indication.TherapyArea].count += 1;
          if (this.therapyAreasData.areas[indication.TherapyArea].indications[indication.Title]) {
            this.therapyAreasData.areas[indication.TherapyArea].indications[indication.Title] += 1;
          } else {
            this.therapyAreasData.areas[indication.TherapyArea].indications[indication.Title] = 1;
          }
        } else {
          this.therapyAreasData.areas[indication.TherapyArea] = {
            count: 1,
            indications: {}
          };
          this.therapyAreasData.areas[indication.TherapyArea].indications[indication.Title] = 1;
        }
      }
    }
  }

}
