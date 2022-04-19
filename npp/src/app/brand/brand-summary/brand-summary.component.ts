import { Component, OnInit } from '@angular/core';
import * as Highcharts from 'highcharts';
import { User } from '@shared/models/user';
import { Opportunity } from '@shared/models/entity';
import { NPPNotification } from '@shared/models/notification';
import { EntitiesService } from 'src/app/services/entities.service';
import { AppDataService } from '@services/app/app-data.service';
import { TeamsService } from '@services/microsoft-data/teams.service';

@Component({
  selector: 'app-brand-summary',
  templateUrl: './brand-summary.component.html',
  styleUrls: ['./brand-summary.component.scss']
})
export class BrandSummaryComponent implements OnInit {

  notificationsList: NPPNotification[] = [];
  therapyAreasData: any = {};
  currentUser: User | undefined = undefined;
  currentTherapyArea: string = '';
  brands: Opportunity[] = [];
  brandData: {
    Id: number,
    brandName: string,
    cycle: string,
    modelsCount: number,
    approvedModelsCount: number
  }[] = [];

  usersList: User[] = [];
  usersOpportunitiesListItem: { type: string | null, userId: number | null, list: Opportunity[] } = {
    type: null,
    userId: null,
    list: []
  };
  generalSeatsCount: {
    TotalSeats: number,
    AssignedSeats: number,
    AvailableSeats: number
  } | null = null;
  generatingSeatsTable = true;

  constructor(
    // private teams: TeamsService,
    private readonly entities: EntitiesService,
    private readonly appData: AppDataService
  ) { }

  async ngOnInit(): Promise<void> {
    // try {
    //   if (this.teams.initialized) this.init();
    //   else {
    //     this.teams.statusSubject.subscribe(async (msg) => {
    //       setTimeout(async () => {
    //         this.init();
    //       }, 500);
    //     });
    //   }
    // } catch (e) {
    //   console.log(e);
    // }
    this.init();
  }

  async init() {
    //@ts-ignore
    window.SummaryComponent = this;

    const user = await this.appData.getCurrentUserInfo();
    this.notificationsList = await this.appData.getUserNotifications(user.Id);
    this.therapyAreasData = { areas: {}, total: 0 };

    this.brands = await this.entities.getAll();

    this.brands.forEach(async (el, index) => {

      //populate therapyAreasData
      if (el.Indication && el.Indication.length) {
        for (let i = 0; i < el.Indication.length; i++) {
          this.therapyAreasData.total += 1;
          let indication = el.Indication[i];
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
    });

    this.renderTherapyAreasGraph();

    this.brands.forEach(async (el, index) => {

      this.brandData.push({
        Id: el.ID,
        brandName: el.Title,
        cycle: el.ForecastCycle?.Title + " " + el.Year,
        modelsCount: await this.entities.getModelsCount(el),
        approvedModelsCount: await this.entities.getApprovedModelsCount(el),
      });

    });

    // seats
    this.currentUser = await this.appData.getCurrentUserInfo();
    if (this.currentUser.IsSiteAdmin) this.loadSeatsInfo();
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
        text: 'Therapy Areas: ' + this.therapyAreasData.total + ' brands',
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
        name: 'Therapy Areas',
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

    //@ts-ignore
    if (Object.keys(this.therapyAreasData.areas).length) Highcharts.chart('chartTherapyAreas', optionsTherapyAreas);
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
        text: 'Indications for ' + self.currentTherapyArea + ': ' + self.therapyAreasData.areas[self.currentTherapyArea].count + ' brands'
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
        name: 'Indications for ' + self.currentTherapyArea,
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

  private async loadSeatsInfo() {
    this.generatingSeatsTable = true;
    this.usersList = await this.appData.getUsers();
    this.usersList = this.usersList.filter(el => el.Email);

    for (let index = 0; index < this.usersList.length; index++) {
      const user: any = this.usersList[index];
      const result = await this.appData.getSeats(user.Email);
      if (index == 0 && result) {
        this.generalSeatsCount = {
          AssignedSeats: result?.AssignedSeats,
          TotalSeats: result?.TotalSeats,
          AvailableSeats: result?.AvailableSeats
        }
      }
      user['seats'] = result?.UserGroupsCount;
      const groups = await this.appData.getUserGroups(user.Id);
      const OUgroups = groups.filter(g => g.Title.startsWith('OU-'));
      const OOgroups = groups.filter(g => g.Title.startsWith('OO-'));
      user['opportunities'] = OUgroups.length;
      user['owner'] = OOgroups.length;
    }

    this.generatingSeatsTable = false;
  }

  async listOpportunities(userId: number, group: 'OU' | 'OO') {
    if (this.usersOpportunitiesListItem.type == group && this.usersOpportunitiesListItem.userId == userId) {
      this.usersOpportunitiesListItem.type = null;
      this.usersOpportunitiesListItem.userId = null;
      this.usersOpportunitiesListItem.list = [];
      return;
    }
    const groups = await this.appData.getUserGroups(userId);
    const OUgroups = groups.filter(g => g.Title.startsWith(group + '-'));
    const allOpportunities = await this.appData.getAllOpportunities(false, false);
    const oppsList = OUgroups.map(e => {
      const splittedName = e.Title.split('-');
      return splittedName[1];
    });
    const oppsListRelated = allOpportunities.filter(opp => oppsList.includes(opp.ID.toString()));
    this.usersOpportunitiesListItem.type = group;
    this.usersOpportunitiesListItem.userId = userId;
    this.usersOpportunitiesListItem.list = oppsListRelated;
  }
}
