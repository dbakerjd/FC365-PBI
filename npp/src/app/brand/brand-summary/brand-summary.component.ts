import { Component, OnInit } from '@angular/core';
import { Brand, NPPNotification, SharepointService } from 'src/app/services/sharepoint.service';
import * as Highcharts from 'highcharts';
import { TeamsService } from 'src/app/services/teams.service';
import { InlineNppDisambiguationService } from 'src/app/services/inline-npp-disambiguation.service';

@Component({
  selector: 'app-brand-summary',
  templateUrl: './brand-summary.component.html',
  styleUrls: ['./brand-summary.component.scss']
})
export class BrandSummaryComponent implements OnInit {

  notificationsList: NPPNotification[] = [];
  therapyAreasData: any = {};
  currentTherapyArea: string = '';
  brands: Brand[] = [];
  brandData: {
    brandName: string,
    cycle: string,
    modelsCount: number,
    approvedModelsCount: number
  }[] = [];

  constructor(
    private sharepoint: SharepointService, 
    private teams: TeamsService,
    private disambiguator: InlineNppDisambiguationService
  ) { }

  async ngOnInit(): Promise<void> {
    try {
      if(this.teams.initialized) this.init();
      else {
        this.teams.statusSubject.subscribe(async (msg) => {
          setTimeout(async () => {
            this.init();
          }, 500);
        });
      }
    } catch(e) {
      console.log(e);
    }
  } 

  async init() {
    //@ts-ignore
    window.SummaryComponent = this;

    const user = await this.sharepoint.getCurrentUserInfo();
    this.notificationsList = await this.sharepoint.getUserNotifications(user.Id);
    this.therapyAreasData  = { areas: {}, total: 0 };

    this.brands = await this.disambiguator.getEntities() as Brand[];

    this.brands.forEach(async (el, index) => {
      
      //populate therapyAreasData
      if(el.Indication && el.Indication.length) {
        for(let i=0; i < el.Indication.length; i++) {
          this.therapyAreasData.total += 1;
          let indication = el.Indication[i];
          if(this.therapyAreasData.areas[indication.TherapyArea]) {
            this.therapyAreasData.areas[indication.TherapyArea].count += 1;
            if(this.therapyAreasData.areas[indication.TherapyArea].indications[indication.Title]) {
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
        brandName: el.Title,
        cycle: el.ForecastCycle?.Title + " " + el.Year,
        modelsCount: await this.sharepoint.getBrandModelsCount(el),
        approvedModelsCount: await this.sharepoint.getBrandApprovedModelsCount(el),
      });
      
    });
  }

  renderTherapyAreasGraph() {
    let optionsTherapyAreas = {
      credits: {
        enabled: false
      },
      chart: {
          plotShadow: true,
          backgroundColor: "#ebebeb",
          type: 'pie'
      },
      title: {
          text: 'Therapy Areas: '+this.therapyAreasData.total+' brands'
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
            if(!this.currentTherapyArea) this.currentTherapyArea = key;
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
    if(Object.keys(this.therapyAreasData.areas).length) Highcharts.chart('chartTherapyAreas', optionsTherapyAreas);  
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
          text: 'Indications for '+self.currentTherapyArea+': '+self.therapyAreasData.areas[self.currentTherapyArea].count+' brands'
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
          name: 'Indications for '+self.currentTherapyArea,
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
    if(Object.keys(self.therapyAreasData.areas).length) Highcharts.chart('chartIndications', optionsIndications);  
  }
}
