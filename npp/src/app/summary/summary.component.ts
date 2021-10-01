import { Component, OnInit } from '@angular/core';
import { NPPNotification, Opportunity, SharepointService } from '../services/sharepoint.service';
import * as Highcharts from 'highcharts';
import { TeamsService } from '../services/teams.service';

@Component({
  selector: 'app-summary',
  templateUrl: './summary.component.html',
  styleUrls: ['./summary.component.scss']
})
export class SummaryComponent implements OnInit {

  notificationsList: NPPNotification[] = [];
  gateProjects: Opportunity[] = [];
  phaseProjects: Opportunity[] = [];
  gateCount: any = {};
  phaseCount: any = {};
  therapyAreasData: any = {};
  currentTherapyArea: string = '';
  projectsStats: {
    total: number,
    active: number,
    archived: number
  } | null = null;

  constructor(
    private sharepoint: SharepointService, 
    private teams: TeamsService
  ) { }

  async ngOnInit(): Promise<void> {
    try {
      const user = await this.sharepoint.getCurrentUserInfo();
      this.notificationsList = await this.sharepoint.getUserNotifications(user.Id);

      const opportunities = await this.sharepoint.getOpportunities(true, true);
      const gates = await this.sharepoint.getAllStages();

      this.therapyAreasData  = { areas: {}, total: 0 };

      opportunities.forEach(el => {
        //populate gates/phases and isGateType
        let filteredGates = gates.filter(g => {
          return g.OpportunityNameId == el.ID;
        });
        el.gates = filteredGates;
        if(el.gates.length > 0) {
          el.isGateType = el.gates[0].Title.indexOf('Gate') != -1;
        }

        //populate therapyAreasData
        if(el.Indication && el.Indication.TherapyArea) {
          this.therapyAreasData.total += 1;
          if(this.therapyAreasData.areas[el.Indication.TherapyArea]) {
            this.therapyAreasData.areas[el.Indication.TherapyArea].count += 1;
            if(this.therapyAreasData.areas[el.Indication.TherapyArea].indications[el.Indication.Title]) {
              this.therapyAreasData.areas[el.Indication.TherapyArea].indications[el.Indication.Title] += 1;
            } else {
              this.therapyAreasData.areas[el.Indication.TherapyArea].indications[el.Indication.Title] = 1;
            }
          } else {
            this.therapyAreasData.areas[el.Indication.TherapyArea] = {
              count: 1,
              indications: {}
            };
            this.therapyAreasData.areas[el.Indication.TherapyArea].indications[el.Indication.Title] = 1;
          }
        }
        
      });

      
      this.teams.hackyConsole += "       **********Therapy Areas*************      "+JSON.stringify(this.therapyAreasData)+ "        **********************     ";

      this.gateProjects = opportunities.filter(el => el.isGateType);
      this.phaseProjects = opportunities.filter(el => !el.isGateType);

      this.gateCount = {gates: {}, Total: 0};
      this.gateProjects.forEach(p => {
        let numGates = p.gates?.length;
        if(numGates) {
          if(this.gateCount.gates["Gate "+numGates]) {
            this.gateCount.gates["Gate "+numGates] += 1;
          } else {
            this.gateCount.gates["Gate "+numGates] = 1;
          }

          this.gateCount.Total += 1;
          
        }
      });

      this.phaseCount = {phases: {}, Total: 0};
      this.phaseProjects.forEach(p => {
        let numPhases = p.gates?.length;
        if(numPhases) {
          if(this.phaseCount.phases["Phase "+numPhases]) {
            this.phaseCount.phases["Phase "+numPhases] += 1;
          } else {
            this.phaseCount.phases["Phase "+numPhases] = 1;
          }

          this.phaseCount.Total += 1;
        }
      });

      let optionsGateProjects = {
        chart: {
            plotShadow: true,
            type: 'pie'
        },
        title: {
            text: 'Current Gate: '+this.gateCount.Total+' Projects'
        },
        tooltip: {
            pointFormat: '{series.name}: <b>{point.value} projects</b>'
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
                    format: '<b>{point.name}</b>: {point.value} projects'
                }
            }
        },
        series: [{
            name: 'Current Gate',
            colorByPoint: true,
            data: Object.keys(this.gateCount.gates).map(key => {
              return {
                name: key,
                y: this.gateCount.gates[key] * 100 / this.gateCount.Total,
                value: this.gateCount.gates[key],
                sliced: true
              }
            })
        }]
      };

      let optionsPhaseProjects = {
        chart: {
            plotShadow: true,
            type: 'pie'
        },
        title: {
            text: 'Current Phase: '+this.phaseCount.Total+' Projects'
        },
        tooltip: {
            pointFormat: '{series.name}: <b>{point.value} projects</b>'
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
                    format: '<b>{point.name}</b>: {point.value} projects'
                }
            }
        },
        series: [{
            name: 'Current Phase',
            colorByPoint: true,
            data: Object.keys(this.phaseCount.phases).map(key => {
              return {
                name: key,
                y: this.phaseCount.phases[key] * 100 / this.phaseCount.Total,
                value: this.phaseCount.phases[key],
                sliced: true
              }
            })
        }]
      };

      let optionsTherapyAreas = {
        chart: {
            plotShadow: true,
            type: 'pie'
        },
        title: {
            text: 'Therapy Areas: '+this.therapyAreasData.total+' Projects'
        },
        tooltip: {
            pointFormat: '{series.name}: <b>{point.value} projects</b>'
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
                    format: '<b>{point.name}</b>: {point.value} projects'
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
                sliced: true
              }
            })
        }]
      };

      
      //@ts-ignore
      if(this.gateProjects.length) Highcharts.chart('chart', optionsGateProjects);
      //@ts-ignore
      if(this.phaseProjects.length) Highcharts.chart('chart-2', optionsPhaseProjects);
      //@ts-ignore
      
      if(Object.keys(this.therapyAreasData.areas).length) Highcharts.chart('chart-3', optionsTherapyAreas);
      
      if(this.currentTherapyArea) {
        this.currentTherapyArea = 'Haematology';
        let optionsIndications = {
          chart: {
              plotShadow: true,
              type: 'pie'
          },
          title: {
              text: 'Indications for '+this.currentTherapyArea+': '+this.therapyAreasData.areas[this.currentTherapyArea].count+' Projects'
          },
          tooltip: {
              pointFormat: '{series.name}: <b>{point.value} projects</b>'
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
                      format: '<b>{point.name}</b>: {point.value} projects'
                  }
              }
          },
          series: [{
              name: 'Indications for '+this.currentTherapyArea,
              colorByPoint: true,
              data: Object.keys(this.therapyAreasData.areas[this.currentTherapyArea].indications).map(key => {
                return {
                  name: key,
                  y: this.therapyAreasData.areas[this.currentTherapyArea].indications[key] * 100 / this.therapyAreasData.areas[this.currentTherapyArea].count,
                  value: this.therapyAreasData.areas[this.currentTherapyArea].indications[key],
                  sliced: true
                }
              })
          }]
        };
        //@ts-ignore
        if(Object.keys(this.therapyAreasData.areas).length) Highcharts.chart('chart-4', optionsIndications);  
      }

    } catch(e) {
      this.teams.hackyConsole += "********RUNTIME ERROR********    "+JSON.stringify(e);
    }
  } 
}
