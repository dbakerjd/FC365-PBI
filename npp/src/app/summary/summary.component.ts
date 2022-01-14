import { Component, OnInit } from '@angular/core';
import { NPPNotification, Opportunity, SharepointService } from '../services/sharepoint.service';
import * as Highcharts from 'highcharts';
import { TeamsService } from '../services/teams.service';
import { NotificationsService } from '../services/notifications.service';

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
  opportunities: Opportunity[] = [];
  currentTasks: {
    opportunityName: string,
    taskName: string;
    taskDeadLine: Date | undefined;
  }[] = [];
  projectsStats: {
    total: number,
    active: number,
    archived: number
  } | null = null;

  constructor(
    private sharepoint: SharepointService, 
    private notifications: NotificationsService,
    private teams: TeamsService
  ) { }

  async ngOnInit(): Promise<void> {
    try {
      this.notificationsList = await this.notifications.getNotifications();

      this.opportunities = await this.sharepoint.getOpportunities(true, true);
      const gates = await this.sharepoint.getAllStages();

      this.therapyAreasData  = { areas: {}, total: 0 };
      this.currentTasks = [];
      this.opportunities.forEach(async (el, index) => {
        //populate gates/phases and isGateType
        let filteredGates = gates.filter(g => {
          return g.EntityNameId == el.ID;
        });
        el.gates = filteredGates;
        if(el.gates.length > 0) {
          el.isGateType = el.gates[0].Title.indexOf('Gate') != -1;
        }
        
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

        let lastGate = el.gates[el.gates.length - 1];
        let lastGateTasks = await this.sharepoint.getActionsRaw(lastGate.EntityNameId, lastGate.StageNameId);
        let lastTask = lastGateTasks.find(el => !el.Complete);
        let taskInfo = {
          opportunityName: el.Title,
          taskName: lastGate.Title+" - "+lastTask?.Title,
          taskDeadLine: lastTask?.ActionDueDate
        }
        this.currentTasks.push(taskInfo);

        this.currentTasks = this.currentTasks.sort((obj1, obj2) => {
          if(obj1.taskDeadLine && !obj2.taskDeadLine) {
            return 1;
          }

          if(!obj1.taskDeadLine && obj2.taskDeadLine) {
            return -1;
          }

          if(!obj1.taskDeadLine && !obj2.taskDeadLine) {
            return 0;
          }

          if(!obj1.taskDeadLine || !obj2.taskDeadLine) {
            return 0;
          }

          if (obj1.taskDeadLine > obj2.taskDeadLine) {
              return 1;
          }
      
          if (obj1.taskDeadLine < obj2.taskDeadLine) {
              return -1;
          }
      
          return 0;
        })
      });

      //@ts-ignore
      window.SummaryComponent = this;

      this.gateProjects = this.opportunities.filter(el => el.isGateType);
      this.phaseProjects = this.opportunities.filter(el => !el.isGateType);

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
        credits: {
          enabled: false
        },
        chart: {
            plotShadow: true,
            plotBorderColor: "#ff0000",
            backgroundColor: "#ebebeb",
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
        credits: {
          enabled: false
        },
        chart: {
            plotShadow: true,
            backgroundColor: "#ebebeb",
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
        credits: {
          enabled: false
        },
        chart: {
            plotShadow: true,
            backgroundColor: "#ebebeb",
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
        if(Object.keys(this.therapyAreasData.areas).length) this.renderIndicationsGraph();
      }

    } catch(e) {
      this.teams.hackyConsole += "********RUNTIME ERROR********    "+JSON.stringify(e);
    }
  } 

  async ngAfterViewInit() {
    this.notifications.updateUnreadNotifications();
  }

  async getOpportunityCurrentTaskName(op: Opportunity) {
    if(!op.gates) return;

    let currentGate = op.gates[op.gates.length - 1];


  }

  async getOpportunityCurrentTaskDeadline(op: Opportunity) {

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
          text: 'Indications for '+self.currentTherapyArea+': '+self.therapyAreasData.areas[self.currentTherapyArea].count+' Projects'
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
          name: 'Indications for '+self.currentTherapyArea,
          colorByPoint: true,
          data: Object.keys(self.therapyAreasData.areas[self.currentTherapyArea].indications).map(key => {
            return {
              name: key,
              y: self.therapyAreasData.areas[self.currentTherapyArea].indications[key] * 100 / self.therapyAreasData.areas[self.currentTherapyArea].count,
              value: self.therapyAreasData.areas[self.currentTherapyArea].indications[key],
              sliced: true
            }
          })
      }]
    };
    //@ts-ignore
    if(Object.keys(self.therapyAreasData.areas).length) Highcharts.chart('chart-4', optionsIndications);  
  }
}
