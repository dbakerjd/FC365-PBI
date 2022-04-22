import { Component, OnInit } from '@angular/core';
import * as Highcharts from 'highcharts';
import { NPPNotification } from '@shared/models/notification';
import { User } from '@shared/models/user';
import { Opportunity } from '@shared/models/entity';
import { NotificationsService } from '@services/notifications.service';
import { AppControlService } from '@services/app/app-control.service';
import { EntitiesService } from '@services/entities.service';
import { ErrorService } from '@services/app/error.service';
import { PermissionsService } from '@services/permissions.service';

@Component({
  selector: 'app-summary',
  templateUrl: './summary.component.html',
  styleUrls: ['./summary.component.scss']
})
export class SummaryComponent implements OnInit {

  loadingGraphics = true;
  loadingTasksTable = true;
  currentUser: User | undefined = undefined;
  notificationsList: NPPNotification[] = [];
  gateProjects: Opportunity[] = [];
  phaseProjects: Opportunity[] = [];
  gateCount: any = {};
  phaseCount: any = {};
  therapyAreasData: any = {};
  currentTherapyArea: string = '';
  currentTasks: {
    opportunityName: string,
    taskName: string;
    taskDeadLine: Date | undefined;
  }[] = [];

  constructor(
    private notifications: NotificationsService,
    private readonly permissions: PermissionsService,
    private readonly entities: EntitiesService,
    private readonly appControl: AppControlService,
    private readonly error: ErrorService
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

  async ngAfterViewInit() {
    this.notifications.updateUnreadNotifications();
  }

  async init() {
    this.notificationsList = await this.notifications.getNotifications();
    this.currentUser = await this.permissions.getCurrentUserInfo();

    await this.prepareData();
    this.createGraphics();
    this.tasksTable();
  }

  private async prepareData() {
    const opportunities = await this.entities.getAll();
    const gates = await this.entities.getAllStages();

    this.therapyAreasData  = { areas: {}, total: 0 };
    this.currentTasks = [];
    for (const el of opportunities) {
      //populate gates/phases and isGateType
      let filteredGates = gates.filter(g => {
        return g.EntityNameId == el.ID;
      });
      el.gates = filteredGates;
      if (el.gates.length < 1) {
        continue;
      }
      if(el.gates.length > 0) {
        el.isGateType = el.gates[0].Title.indexOf('Gate') != -1;
      }
      
      this.populateTherapyAreasData(el);

      let lastGate = el.gates[el.gates.length - 1];
      let lastGateTasks = await this.entities.getStageActionsRaw(lastGate.EntityNameId, lastGate.StageNameId);
      let lastTask = lastGateTasks.find(el => !el.Complete);
      let taskInfo = {
        opportunityName: el.Title,
        taskName: lastGate.Title + (lastTask?.Title ? " - " + lastTask.Title : ''),
        taskDeadLine: lastTask?.ActionDueDate
      }
      this.currentTasks.push(taskInfo);
      
    }
    
    this.gateProjects = opportunities.filter(el => el.isGateType);
    this.phaseProjects = opportunities.filter(el => !el.isGateType);
  }

  private tasksTable() {
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
    });

    this.loadingTasksTable = false;
  }

  private createGraphics() {
    try {
      //@ts-ignore
      window.SummaryComponent = this;

      this.gateCount = { gates: {}, Total: 0 };
      this.gateProjects.forEach(p => {
        let numGates = p.gates?.length;
        if (numGates) {
          if (this.gateCount.gates["Gate " + numGates]) {
            this.gateCount.gates["Gate " + numGates] += 1;
          } else {
            this.gateCount.gates["Gate " + numGates] = 1;
          }

          this.gateCount.Total += 1;

        }
      });

      this.phaseCount = { phases: {}, Total: 0 };
      this.phaseProjects.forEach(p => {
        let numPhases = p.gates?.length;
        if (numPhases) {
          if (this.phaseCount.phases["Phase " + numPhases]) {
            this.phaseCount.phases["Phase " + numPhases] += 1;
          } else {
            this.phaseCount.phases["Phase " + numPhases] = 1;
          }

          this.phaseCount.Total += 1;
        }
      });

      let optionsGateProjects = {
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
          text: 'Current Gate: ' + this.gateCount.Total + ' Projects',
          style: {
            "fontSize": "1.2rem",
            "color": "#000"
          }
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
              format: '<b>{point.name}</b>: {point.value} projects',
              style: {
                "fontSize": "0.8rem",
                "color": "#333",
                "fontWeight": "normal"
              }
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
          plotBorderWidth: null,
          plotShadow: false,
          backgroundColor: "#fff",
          type: 'pie'
        },
        title: {
          text: 'Current Phase: ' + this.phaseCount.Total + ' Projects',
          style: {
            "fontSize": "1.2rem",
            "color": "#000"
          }
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
              format: '<b>{point.name}</b>: {point.value} projects',
              style: {
                "fontSize": "0.8rem",
                "color": "#333",
                "fontWeight": "normal"
              }
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
          plotBorderWidth: null,
          plotShadow: false,
          backgroundColor: "#fff",
          type: 'pie'
        },
        title: {
          text: 'Therapy Areas: ' + this.therapyAreasData.total + ' Projects',
          style: {
            "fontSize": "1.2rem",
            "color": "#000"
          }
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
              format: '<b>{point.name}</b>: {point.value} projects',
              style: {
                "fontSize": "0.8rem",
                "color": "#333",
                "fontWeight": "normal"
              }
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
              sliced: true
            }
          })
        }]
      };

      this.loadingGraphics = false;

      //@ts-ignore
      if (this.gateProjects.length) Highcharts.chart('chart', optionsGateProjects);
      //@ts-ignore
      if (this.phaseProjects.length) Highcharts.chart('chart-2', optionsPhaseProjects);
      //@ts-ignore
      if (Object.keys(this.therapyAreasData.areas).length) Highcharts.chart('chart-3', optionsTherapyAreas);

      if (this.currentTherapyArea) {
        if (Object.keys(this.therapyAreasData.areas).length) this.renderIndicationsGraph();
      }
      this.loadingGraphics = false;

    } catch (e) {
      this.error.handleError(e);
      this.loadingGraphics = false;
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
          plotBorderWidth: null,
          plotShadow: false,
          backgroundColor: "#fff",
          type: 'pie'
      },
      title: {
          text: 'Indications for '+self.currentTherapyArea+': '+self.therapyAreasData.areas[self.currentTherapyArea].count+' Projects',
          style: {
            "fontSize": "1.2rem",
            "color": "#000"
          }
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
                  format: '<b>{point.name}</b>: {point.value} projects',
                  style: {
                    "fontSize": "0.8rem",
                    "color": "#333",
                    "fontWeight": "normal"
                  }
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

  private populateTherapyAreasData(opp: Opportunity) {
    if (opp.Indication && opp.Indication.length) {
      for (let i = 0; i < opp.Indication.length; i++) {
        this.therapyAreasData.total += 1;
        let indication = opp.Indication[i];
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
