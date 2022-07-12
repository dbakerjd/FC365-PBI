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
import { StringMapperService } from '@services/string-mapper.service';
import { Router } from '@angular/router';

@Component({
  selector: 'app-summary',
  templateUrl: './summary.component.html',
  styleUrls: ['./summary.component.scss']
})
export class SummaryComponent implements OnInit {

  loadingGraphics = true;
  loadingTasksTable = true;
  seatsTableOption: 'All Users' | 'Admin Only' | 'Off' = 'All Users';
  currentUser: User | undefined = undefined;
  notificationsList: NPPNotification[] = [];
  gateProjects: Opportunity[] = [];
  allProjects: Opportunity[] = [];
  gateCount: any = {};
  // phaseCount: any = {};
  clinicalTrialPhaseCount: any = {};
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

  async ngAfterViewInit() {
    this.notifications.updateUnreadNotifications();
  }

  async init() {
    if (!await this.appControl.userHasAccessToEntities()) {
      this.router.navigate(['splash/reports']); return;
    }
    this.notificationsList = await this.notifications.getNotifications();
    this.currentUser = await this.permissions.getCurrentUserInfo();
    this.seatsTableOption = await this.appControl.getAppConfigValue('SeatsTable');

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
    this.allProjects = opportunities;
    // this.phaseProjects = opportunities.filter(el => !el.isGateType);
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

      this.clinicalTrialPhaseCount = { ctp: {}, Total: 0 };
      this.allProjects.forEach(p => {
        const clinicalTrialPhase = p.ClinicalTrialPhase;
        
        if (clinicalTrialPhase) {
          if (this.clinicalTrialPhaseCount.ctp[clinicalTrialPhase.Title]) this.clinicalTrialPhaseCount.ctp[clinicalTrialPhase.Title] += 1;
          else this.clinicalTrialPhaseCount.ctp[clinicalTrialPhase.Title] = 1;

          this.clinicalTrialPhaseCount.Total += 1;
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
          text: 'Current Phase: ' + this.clinicalTrialPhaseCount.Total + ' Projects',
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
          name: 'Clinical Trial Phases',
          colorByPoint: true,
          data: Object.keys(this.clinicalTrialPhaseCount.ctp).map(key => {
            return {
              name: key,
              y: this.clinicalTrialPhaseCount.ctp[key] * 100 / this.clinicalTrialPhaseCount.Total,
              value: this.clinicalTrialPhaseCount.ctp[key],
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
          text: this.stringMapper.getString('Therapy Areas') + ': ' + this.therapyAreasData.total + ' Projects',
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
          name: this.stringMapper.getString('Therapy Areas'),
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
      if (this.allProjects.length) Highcharts.chart('chart-2', optionsPhaseProjects);
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
          text: this.stringMapper.getString('Indications') + ' for '+self.currentTherapyArea+': '+self.therapyAreasData.areas[self.currentTherapyArea].count+' Projects',
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
          name: this.stringMapper.getString('Indications') + ' for '+self.currentTherapyArea,
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

  /** get the deadline date class */
  getDeadlineClass(date: Date) {
    const today = new Date();
    const nextWeek = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 7).getTime();
    const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, today.getDate()).getTime();
    const deadlineDate = new Date(date).getTime();
    if (deadlineDate < today.getTime() || deadlineDate < nextWeek) {
      return 'deadline late';
    } else if (deadlineDate < nextMonth) {
      return 'deadline soon';
    }
    return 'deadline';
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
