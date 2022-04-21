import { Component, OnInit } from '@angular/core';
import * as Highcharts from 'highcharts';
import { TeamsService } from '@services/microsoft-data/teams.service';
import { NotificationsService } from '../services/notifications.service';
import { NPPNotification } from '@shared/models/notification';
import { User } from '@shared/models/user';
import { Opportunity } from '@shared/models/entity';
import { AppDataService } from '../services/app/app-data.service';
import { AppControlService } from '@services/app/app-control.service';

@Component({
  selector: 'app-summary',
  templateUrl: './summary.component.html',
  styleUrls: ['./summary.component.scss']
})
export class SummaryComponent implements OnInit {

  currentUser: User | undefined = undefined;
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
    private notifications: NotificationsService,
    private teams: TeamsService,
    private readonly appData: AppDataService,
    private readonly appControl: AppControlService
  ) { }

  async ngOnInit(): Promise<void> {
    if(this.appControl.isReady) {
      this.init();
    }else {
      this.appControl.readySubscriptions.subscribe(val => {
        this.init();
      });
    }
    this.init();
    
  } 

  async init() {
    try {
      this.notificationsList = await this.notifications.getNotifications();
      this.opportunities = await this.appData.getAllOpportunities(true, true);
      const gates = await this.appData.getAllStages();

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
        let lastGateTasks = await this.appData.getActionsRaw(lastGate.EntityNameId, lastGate.StageNameId);
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
            plotBorderWidth: null,
            plotShadow: false,
            plotBorderColor: "#ff0000",
            backgroundColor: "#fff",
            type: 'pie'
        },
        title: {
            text: 'Current Gate: '+this.gateCount.Total+' Projects',
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
            text: 'Current Phase: '+this.phaseCount.Total+' Projects',
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
            text: 'Therapy Areas: '+this.therapyAreasData.total+' Projects',
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

      // seats
      this.currentUser = await this.appData.getCurrentUserInfo();
      if (this.currentUser.IsSiteAdmin) this.loadSeatsInfo();

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

  private async loadSeatsInfo() {
    /** seats */
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
    /** endseats */
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
