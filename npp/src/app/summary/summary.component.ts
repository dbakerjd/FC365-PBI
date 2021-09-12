import { Component, OnInit } from '@angular/core';
import { NPPNotification, SharepointService } from '../services/sharepoint.service';
import * as Highcharts from 'highcharts';

@Component({
  selector: 'app-summary',
  templateUrl: './summary.component.html',
  styleUrls: ['./summary.component.scss']
})
export class SummaryComponent implements OnInit {

  notificationsList: NPPNotification[] = [];
  projectsStats: {
    total: number,
    active: number,
    archived: number
  } | null = null;

  constructor(
    private sharepoint: SharepointService, 
  ) { }

  async ngOnInit(): Promise<void> {
    const user = await this.sharepoint.getCurrentUserInfo();
    this.notificationsList = await this.sharepoint.getUserNotifications(user.Id);

    const opportunities = await this.sharepoint.getOpportunities(false);
    this.projectsStats = {
      total: opportunities.length,
      active: opportunities.filter(o => o.OpportunityStatus === 'Active').length,
      archived: opportunities.filter(o => o.OpportunityStatus === 'Archive').length
    }

    let options = {
      chart: {
          plotShadow: true,
          type: 'pie'
      },
      title: {
          text: 'Current Project Stats: '+this.projectsStats.total+' Projects'
      },
      tooltip: {
          pointFormat: '{series.name}: <b>{point.value}</b>'
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
                  format: '<b>{point.name}</b>: {point.value}'
              }
          }
      },
      series: [{
          name: 'Projects',
          colorByPoint: true,
          data: [{
              name: 'Active',
              y: this.projectsStats.active * 100 / this.projectsStats.total,
              value: this.projectsStats.active,
              sliced: true,
              selected: true,
              color: '#ffa13e'
          }, {
              name: 'Archived',
              value: this.projectsStats.archived,
              y: this.projectsStats.archived * 100 / this.projectsStats.total,
          }]
      }]
    };

    //@ts-ignore
    Highcharts.chart('chart', options);
  }

}
