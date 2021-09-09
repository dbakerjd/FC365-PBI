import { Component, OnInit } from '@angular/core';
import { NPPNotification, SharepointService } from '../services/sharepoint.service';

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
  }

}
