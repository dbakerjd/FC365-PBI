import { Component, OnInit } from '@angular/core';
import { Opportunity, SharepointService } from 'src/app/services/sharepoint.service';

@Component({
  selector: 'app-opportunity-list',
  templateUrl: './opportunity-list.component.html',
  styleUrls: ['./opportunity-list.component.scss']
})
export class OpportunityListComponent implements OnInit {
  opportunities: Opportunity[] = [];
  constructor(private sharepoint: SharepointService) { }

  async ngOnInit() {
    this.opportunities = await this.sharepoint.getOpportunities();
  }

}
