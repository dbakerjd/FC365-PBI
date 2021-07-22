import { Component, OnInit } from '@angular/core';
import { Router } from '@angular/router';
import { Opportunity, SharepointService } from 'src/app/services/sharepoint.service';

@Component({
  selector: 'app-opportunity-list',
  templateUrl: './opportunity-list.component.html',
  styleUrls: ['./opportunity-list.component.scss']
})
export class OpportunityListComponent implements OnInit {
  opportunities: Opportunity[] = [];
  constructor(private sharepoint: SharepointService, private router: Router) { }

  async ngOnInit() {
    this.opportunities = await this.sharepoint.getOpportunities();
    let lists = await this.sharepoint.getLists();
    console.log(lists);
  }

  navigateTo(item: Opportunity) {
    this.router.navigate(['opportunities', item.Id, 'actions']);
  }
}
