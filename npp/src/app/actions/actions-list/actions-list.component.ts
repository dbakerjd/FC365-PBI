import { Component, OnInit } from '@angular/core';
import { ActivatedRoute } from '@angular/router';
import { Gate, Opportunity, SharepointService } from 'src/app/services/sharepoint.service';

@Component({
  selector: 'app-actions-list',
  templateUrl: './actions-list.component.html',
  styleUrls: ['./actions-list.component.scss']
})
export class ActionsListComponent implements OnInit {
  gates: Gate[] = [];
  opportunityId = 0;
  opportunity: Opportunity | undefined = undefined;

  constructor(private sharepoint: SharepointService, private route: ActivatedRoute) { }

  ngOnInit(): void {
    this.route.params.subscribe(async (params) => {
      if(params.id && params.id != this.opportunityId) {
        this.opportunityId = params.id;
        this.opportunity = await this.sharepoint.getOpportunity(params.id);
        this.gates = await this.sharepoint.getGates(params.id);
        this.gates.forEach(async el => {
          el.actions = await this.sharepoint.getActions(el.id);
        });
      }
    });
  }

}
