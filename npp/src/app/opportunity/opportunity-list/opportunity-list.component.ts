import { Component, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MatDialog } from '@angular/material/dialog';
import { Router } from '@angular/router';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { CreateOpportunityComponent } from 'src/app/modals/create-opportunity/create-opportunity.component';
import { Opportunity, SharepointService } from 'src/app/services/sharepoint.service';

@Component({
  selector: 'app-opportunity-list',
  templateUrl: './opportunity-list.component.html',
  styleUrls: ['./opportunity-list.component.scss']
})
export class OpportunityListComponent implements OnInit {
  opportunities: Opportunity[] = [];
  form = new FormGroup({});
  model: any = { };
  fields: FormlyFieldConfig[] = [];
  dialogInstance: any;
  loading = true;

  constructor(private sharepoint: SharepointService, private router: Router, public matDialog: MatDialog) { }

  async ngOnInit() {

    this.sharepoint.getLists();
    
    let indications = await this.sharepoint.getIndicationsList();
    let opportunityTypes = await this.sharepoint.getOpportunityTypesList();
    let opportunityFields = await this.sharepoint.getOpportunityFields();
    
    this.fields = [{
        key: 'search',
        type: 'input',
        templateOptions: {
          placeholder: 'Search all opportunities'
        }
      },{
        key: 'status',
        type: 'select',
        templateOptions: {
          placeholder: 'All',
          options: [
            { value: 'processing', label: 'Processing' },
            { value: 'active', label: 'Active' },
            { value: 'archived', label: 'Archived' },
            { value: 'approved', label: 'Approved' },
          ]
        }
      },{
        key: 'type',
        type: 'select',
        templateOptions: {
          placeholder: 'Filter by type',
          options: opportunityTypes
        }
      },{
        key: 'indication',
        type: 'select',
        templateOptions: {
          placeholder: 'Filter by indication',
          options: indications
        }
      },{
        key: 'sort_by',
        type: 'select',
        templateOptions: {
          placeholder: 'Sort by',
          options: opportunityFields
        }
      }
    ];

    this.opportunities = await this.sharepoint.getOpportunities();
    this.loading = false;
    for (let op of this.opportunities) {
      op.progress = await this.computeProgress(op);
    }
  }

  createOpportunity() {
    this.dialogInstance = this.matDialog.open(CreateOpportunityComponent, {
      height: '700px',
      width: '405px'
    });
  }

  onSubmit() {
    return;
  }

  navigateTo(item: Opportunity) {
    if (item.OpportunityStatus === "Processing") return;
    this.router.navigate(['opportunities', item.ID, 'actions']);
  }

  async computeProgress(opportunity: Opportunity): Promise<number> {
    let actions = await this.sharepoint.getActions(opportunity.ID);
    if (actions.length) {
      let gates: {'total': number; 'completed': number}[] = [];
      let currentGate = 0;
      let gateIndex = 0;
      for(let act of actions) {
        if (act.StageNameId == currentGate) {
          gates[gateIndex-1]['total']++;
          if (act.Complete) gates[gateIndex-1]['completed']++;
        } else {
          currentGate = act.StageNameId;
          if (act.Complete) gates[gateIndex] = {'total': 1, 'completed': 1};
          else gates[gateIndex] = {'total': 1, 'completed': 0};
          gateIndex++;
        }
      }

      let gatesMedium = gates.map(function(x) { return x.completed / x.total; });
      return Math.round((gatesMedium.reduce((a, b) => a + b, 0) / gatesMedium.length) * 10000) / 100;
    }
    return 0;
  }
}
