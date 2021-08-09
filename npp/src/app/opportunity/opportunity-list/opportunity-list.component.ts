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
  model = { };
  fields: FormlyFieldConfig[] = [];
  dialogInstance: any;

  constructor(private sharepoint: SharepointService, private router: Router, public matDialog: MatDialog) { }

  async ngOnInit() {

    let indications = await this.sharepoint.getIndications();
    let opportunityTypes = await this.sharepoint.getOpportunityTypes();
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
            { value: 'active', label: 'Active' },
            { value: 'archived', label: 'Archived' },
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

    let objOpportunities = await this.sharepoint.getOpportunities();
    console.log(objOpportunities);
    /*this.opportunities = opt.map(el => {
      console.log(el);
      return el;
    });
   /*let lists = await this.sharepoint.getLists();
    console.log(lists);*/
  }

  createOpportunity() {
    this.dialogInstance = this.matDialog.open(CreateOpportunityComponent, {
      height: '700px',
      width: '405px'
    })
  }

  onSubmit() {
    console.log(this.model);
  }

  navigateTo(item: Opportunity) {
    this.router.navigate(['opportunities', item.Id, 'actions']);
  }
}
