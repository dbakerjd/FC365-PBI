import { Component, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { SharepointService } from 'src/app/services/sharepoint.service';

@Component({
  selector: 'app-opportunity-filter',
  templateUrl: './opportunity-filter.component.html',
  styleUrls: ['./opportunity-filter.component.scss']
})
export class OpportunityFilterComponent implements OnInit {
  
  constructor(private sharepoint: SharepointService) { }

  form = new FormGroup({});
  model = { };
  fields: FormlyFieldConfig[] = [];

  onSubmit() {
    console.log(this.model);
  }

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
  }

}
