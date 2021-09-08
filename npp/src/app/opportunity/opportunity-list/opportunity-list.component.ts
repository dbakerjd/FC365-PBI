import { Component, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MatDialog } from '@angular/material/dialog';
import { Router } from '@angular/router';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { ToastrService } from 'ngx-toastr';
import { CreateOpportunityComponent } from 'src/app/modals/create-opportunity/create-opportunity.component';
import { Opportunity, SharepointService, User } from 'src/app/services/sharepoint.service';

@Component({
  selector: 'app-opportunity-list',
  templateUrl: './opportunity-list.component.html',
  styleUrls: ['./opportunity-list.component.scss']
})
export class OpportunityListComponent implements OnInit {
  currentUser: User | undefined = undefined;
  opportunities: Opportunity[] = [];
  form = new FormGroup({});
  model: any = { };
  fields: FormlyFieldConfig[] = [];
  dialogInstance: any;
  loading = true;

  constructor(
    private sharepoint: SharepointService, 
    private toastr: ToastrService,
    private router: Router, 
    public matDialog: MatDialog
    ) { }

  async ngOnInit() {

    /**TODEL */
    // this.sharepoint.testAddGroup();
    // this.sharepoint.testAddGroupToOpportunity();
    // this.sharepoint.testEndpoint();
    // this.sharepoint.getLists();

    // this.sharepoint.createStageActions(1, 1);
    // await this.sharepoint.createFolder('/6');
    // this.sharepoint.createOpportunityGroups(13, 5, 3);
    
    // const r = await this.sharepoint.getGroupPermissions("Opportunities");
    // console.log('r', r);

    // const all = await this.sharepoint.getGroupPermissions();
    // console.log('all', all);
    // console.log('filtered', all.filter(el => el.ListFilter === 'List'));

    const folder = await this.sharepoint.createFolder('/testfolder');
    console.log('folder', folder);

    console.log('master', await this.sharepoint.getStageFolders(3));
    this.sharepoint.readGroups();

    await this.sharepoint.getRoleDefinitionId('ListEdit');
    await this.sharepoint.getRoleDefinitionId('ListEdit');
    await this.sharepoint.getRoleDefinitionId('ListRead');
    /**TODEL */

    
    this.currentUser = await this.sharepoint.getCurrentUserInfo();
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

    this.dialogInstance.afterClosed().subscribe(async (result: any) => {
      if (result.success) {
        this.toastr.success("A opportunity was created successfully", result.data.opportunity.Title);
        let opp = await this.sharepoint.getOpportunity(result.data.opportunity.ID);
        this.opportunities.push(opp);
        this.sharepoint.initializeOpportunity(result.data.opportunity, result.data.stage).then(async r => {
          // set active
          await this.sharepoint.setOpportunityStatus(opp.ID, 'Active');
          opp.OpportunityStatus = 'Active';
          this.toastr.success("The opportunity is now active", opp.Title);
        });
      } else {
        this.toastr.error("The opportunity couldn't be created", "Try again");
      }
    });
  }

  async editOpportunity(opp: Opportunity) {
    this.dialogInstance = this.matDialog.open(CreateOpportunityComponent, {
      height: '700px',
      width: '405px',
      data: {
        opportunity: opp
      }
    });

    this.dialogInstance.afterClosed().subscribe(async (result: any) => {
      if (result.success) {
        this.toastr.success("The opportunity was updated successfully", result.data.Title);
        Object.assign(opp, await this.sharepoint.getOpportunity(opp.ID));
      } else {
        this.toastr.error("The opportunity couldn't be updated", "Try again");
      }
    });

  }

  onSubmit() {
    return; // filtering done with pipes
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

  async archiveOpportunity(opp: Opportunity) {
    console.log(opp);
    const success = await this.sharepoint.setOpportunityStatus(opp.ID, "Archive");
    if (success) {
      this.toastr.success("The opportunity has been archived");
    } else {
      this.toastr.error("The opportunity couldn't be archived", "Try again");
    }
  }

  async restoreOpportunity(opp: Opportunity) {
    console.log(opp);
    const success = await this.sharepoint.setOpportunityStatus(opp.ID, "Active");
    if (success) {
      this.toastr.success("The opportunity has been restored");
    } else {
      this.toastr.error("The opportunity couldn't be restored", "Try again");
    }
  }


}
