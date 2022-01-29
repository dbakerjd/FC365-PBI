import { Component, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MatDialog } from '@angular/material/dialog';
import { Router } from '@angular/router';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { ToastrService } from 'ngx-toastr';
import { take } from 'rxjs/operators';
import { ConfirmDialogComponent } from 'src/app/modals/confirm-dialog/confirm-dialog.component';
import { CreateOpportunityComponent } from 'src/app/modals/create-opportunity/create-opportunity.component';
import { Opportunity, OpportunityType, SharepointService, User } from 'src/app/services/sharepoint.service';
import { NotificationsService } from 'src/app/services/notifications.service';
import { TeamsService } from 'src/app/services/teams.service';
import { WorkInProgressService } from 'src/app/services/work-in-progress.service';
import { InlineNppDisambiguationService } from 'src/app/services/inline-npp-disambiguation.service';

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
  opportunityTypes: OpportunityType[] = [];

  constructor(
    private sharepoint: SharepointService, 
    private notifications: NotificationsService,
    private toastr: ToastrService,
    private router: Router, 
    public matDialog: MatDialog,
    public jobs: WorkInProgressService,
    public teams: TeamsService,
    public disambiguator: InlineNppDisambiguationService
    ) { }

  async ngOnInit() {
    if(this.disambiguator.isReady) {
      this.init();
    }else {
      this.disambiguator.readySubscriptions.subscribe(val => {
        this.init();
      });
    }
  }

  async init() {
    this.currentUser = await this.sharepoint.getCurrentUserInfo();
    let indications = await this.sharepoint.getIndicationsList();
    this.opportunityTypes = await this.sharepoint.getOpportunityTypes();
    let opportunityTypes = this.opportunityTypes.map(t => { return { value: t.ID, label: t.Title } });
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
            // { value: 'Processing', label: 'Processing' },
            { value: 'Active', label: 'Active' },
            { value: 'Archive', label: 'Archived' },
            { value: 'Approved', label: 'Approved' },
          ],        
        },
        defaultValue: 'Active'
      },{
        key: 'type',
        type: 'select',
        templateOptions: {
          placeholder: 'All types',
          options: opportunityTypes
        }
      },{
        key: 'indication',
        type: 'select',
        templateOptions: {
          placeholder: 'All indications',
          options: indications,
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
    this.opportunities.forEach(el => {
      this.initIndicationString(el);
    })
    this.loading = false;
    for (let op of this.opportunities) {
      op.progress = await this.computeProgress(op);
    }
  }

  initIndicationString(el: Opportunity) {
    if(el.Indication && el.Indication.length) {
      (el as any).IndicationAsString = el.Indication.map(el => el.Title).join(", ");
      (el as any).TherapyAreaAsString = el.Indication.map(el => el.TherapyArea).join(", ");
    }
  }

  createOpportunity(fromOpp: Opportunity | null, forceType = false) {
    this.dialogInstance = this.matDialog.open(CreateOpportunityComponent, {
      height: '75vh',
      width: '500px',
      data: {
        opportunity: fromOpp ? fromOpp : null,
        createFrom: fromOpp ? true : false,
        forceType
      },
      panelClass: 'config-dialog-container'
    });

    this.dialogInstance.afterClosed()
    .pipe(take(1))
    .subscribe(async (result: any) => {
     
      if (result.success) {
        this.toastr.success("A opportunity was created successfully", result.data.opportunity.Title);
        let opp = await this.sharepoint.getOpportunity(result.data.opportunity.ID);
        opp.progress = 0;
        let job = this.jobs.startJob(
          "initialize opportunity "+result.data.opportunity.id,
          'The new opportunity is being initialized. Stages and permissions are being created.'
          );
        this.sharepoint.initializeOpportunity(result.data.opportunity, result.data.stage).then(async r => {
          // set active
          await this.sharepoint.setOpportunityStatus(opp.ID, 'Active');
          opp.OpportunityStatus = 'Active';
          this.initIndicationString(opp);
          this.opportunities = [...this.opportunities, opp];
          this.jobs.finishJob(job.id);
          this.toastr.success("The opportunity is now active", opp.Title);
          await this.notifications.opportunityOwnerNotification(result.data.opportunity);
          if(result.data.stage) await this.notifications.newOpportunityAccessNotification(result.data.stage.StageUsersId, result.data.opportunity);
        }).catch(e => {
          this.jobs.finishJob(job.id);
          this.toastr.error((e as Error).message);
        });
      } else if (result.success === false) {
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
      },
      panelClass: 'config-dialog-container'
    });

    this.dialogInstance.afterClosed()
    .pipe(take(1))
    .subscribe(async (result: any) => {
      if (result.success) {
        this.toastr.success("The opportunity was updated successfully", result.data.Title);
        if (opp.EntityOwnerId !== result.data.EntityOwnerId) {
          await this.notifications.opportunityOwnerNotification(result.data);
        }
        Object.assign(opp, await this.sharepoint.getOpportunity(opp.ID));
      } else if (result.success === false) {
        this.toastr.error("The opportunity couldn't be updated", "Try again");
      }
    });
  }

  onSubmit() {
    return; // filtering done with pipes
  }

  navigateTo(item: Opportunity) {
    if (item.OpportunityStatus === "Processing") return;
    let opType = this.opportunityTypes.find(el => el.Title == item.OpportunityType?.Title);
    if(opType?.isInternal) {
      this.router.navigate(['opportunities', item.ID, 'files']);
    } else {
      this.router.navigate(['opportunities', item.ID, 'actions']);
    }
    
  }

  async computeProgress(opportunity: Opportunity): Promise<number> {
    let opType = this.opportunityTypes.find(el => el.Title == opportunity.OpportunityType?.Title);
    if(opType?.isInternal) {
      return -1; // progress no applies
    }
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
    const success = await this.sharepoint.setOpportunityStatus(opp.ID, "Archive");
    if (success) {
      opp.OpportunityStatus = 'Archive';
      this.toastr.success("The opportunity has been archived");
    } else {
      this.toastr.error("The opportunity couldn't be archived", "Try again");
    }
  }

  async restoreOpportunity(opp: Opportunity) {
    const dialogRef = this.matDialog.open(ConfirmDialogComponent, {
      maxWidth: "400px",
      height: "200px",
      data: {
        message: `The opportunity <em>${opp.Title}</em> will be restored.<br />Do you want to create a new copy of the opportunity to do so?`,
        confirmButtonText: 'Yes, create a new one',
        cancelButtonText: 'No, restore only'
      }
    });

    dialogRef.afterClosed()
      .pipe(take(1))
      .subscribe(async response => {
        if (response) {
          // create new
          this.createOpportunity(opp);
        } else if (response === false) {
          const success = await this.sharepoint.setOpportunityStatus(opp.ID, "Active");
          if (success) {
            opp.OpportunityStatus = 'Active';
            this.toastr.success("The opportunity has been restored");
          } else {
            this.toastr.error("The opportunity couldn't be restored", "Try again");
          }
        }
      });
  }


}
