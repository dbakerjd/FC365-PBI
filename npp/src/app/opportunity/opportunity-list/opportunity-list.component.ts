import { Component, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MatDialog } from '@angular/material/dialog';
import { Router } from '@angular/router';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { ToastrService } from 'ngx-toastr';
import { take } from 'rxjs/operators';
import { ConfirmDialogComponent } from 'src/app/modals/confirm-dialog/confirm-dialog.component';
import { CreateOpportunityComponent } from 'src/app/modals/create-opportunity/create-opportunity.component';
import { NotificationsService } from 'src/app/services/notifications.service';
import { WorkInProgressService } from '@services/app/work-in-progress.service';
import { AppControlService } from '@services/app/app-control.service';
import { User } from '@shared/models/user';
import { Opportunity, OpportunityType } from '@shared/models/entity';
import { AppDataService } from '@services/app/app-data.service';
import { PermissionsService } from 'src/app/services/permissions.service';
import { EntitiesService } from '@services/entities.service';
import { SelectListsService } from '@services/select-lists.service';

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
  canCreate = false;

  constructor(
    private permissions: PermissionsService, 
    private notifications: NotificationsService,
    private toastr: ToastrService,
    private router: Router, 
    public matDialog: MatDialog,
    public jobs: WorkInProgressService,
    private readonly appControl: AppControlService,
    private readonly appData: AppDataService,
    private readonly entities: EntitiesService,
    private readonly selectLists: SelectListsService
    ) { }

  async ngOnInit() {
    if(this.appControl.isReady) {
      this.init();
    }else {
      this.appControl.readySubscriptions.subscribe(val => {
        this.init();
      });
    }
  }

  async init() {
    this.currentUser = await this.appData.getCurrentUserInfo();
    this.canCreate = this.appControl.getAppConfigValue('AllowCreation') && !!this.currentUser?.IsSiteAdmin;

    let indications = await this.selectLists.getIndicationsList();
    let opportunityFields = await this.selectLists.getOpportunityFilterFields();
    
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
          options: await this.selectLists.getOpportunityTypesList()
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

    this.opportunities = await this.entities.getAll();
    this.opportunities.forEach(el => {
      this.initIndicationString(el);
    })
    this.loading = false;
    for (let op of this.opportunities) {
      op.progress = await this.entities.getProgress(op);
    }

    // this.appData.deleteAllGroups();
    // this.appData.removeUserSeat('demouser@jdforecasting.com');
    
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
        this.toastr.success("An opportunity was created successfully", result.data.opportunity.Title);
        let opp = await this.appData.getEntity(result.data.opportunity.ID);
        opp.progress = 0;
        if (await this.entities.isInternalOpportunity(opp.OpportunityTypeId)) {
          opp.progress = -1;
        }
        let job = this.jobs.startJob(
          "initialize opportunity "+result.data.opportunity.id
          );
        this.permissions.initializeOpportunity(result.data.opportunity, result.data.stage, result.users).then(async r => {
          if (r) {
            await this.entities.activeEntity(opp.ID);
            opp.OpportunityStatus = 'Active';
            this.initIndicationString(opp);
            this.opportunities = [...this.opportunities, opp];
            this.jobs.finishJob(job.id);
            this.toastr.success("The opportunity is now active", opp.Title);
            await this.notifications.opportunityOwnerNotification(result.data.opportunity);
            if(result.data.stage) await this.notifications.newOpportunityAccessNotification(result.users, result.data.opportunity);
          } else {
            this.appData.deleteOpportunity(opp.ID);
            this.jobs.finishJob(job.id);
            this.toastr.error("The opportunity couldn't be created", "Try again");
          }
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
      height: '75vh',
      width: '500px',
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
        Object.assign(opp, await this.appData.getEntity(opp.ID));
      } else if (result.success === false) {
        this.toastr.error("The opportunity couldn't be updated", "Try again");
      }
    });
  }

  onSubmit() {
    return; // filtering done with pipes
  }

  async navigateTo(item: Opportunity) {
    if (item.OpportunityStatus === "Processing" || !item.OpportunityTypeId) return;
    if(await this.entities.isInternalOpportunity(item.OpportunityTypeId)) {
      this.router.navigate(['opportunities', item.ID, 'files']);
    } else {
      this.router.navigate(['opportunities', item.ID, 'actions']);
    }
    
  }

  async archiveOpportunity(opp: Opportunity) {
    const success = await this.entities.archiveEntity(opp.ID);
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
          const success = await this.entities.activeEntity(opp.ID);
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
