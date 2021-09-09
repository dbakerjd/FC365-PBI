import { Component, OnInit } from '@angular/core';
import { MatDialog } from '@angular/material/dialog';
import { ActivatedRoute, Router } from '@angular/router';
import { DatepickerOptions } from 'ng2-datepicker';
import { ToastrService } from 'ngx-toastr';
import { take } from 'rxjs/operators';
import { ConfirmDialogComponent } from 'src/app/modals/confirm-dialog/confirm-dialog.component';
import { CreateScenarioComponent } from 'src/app/modals/create-scenario/create-scenario.component';
import { SendForApprovalComponent } from 'src/app/modals/send-for-approval/send-for-approval.component';
import { ShareDocumentComponent } from 'src/app/modals/share-document/share-document.component';
import { StageSettingsComponent } from 'src/app/modals/stage-settings/stage-settings.component';
import { UploadFileComponent } from 'src/app/modals/upload-file/upload-file.component';
import { Action, Stage, NPPFile, NPPFolder, Opportunity, SharepointService, User } from 'src/app/services/sharepoint.service';

@Component({
  selector: 'app-actions-list',
  templateUrl: './actions-list.component.html',
  styleUrls: ['./actions-list.component.scss']
})
export class ActionsListComponent implements OnInit {
  currentUser: User | undefined = undefined;
  isOwner = false;
  gates: Stage[] = [];
  opportunityId = 0;
  opportunity: Opportunity | undefined = undefined;
  currentGate: Stage | undefined = undefined;
  lastStageId: number | undefined = undefined; // next stage button control
  currentActions: Action[] | undefined = undefined;
  currentGateProgress: number = 0;
  dateOptions: DatepickerOptions = {
    format: 'Y-M-d'
  };
  currentSection = 'actions';
  dateListener: any;
  currentFiles: NPPFile[] = [];
  currentFolders: NPPFolder[] = [];
  currentFolder: NPPFolder | undefined = undefined;
  currentFolderUri: string = '';
  displayingModels: boolean = false;
  dialogInstance: any; 
  loading = false;

  constructor(
    private readonly sharepoint: SharepointService, 
    private route: ActivatedRoute, 
    private router: Router,
    public matDialog: MatDialog,
    private toastr: ToastrService
    ) { }

  ngOnInit(): void {
    this.sharepoint.getTest();
    this.route.params.subscribe(async (params) => {
      if(params.id && params.id != this.opportunityId) {
        this.opportunityId = params.id;
        this.opportunity = await this.sharepoint.getOpportunity(params.id);
        if (!this.opportunity) {
          this.router.navigate(['notfound']);
        }
        this.currentUser = await this.sharepoint.getCurrentUserInfo();
        this.isOwner = this.currentUser.Id === this.opportunity.OpportunityOwnerId;

        if (this.opportunity.OpportunityOwner) {
          this.opportunity.OpportunityOwner.profilePicUrl = await this.sharepoint.getUserProfilePic(this.opportunity.OpportunityOwnerId);
        }
        this.gates = await this.sharepoint.getStages(params.id);
        this.gates.forEach(async (el, index) => {
          el.actions = await this.sharepoint.getActions(params.id, el.StageNameId);
          el.folders = await this.sharepoint.getStageFolders(el.StageNameId, this.opportunityId);
          // el.folders = await this.sharepoint.getSubfolders(`/${this.opportunityId}/${el.StageNameId}`);
          console.log('folders', el.folders);
          this.setStatus(el.actions);

          //set current gate
          if(index < (this.gates.length - 1)) {
            let uncompleted = el.actions.filter(a => !a.Complete);
            if(!this.currentGate && uncompleted && (uncompleted.length > 0)) {
              this.setGate(el.ID);
            } 
          } else {
            if(!this.currentGate) {
              this.setGate(el.ID);
            }
            this.lastStageId = el.ID;
          }

        });

        this.setUpDateListener();
      }
    });
  }

  async openUploadDialog() {
    this.dialogInstance = this.matDialog.open(UploadFileComponent, {
      height: '600px',
      width: '405px',
      data: {
        folderList: this.currentFolders,
        countries: await this.sharepoint.getCountriesList(),
        scenarios: await this.sharepoint.getScenariosList(),
        masterStageId: this.currentGate?.StageNameId,
        opportunityId: this.opportunityId
      }
    })
  }

  sendForApproval(file: NPPFile) {
    this.dialogInstance = this.matDialog.open(SendForApprovalComponent, {
      height: '300px',
      width: '405px',
      data: {
        file: file
      }
    })
  }

  createScenario(file: NPPFile) {
    this.dialogInstance = this.matDialog.open(CreateScenarioComponent, {
      height: '400px',
      width: '405px',
      data: {
        file: file
      }
    })
  }

  openStageSettings() {
    this.dialogInstance = this.matDialog.open(StageSettingsComponent, {
      height: '400px',
      width: '405px',
      data: {
        stage: this.currentGate
      }
    });

    this.dialogInstance.afterClosed().subscribe(async (result:any) => {
      console.log('result', result);
      if (this.currentGate && result.success) {
        this.currentGate.StageUsersId = result.data.StageUsersId;
        this.currentGate.StageReview = result.data.StageReview;
      }
    });
  }

  setUpDateListener() {
    this.dateListener = setInterval(()=>{
      if(this.currentActions) {
        this.setStatus(this.currentActions)
      };
    }, 1000);
  }

  setSection(section: string) {
    this.currentSection = section;
  }

  showFolders() {
    this.currentSection = 'documents';
    if (this.currentFolders.length > 0) this.setFolder(this.currentFolders[0].ID);
    else this.setFolder(null)
  }

  showModels() {
    this.setSection('documents');
    let modelsFolder = this.currentFolders.find(el => el.containsModels === true);
    if (modelsFolder)  this.setFolder(modelsFolder.ID);
  }

  setStatus(actions: Action[]) {
    actions.forEach(a => {
      this.computeStatus(a);
    });
  }

  computeStatus(a: Action) {
    let today = new Date().getTime();
    if (a.ActionDueDate) a.ActionDueDate = new Date(a.ActionDueDate); // set to date format for datepicker

    if(a.Complete) a.status = 'completed';
    else if (a.ActionDueDate) {
      let dueDate = a.ActionDueDate.getTime();
      if(dueDate < today) {
        a.status = 'late';
      } else {
        a.status = 'pending';
      }
    } else {
      a.status = 'pending';
    }
    this.computeProgress();
  }

  async toggleStatus(action: Action) {
    let done = false;
    if (!this.currentUser) this.currentUser = await this.sharepoint.getCurrentUserInfo(); // no tenim ID user al sharepoint
    
    if (action.Complete) done = await this.sharepoint.uncompleteAction(action.Id);
    else {
      done = await this.sharepoint.completeAction(action.Id, this.currentUser.Id);
    }

    if (done) {
      action.Complete = !action.Complete;
      this.computeStatus(action);
      if (action.Complete) {
        action.Timestamp = new Date();
        action.TargetUser = {
          Id: this.currentUser.Id,
          FirstName: this.currentUser.Title,
          LastName: ''
        };
      }
    }
  }

  onDueDateChange(actionId: number, value: string) {
    this.sharepoint.setActionDueDate(actionId, value);
  }

  async nextStage() {
    if (!this.currentGate) return;
    let nextStage = await this.sharepoint.getNextStage(this.currentGate.StageNameId);
    if (nextStage) {
      this.dialogInstance = this.matDialog.open(StageSettingsComponent, {
        height: '400px',
        width: '405px',
        data: {
          next: nextStage,
          opportunityId: this.opportunityId
        }
      });
      this.dialogInstance.afterClosed().subscribe(async (result: any) => {
        if (result.success) {
          let opp = await this.sharepoint.getOpportunity(result.data.OpportunityNameId);
          this.sharepoint.initializeStage(opp, result.data).then(async r => {
            // set active
            // await this.sharepoint.setOpportunityStatus(opp.ID, 'Active');
            // opp.OpportunityStatus = 'Active';
            this.toastr.success("Next stage has been created successfully", result.data.Title);
          });
        } else {
          this.toastr.error("The next stage couldn't be created", "Try again");
        }
      });
    }
    else {
      // TODO Complete Opportunity
    }
  }

  computeProgress() {
    if(this.currentActions && this.currentActions.length) {
      let completed = this.currentActions.filter(el => el.Complete);
      this.currentGateProgress = Math.round((completed.length / this.currentActions.length) * 10000) / 100;
    } else {
      this.currentGateProgress = 0;
    }
  }

  setGate(gateId: number) {
    let gate = this.gates.find(el => el.ID == gateId);
    if(gate && gate != this.currentGate) {
      this.currentGate = gate;
      this.currentActions = gate.actions;
      this.computeProgress();
      this.getFolders();
    } else if(gate && gate == this.currentGate) {
      if(this.displayingModels || this.currentSection == 'documents') {
        this.setSection('actions');
      } else {
        this.setSection('documents');
      }
    }
  }

  async getFolders() {
    if (!this.currentGate?.folders) {
      this.currentFolder = undefined;
      this.currentFiles = [];
    } else {
      this.currentFolders = this.currentGate.folders;
      if (this.currentFolders.length) this.setFolder(this.currentFolders[0].ID);
    }
  }

  async setFolder(folderId: number | null) {
    this.currentFiles = [];
    if (folderId) {
      this.loading = true;
      this.currentFolder = this.currentFolders.find(el => el.ID === folderId);
      this.currentFolderUri = `${this.opportunityId}/${this.currentGate?.StageNameId}/`+folderId;
      this.currentFiles = await this.sharepoint.readFolderFiles(this.currentFolderUri, true);
  
      this.displayingModels = false;
      if (this.currentFolder) {
        this.displayingModels = !!this.currentFolder.containsModels;
      }
      this.loading = false;
    } else {
      // no folders
      this.currentFolder = undefined;
      this.currentFolderUri = '';
      this.displayingModels = false;
    }

  }

  async openFile(fileId: number, forceDownload = false) {
    const fileInfo = this.currentFiles.find(f => f.ListItemAllFields?.ID === fileId);
    if (!fileInfo) return;


    const response = await this.sharepoint.readFile(fileInfo.ServerRelativeUrl);
    var newBlob = new Blob([response]);

    if (forceDownload) {
      var link = document.createElement('a');
      link.href = window.URL.createObjectURL(newBlob);
      link.download = fileInfo.Name;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } else {
      const data = window.URL.createObjectURL(newBlob);
      window.open(data);
    }
  }

  async shareFile(fileId: number, folderId: number) {
    const file = this.currentFiles.find(f => f.ListItemAllFields?.ID === fileId);
    if (!file) return;
    let folderGroup = `${this.opportunityId}-${this.currentGate?.StageNameId}-${folderId}`;
    folderGroup = 'Beta Test Group'; // TODO
    let folderUsersList = await this.sharepoint.getUsersFromGroup(folderGroup);
    this.matDialog.open(ShareDocumentComponent, {
      height: '250px',
      width: '405px',
      data: {
        file,
        folderUsersList
      }
    });
  }

  async deleteFile(fileId: number) {
    const fileInfo = this.currentFiles.find(f => f.ListItemAllFields?.ID === fileId);
    if (!fileInfo) return;

    const dialogRef = this.matDialog.open(ConfirmDialogComponent, {
      maxWidth: "400px",
      height: "200px",
      data: {
        message: 'Are you sure you want to delete the file <em>' + fileInfo.Name + '</em> ?',
        confirmButtonText: 'Yes, delete'
      }
    });

    dialogRef.afterClosed()
      .pipe(take(1))
      .subscribe(async deleteConfirmed => {
        if (deleteConfirmed) {
          if (await this.sharepoint.deleteFile(fileInfo.ServerRelativeUrl)) {
            // remove file for the current files list
            this.currentFiles = this.currentFiles.filter(f => f.ListItemAllFields?.ID !== fileId);
            console.log('File deleted');
          }
        }
      });
  }

  ngOnDestroy() {
    clearTimeout(this.dateListener);
  }
}
