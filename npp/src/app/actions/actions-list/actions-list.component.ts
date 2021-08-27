import { Component, OnInit } from '@angular/core';
import { MatDialog } from '@angular/material/dialog';
import { ActivatedRoute, Router } from '@angular/router';
import { DatepickerOptions } from 'ng2-datepicker';
import { take } from 'rxjs/operators';
import { ConfirmDialogComponent } from 'src/app/modals/confirm-dialog/confirm-dialog.component';
import { CreateScenarioComponent } from 'src/app/modals/create-scenario/create-scenario.component';
import { SendForApprovalComponent } from 'src/app/modals/send-for-approval/send-for-approval.component';
import { StageSettingsComponent } from 'src/app/modals/stage-settings/stage-settings.component';
import { UploadFileComponent } from 'src/app/modals/upload-file/upload-file.component';
import { Action, Stage, NPPFile, NPPFolder, Opportunity, SharepointService } from 'src/app/services/sharepoint.service';

@Component({
  selector: 'app-actions-list',
  templateUrl: './actions-list.component.html',
  styleUrls: ['./actions-list.component.scss']
})
export class ActionsListComponent implements OnInit {
  gates: Stage[] = [];
  opportunityId = 0;
  opportunity: Opportunity | undefined = undefined;
  currentGate: Stage | undefined = undefined;
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
  uploadDialogInstance: any; 
  loading = false;

  constructor(
    private readonly sharepoint: SharepointService, 
    private route: ActivatedRoute, 
    private router: Router,
    public matDialog: MatDialog
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
        if (this.opportunity.OpportunityOwner) {
          this.opportunity.OpportunityOwner.profilePicUrl = await this.sharepoint.getUserProfilePic(this.opportunity.OpportunityOwnerId);
        }
        this.gates = await this.sharepoint.getStages(params.id);
        this.gates.forEach(async (el, index) => {
          el.actions = await this.sharepoint.getActions(params.id, el.StageNameId);
          el.folders = await this.sharepoint.getFolders(el.StageNameId);
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
          }

        });

        this.setUpDateListener();
      }
    });
  }

  async openUploadDialog() {
    this.uploadDialogInstance = this.matDialog.open(UploadFileComponent, {
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
    this.uploadDialogInstance = this.matDialog.open(SendForApprovalComponent, {
      height: '300px',
      width: '405px',
      data: {
        file: file
      }
    })
  }

  createScenario(file: NPPFile) {
    this.uploadDialogInstance = this.matDialog.open(CreateScenarioComponent, {
      height: '400px',
      width: '405px',
      data: {
        file: file
      }
    })
  }

  openStageSettings() {
    this.uploadDialogInstance = this.matDialog.open(StageSettingsComponent, {
      height: '400px',
      width: '405px',
      data: {
        gate: this.currentGate
      }
    })
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
    this.setFolder(this.currentFolders[0].ID);
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
    let currentUser = null;
    if (action.Complete) done = await this.sharepoint.uncompleteAction(action.Id);
    else {
      currentUser = await this.sharepoint.getCurrentUserInfo(); // no tenim ID user al sharepoint
      done = await this.sharepoint.completeAction(action.Id, currentUser.Id);
    }

    if (done) {
      action.Complete = !action.Complete;
      this.computeStatus(action);
      if (action.Complete) {
        action.Timestamp = new Date();
        action.TargetUser = {
          ID: currentUser.Id,
          FirstName: currentUser.Title,
          LastName: ''
        };
      }
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

  async setFolder(folderId: number) {
    this.currentFiles = [];
    this.loading = true;
    this.currentFolder = this.currentFolders.find(el => el.ID === folderId);
    this.currentFolderUri = `${this.opportunityId}/${this.currentGate?.StageNameId}/`+folderId;
    this.currentFiles = await this.sharepoint.readFolderFiles(this.currentFolderUri, true);

    this.displayingModels = false;
    if (this.currentFolder) {
      this.displayingModels = !!this.currentFolder.containsModels;
    }
    this.loading = false;
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
