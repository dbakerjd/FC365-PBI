import { Component, OnInit } from '@angular/core';
import { MatDialog } from '@angular/material/dialog';
import { ActivatedRoute } from '@angular/router';
import { DatepickerOptions } from 'ng2-datepicker';
import { CreateScenarioComponent } from 'src/app/modals/create-scenario/create-scenario.component';
import { SendForApprovalComponent } from 'src/app/modals/send-for-approval/send-for-approval.component';
import { StageSettingsComponent } from 'src/app/modals/stage-settings/stage-settings.component';
import { UploadFileComponent } from 'src/app/modals/upload-file/upload-file.component';
import { Action, Gate, NPPFile, NPPFolder, Opportunity, SharepointService } from 'src/app/services/sharepoint.service';

@Component({
  selector: 'app-actions-list',
  templateUrl: './actions-list.component.html',
  styleUrls: ['./actions-list.component.scss']
})
export class ActionsListComponent implements OnInit {
  gates: Gate[] = [];
  opportunityId = 0;
  opportunity: Opportunity | undefined = undefined;
  currentGate: Gate | undefined = undefined;
  currentActions: Action[] | undefined = undefined;
  currentGateProgress: number = 0;
  dateOptions: DatepickerOptions = {
    format: 'Y-M-d'
  };
  currentSection = 'actions';
  dateListener: any;
  currentFiles: NPPFile[] = [];
  currentFolders: NPPFolder[] = [];
  currentFolder: number | undefined = undefined;
  displayingModels: boolean = false;
  uploadDialogInstance: any; 

  constructor(private sharepoint: SharepointService, private route: ActivatedRoute, public matDialog: MatDialog) { }

  ngOnInit(): void {
    this.route.params.subscribe(async (params) => {
      if(params.id && params.id != this.opportunityId) {
        this.opportunityId = params.id;
        this.opportunity = await this.sharepoint.getOpportunity(params.id);
        this.gates = await this.sharepoint.getGates(params.id);
        this.gates.forEach(async (el, index) => {
          
          el.actions = await this.sharepoint.getActions(el.ID);
          el.folders = await this.sharepoint.folders;
          this.setStatus(el.actions);

          //set current gate
          if(index < (this.gates.length - 1)) {
            let uncompleted = el.actions.filter(a => !a.completed);
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

  openUploadDialog() {
    this.uploadDialogInstance = this.matDialog.open(UploadFileComponent, {
      height: '600px',
      width: '405px',
      data: {
        folderList: this.currentFolders
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
    this.setFolder(this.currentFolders[0].id);
  }
  showModels() {
    this.setSection('documents');
    this.setFolder(this.currentFolders[this.currentFolders.length - 1].id);
  }

  setStatus(actions: Action[]) {
    actions.forEach(a => {
      this.computeStatus(a);
    });
  }

  computeStatus(a: Action) {
    let today = new Date().getTime();

    if(a.completed) a.status = 'completed';
    else if (a.ActionDueDate) {
      a.ActionDueDate = new Date(a.ActionDueDate);
      let dueDate = new Date(a.ActionDueDate).getTime();
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

  toggleStatus(action: Action) {
    action.completed = !action.completed;
    this.computeStatus(action);

  }

  computeProgress() {
    if(this.currentActions && this.currentActions.length) {
      let completed = this.currentActions.filter(el => el.completed);
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

  getFolders() {
    this.currentFolders = this.sharepoint.folders;
    if(this.currentFolders && 
      this.currentFolders.length) {
        this.setFolder(this.currentFolders[0].id);
    } else {
      this.currentFolder = undefined;
      this.currentFiles = [];
    }
  }

  async setFolder(folderId: number) {
    this.currentFolder = folderId;
    this.currentFiles = await this.sharepoint.getFiles(folderId);

    let folder = this.currentFolders.find(el => el.id === folderId);
    this.displayingModels = false;
    if(folder) {
      this.displayingModels = !!folder.containsModels;
    }
  }

  ngOnDestroy() {
    clearTimeout(this.dateListener);
  }
}
