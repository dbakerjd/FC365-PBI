import { HttpResponse } from '@angular/common/http';
import { Component, OnInit } from '@angular/core';
import { MatDialog } from '@angular/material/dialog';
import { ActivatedRoute, Router } from '@angular/router';
import { DatepickerOptions } from 'ng2-datepicker';
import { ToastrService } from 'ngx-toastr';
import { take } from 'rxjs/operators';
import { ConfirmDialogComponent } from 'src/app/modals/confirm-dialog/confirm-dialog.component';
import { CreateOpportunityComponent } from 'src/app/modals/create-opportunity/create-opportunity.component';
import { CreateScenarioComponent } from 'src/app/modals/create-scenario/create-scenario.component';
import { EditFileComponent } from 'src/app/modals/edit-file/edit-file.component';
import { FolderPermissionsComponent } from 'src/app/modals/folder-permissions/folder-permissions.component';
import { RejectModelComponent } from 'src/app/modals/reject-model/reject-model.component';
import { SendForApprovalComponent } from 'src/app/modals/send-for-approval/send-for-approval.component';
import { ShareDocumentComponent } from 'src/app/modals/share-document/share-document.component';
import { StageSettingsComponent } from 'src/app/modals/stage-settings/stage-settings.component';
import { UploadFileComponent } from 'src/app/modals/upload-file/upload-file.component';
import { LicensingService } from 'src/app/services/licensing.service';
import { NotificationsService } from 'src/app/services/notifications.service';
import { PowerBiService } from 'src/app/services/power-bi.service';
import { Action, Stage, NPPFile, NPPFolder, Opportunity, SharepointService, User, SelectInputList, FILES_FOLDER } from 'src/app/services/sharepoint.service';
import { WorkInProgressService } from 'src/app/services/work-in-progress.service';

@Component({
  selector: 'app-actions-list',
  templateUrl: './actions-list.component.html',
  styleUrls: ['./actions-list.component.scss']
})
export class ActionsListComponent implements OnInit {
  currentUser: User | undefined = undefined;
  alreadyGoingNextStage = false;
  isOwner = false;
  isStageUser = false;
  gates: Stage[] = [];
  opportunityId = 0;
  opportunity: Opportunity | undefined = undefined;
  currentGate: Stage | undefined = undefined;
  lastStageId: number | undefined = undefined; // next stage button control
  nextStage: Stage | null = null;
  currentActions: Action[] | undefined = undefined;
  currentGateProgress: number = 0;
  refreshingPowerBi = false;
  dateOptions: DatepickerOptions = {
    format: 'Y-M-d',
    firstCalendarDay: 1
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
  profilePic: string = '/assets/user.svg';

  constructor(
    private readonly sharepoint: SharepointService, 
    private readonly notifications: NotificationsService,
    private route: ActivatedRoute, 
    private router: Router,
    public matDialog: MatDialog,
    private toastr: ToastrService,
    public licensing: LicensingService,
    public jobs: WorkInProgressService,
    public powerBi: PowerBiService
    ) { }

  ngOnInit(): void {
    this.route.params.subscribe(async (params) => {
      if(params.id && params.id != this.opportunityId) {
        this.opportunityId = params.id;
        this.opportunity = await this.sharepoint.getOpportunity(params.id);
        if (!this.opportunity) {
          this.router.navigate(['notfound']);
        }
        this.currentUser = await this.sharepoint.getCurrentUserInfo();
        this.isOwner = this.currentUser.Id === this.opportunity.EntityOwnerId;

        if (this.opportunity.EntityOwner) {
          let pic = await this.sharepoint.getUserProfilePic(this.opportunity.EntityOwnerId);
          this.opportunity.EntityOwner.profilePicUrl = pic ? pic+'' : '/assets/user.svg';
          this.profilePic = this.opportunity.EntityOwner.profilePicUrl;
        }
        this.gates = await this.sharepoint.getStages(params.id);
        this.gates.forEach(async (el, index) => {
          el.actions = await this.sharepoint.getActions(params.id, el.StageNameId);
          el.folders = await this.sharepoint.getStageFolders(el.StageNameId, this.opportunityId, this.opportunity?.BusinessUnitId);
          // el.folders = await this.sharepoint.getSubfolders(`/${this.opportunityId}/${el.StageNameId}`);
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
            this.nextStage = await this.sharepoint.getNextStage(el.StageNameId);
          }

        });

        this.setUpDateListener();
      }
    });
  }

  async openUploadDialog() {
    if (!this.currentGate) return;

    let geographiesList: SelectInputList[] = [];
    const modelFolder = this.currentFolders.find(f => f.containsModels);
    if (modelFolder && this.opportunity) {
      geographiesList = await this.sharepoint.getAccessibleGeographiesList(
        this.opportunity.BusinessUnitId,
        this.opportunityId, 
        this.currentGate.StageNameId,
        modelFolder.ID
      )
    }
    this.dialogInstance = this.matDialog.open(UploadFileComponent, {
      height: '600px',
      width: '405px',
      data: {
        folderList: this.currentFolders,
        selectedFolder: this.currentSection === 'documents' && this.currentFolder ? this.currentFolder.ID : null,
        geographies: geographiesList,
        scenarios: await this.sharepoint.getScenariosList(),
        masterStageId: this.currentGate?.StageNameId,
        opportunityId: this.opportunityId,
        businessUnitId: this.opportunity?.BusinessUnitId
      }
    });

    this.dialogInstance.afterClosed()
    .pipe(take(1))
    .subscribe(async (result: any) => {
      if (result.success && result.uploaded) {
        this.toastr.success(`The file ${result.name} was uploaded successfully`, "File Uploaded");
        if (this.currentFolder?.containsModels) {
          const geoFolders = await this.sharepoint.getSubfolders(this.currentFolderUri);
          this.currentFiles = [];
          for (const geofolder of geoFolders) {
            this.currentFiles.push(...await this.sharepoint.readFolderFiles(this.currentFolderUri + '/' + geofolder.Name+'/0', true));
          }
        } else {
          this.currentFiles = await this.sharepoint.readFolderFiles(this.currentFolderUri+'/0/0', true);
        }
      } else if (result.success === false) {
        this.toastr.error("Sorry, there was a problem uploading your file");
      }
    });

    
  }

  async sendForApproval(file: NPPFile, departmentId: number) {
    this.dialogInstance = this.matDialog.open(SendForApprovalComponent, {
      height: '300px',
      width: '405px',
      data: {
        file,
        rootFolder: FILES_FOLDER,
        entity: this.opportunity
      }
    });

    this.dialogInstance.afterClosed()
      .pipe(take(1))
      .subscribe(async (result: any) => {
        if (result.success) {
          // update view
          if (file.ListItemAllFields?.ApprovalStatus?.Title) {
            file.ListItemAllFields.ApprovalStatus.Title = 'Submitted';
            if (result.comments) {
              file.ListItemAllFields.ModelApprovalComments = result.comments;
            }
          }

          //generate notifications
          this.toastr.success("The model has been sent for approval", "Forecast Model");
          await this.notifications.modelSubmittedNotification(file.Name, this.opportunityId, [
            `DU-${this.opportunityId}-${departmentId}-${file.ListItemAllFields?.EntityGeographyId}`,
            `OO-${this.opportunityId}`,
            `SU-${this.opportunityId}-${this.currentGate?.StageNameId}`,
          ]);
        } else if (result.success === false) {
          this.toastr.error("The model couldn't be sent for approval");
        }
      });
  }

  async approveModel(file: NPPFile, departmentId: number) {
    if (!file.ListItemAllFields) return;
    if (await this.sharepoint.setApprovalStatus(file.ListItemAllFields.ID, "Approved")) {
      file.ListItemAllFields.ApprovalStatus.Title = 'Approved';
      this.toastr.success("The model " + file.Name + " has been approved!", "Forecast Model");
      await this.notifications.modelApprovedNotification(file.Name, this.opportunityId, [
        `DU-${this.opportunityId}-${departmentId}-${file.ListItemAllFields?.EntityGeographyId}`,
        `OO-${this.opportunityId}`,
        `SU-${this.opportunityId}-${this.currentGate?.StageNameId}`,
      ]);
    } else {
      this.toastr.error("There were a problem approving the forecast model", 'Try again');
    }
  }

  async rejectModel(file: NPPFile, departmentId: number) {
    if (!file.ListItemAllFields) return;
    this.dialogInstance = this.matDialog.open(RejectModelComponent, {
      height: '300px',
      width: '405px',
      data: {
        file,
        rootFolder: FILES_FOLDER,
        entity: this.opportunity
      }
    });

    this.dialogInstance.afterClosed()
      .pipe(take(1))
      .subscribe(async (result: any) => {
        if (result.success) {
          // update view
          if (file.ListItemAllFields?.ApprovalStatus?.Title) {
            file.ListItemAllFields.ApprovalStatus.Title = 'In Progress';
            if (result.comments) {
              file.ListItemAllFields.ModelApprovalComments = result.comments;
            }
          }
          this.toastr.warning("The model " + file.Name + " has been rejected", "Forecast Model");
          await this.notifications.modelRejectedNotification(file.Name, this.opportunityId, [
            `DU-${this.opportunityId}-${departmentId}-${file.ListItemAllFields?.EntityGeographyId}`,
            `OO-${this.opportunityId}`,
            `SU-${this.opportunityId}-${this.currentGate?.StageNameId}`,
          ]);
        } else if (result.success === false) {
          this.toastr.error("There were a problem rejecting the forecast model", 'Try again');
        }
      });
  }

  createScenario(file: NPPFile) {
    this.dialogInstance = this.matDialog.open(CreateScenarioComponent, {
      height: '400px',
      width: '405px',
      data: {
        file: file
      }
    });

    this.dialogInstance.afterClosed()
      .pipe(take(1))
      .subscribe(async (success: any) => {
        if (success === true) {
          this.toastr.success(`The new model scenario has been created successfully`, "New Forecast Model");
          const geoFolders = await this.sharepoint.getSubfolders(this.currentFolderUri);
          this.currentFiles = [];
          for (const geofolder of geoFolders) {
            this.currentFiles.push(...await this.sharepoint.readFolderFiles(this.currentFolderUri + '/' + geofolder.Name + '/0', true));
          }
        } else if (success === false) {
          this.toastr.error('The new model scenario could not be created', 'Try Again');
        }
      });
  }

  openStageSettings() {
    this.dialogInstance = this.matDialog.open(StageSettingsComponent, {
      height: '400px',
      width: '405px',
      data: {
        stage: this.currentGate,
        canSetUsers: this.isOwner || this.currentUser?.IsSiteAdmin // only until set permission problem is resolved
      },
      panelClass: 'config-dialog-container'
    });

    this.dialogInstance.afterClosed()
      .pipe(take(1))
      .subscribe(async (result: any) => {
        if (this.currentGate && result.success) {
          // notification to new users
          const currentStageUsers = this.currentGate.StageUsersId;
          const addedStageUsers = result.data.StageUsersId.filter((item: number) => currentStageUsers.indexOf(item) < 0);
          await this.notifications.stageAccessNotification(addedStageUsers, this.currentGate.Title, this.opportunity?.Title);
          // update current info
          this.currentGate.StageUsersId = result.data.StageUsersId;
          this.currentGate.StageReview = result.data.StageReview;
        }
      });
  }

  openFolderPermissions() {
    if (this.isOwner || this.currentUser?.IsSiteAdmin) { // TODO: open to all stage users when using API
      this.dialogInstance = this.matDialog.open(FolderPermissionsComponent, {
        height: '400px',
        width: '405px',
        data: {
          folderList: this.currentFolders,
          opportunityId: this.opportunity?.ID,
          stageId: this.currentGate?.StageNameId
        }
      });
    }
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
    if (modelsFolder) this.setFolder(modelsFolder.ID);
  }

  setStatus(actions: Action[]) {
    actions.forEach(a => {
      this.computeStatus(a);
    });
  }

  computeStatus(a: Action) {
    let today = new Date().getTime();
    if (a.ActionDueDate) a.ActionDueDate = new Date(a.ActionDueDate); // set to date format for datepicker

    if (a.Complete) a.status = 'completed';
    else if (a.ActionDueDate) {
      let dueDate = a.ActionDueDate.getTime();
      if (dueDate < today) {
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
    // only if is the active stage
    if (this.opportunity?.OpportunityStatus !== 'Active' || !this.isActiveStage(action.StageNameId)) return;

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

  async goNextStage() {
    if (!this.currentGate || this.alreadyGoingNextStage) return;

    if (this.nextStage) {
      this.dialogInstance = this.matDialog.open(StageSettingsComponent, {
        height: '400px',
        width: '405px',
        data: {
          next: this.nextStage,
          opportunityId: this.opportunityId,
          canSetUsers: this.isOwner || this.currentUser?.IsSiteAdmin // only until set permission problem is resolved
        },
        panelClass: 'config-dialog-container'
      });
      this.dialogInstance.afterClosed()
        .pipe(take(1))
        .subscribe(async (result: any) => {
          if (result.success) {
            let job = this.jobs.startJob(
              'initialize stage ' + result.data.ID, 
              'The stage is being initialized. The list of actions and starter permissions are being created.'
            );
            let opp = await this.sharepoint.getOpportunity(result.data.EntityNameId);
            const oppGeographies = await this.sharepoint.getOpportunityGeographies(opp.ID);
            this.alreadyGoingNextStage = true;
            this.sharepoint.initializeStage(opp, result.data,oppGeographies).then(async r => {
              await this.jobs.finishJob(job.id);
              this.toastr.success("Next stage has been created successfully", result.data.Title);
              this.alreadyGoingNextStage = false;
              setTimeout(() => {
                window.location.reload();
              }, 1000);
            }).catch(e => {
              this.alreadyGoingNextStage = false;
              this.jobs.finishJob(job.id);
            });
          } else if (result.success === false) {
            this.toastr.error("The next stage couldn't be created", "Try again");
          }
        });
    }
  }

  async completeOpportunity() {
    // Complete Opportunity
    if (!this.opportunity) return;

    const dialogRef = this.matDialog.open(ConfirmDialogComponent, {
      maxWidth: "400px",
      height: "200px",
      data: {
        message: `Do you want to complete the opportunity <em>${this.opportunity.Title}</em>?`,
        confirmButtonText: 'Yes, complete',
      }
    });

    dialogRef.afterClosed()
      .pipe(take(1))
      .subscribe(async response => {
        if (response) {
          // complete opportunity
          if (!this.opportunity) return;

          const stageType = await this.sharepoint.getStageType(this.opportunity.OpportunityTypeId);
          if (await this.sharepoint.getStageType(this.opportunity.OpportunityTypeId) !== 'Phase') {
            const newPhaseDialog = this.matDialog.open(ConfirmDialogComponent, {
              maxWidth: "400px",
              height: "200px",
              data: {
                message: `You can move to a Phase process from this opportunity. Do you want to start it?`,
                confirmButtonText: 'Yes',
                cancelButtonText: 'No, complete only'
              }
            });

            newPhaseDialog.afterClosed()
              .pipe(take(1))
              .subscribe(async response => {
                if (!this.opportunity) return;

                if (response) {

                  // create new
                  this.dialogInstance = this.matDialog.open(CreateOpportunityComponent, {
                    height: '700px',
                    width: '405px',
                    data: {
                      opportunity: { ...this.opportunity },
                      createFrom: true,
                      forceType: true
                    }
                  });

                  this.dialogInstance.afterClosed()
                    .pipe(take(1))
                    .subscribe(async (result: any) => {
                      if (!result || !this.opportunity) return;

                      if (result.success) {
                        // complete current opp
                        await this.sharepoint.setOpportunityStatus(this.opportunity.ID, "Approved");

                        this.toastr.success("A opportunity of type 'phase' was created successfully", result.data.opportunity.Title);
                        let opp = await this.sharepoint.getOpportunity(result.data.opportunity.ID);
                        opp.progress = 0;
                        let job = this.jobs.startJob(
                          "initialize opportunity " + result.data.opportunity.id,
                          'The new opportunity is being initialized. First stage and permissions are being created.'
                        );
                        this.sharepoint.initializeOpportunity(result.data.opportunity, result.data.stage).then(async r => {
                          // set active
                          await this.sharepoint.setOpportunityStatus(opp.ID, 'Active');
                          this.jobs.finishJob(job.id);
                          this.toastr.success("The opportunity is now active", opp.Title);
                          this.router.navigate(['opportunities', opp.ID, 'actions']);
                        }).catch(e => {
                          this.jobs.finishJob(job.id);
                        });
                      } else if (result.success === false) {
                        this.toastr.error("The opportunity couldn't be created", "Try again");
                      }
                    });
                } else if (response === false) {
                  // complete
                  await this.sharepoint.setOpportunityStatus(this.opportunity.ID, "Approved");
                  this.opportunity.OpportunityStatus = 'Approved';
                  this.toastr.success("The opportunity has been completed", this.opportunity.Title);
                }
              });
          } else {
            // complete
            await this.sharepoint.setOpportunityStatus(this.opportunity.ID, "Approved");
            this.opportunity.OpportunityStatus = 'Approved';
            this.toastr.success("The opportunity has been completed", this.opportunity.Title);
          }
        }
      });
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
      this.isStageUser = gate.StageUsersId.some(userId => userId === this.currentUser?.Id);
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
      this.currentFolderUri = `${this.opportunity?.BusinessUnitId}/${this.opportunityId}/${this.currentGate?.StageNameId}/`+folderId;
      if (this.currentFolder?.containsModels) {
        const geoFolders = await this.sharepoint.getSubfolders(this.currentFolderUri);
        this.currentFiles = [];
        for (const geofolder of geoFolders) {
          this.currentFiles.push(...await this.sharepoint.readFolderFiles(this.currentFolderUri + '/' + geofolder.Name + '/0', true));
        }
      } else {
        this.currentFiles = await this.sharepoint.readFolderFiles(this.currentFolderUri+'/0/0', true);
      }
  
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
      this.toastr.success("File downloaded to your Downloads folder.");
    } else {
      // get sharepoint base url TODO

      console.log('file info', fileInfo);
      
      /*
      ms-word:
      ms-powerpoint:
      ms-excel:
      ms-visio:
      ms-access:
      ms-project:
      ms-publisher:
      ms-spd:
      ms-infopath:
      */

      if(!fileInfo.LinkingUri) {
        this.toastr.error("This file type can't be openned online. Try downloading it instead.");
        return;
      }
      
      let arrUrl = fileInfo.LinkingUri.split("?");
      let url = arrUrl[0];
      const arrFile = url.split(".");
      const extension = arrFile[arrFile.length - 1];

      switch(extension) {
        case "xlsx":
        case "xls":
        case "csv":
          url = "ms-excel:"+url;
          break;
        case "docx":
        case "doc":
          url = "ms-word:"+url;
          break;
        case "pptx":
        case "ppt":
          url = "ms-powerpoint"+url;
          break;
        default:
          url = fileInfo.LinkingUri;
      }
      
      const data = window.open(url, '_blank');
      this.toastr.success("Trying to open file with your local Office installation.");
    }
  }

  async shareFile(fileId: number, departmentId: number) {
    const file = this.currentFiles.find(f => f.ListItemAllFields?.ID === fileId);
    if (!file) return;
    
    let folderGroup = `DU-${this.opportunityId}-${departmentId}`;

    // is it a model with geography assigned?
    if (file.ListItemAllFields?.EntityGeographyId) {
      folderGroup += '-' + file.ListItemAllFields?.EntityGeographyId;
    }
    
    // users with access
    let folderUsersList = await this.sharepoint.getGroupMembers(folderGroup);
    folderUsersList = folderUsersList.concat(
      await this.sharepoint.getGroupMembers('OO-' + this.opportunityId),
      await this.sharepoint.getGroupMembers('SU-' + this.opportunityId + '-' + this.currentGate?.StageNameId)
    );

    // clean users list
    let uniqueFolderUsersList = [...new Map(folderUsersList.map(u => [u.Id, u])).values()];
    // remove own user
    const currentUser = await this.sharepoint.getCurrentUserInfo();
    uniqueFolderUsersList = uniqueFolderUsersList.filter(el => el.Id !== currentUser.Id);

    this.matDialog.open(ShareDocumentComponent, {
      height: '300px',
      width: '405px',
      data: {
        file,
        folderUsersList: uniqueFolderUsersList
      }
    });
  }

  async editFile(fileId: number) {
    const fileInfo = this.currentFiles.find(f => f.ListItemAllFields?.ID === fileId);
    if (!fileInfo) return;

    const dialogRef = this.matDialog.open(EditFileComponent, {
      width: "300px",
      height: "225px",
      data: {
        fileInfo,
      }
    });

    dialogRef.afterClosed()
      .pipe(take(1))
      .subscribe(async result => {
        if (result.success) {
          fileInfo.Name = result.filename;
          this.toastr.success(`The file has been renamed`, "File Renamed");
        } else {
          this.toastr.error("Sorry, there was a problem renaming the file");
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
            this.toastr.success(`The file ${fileInfo.Name} has been deleted`, "File Removed");
          } else {
            this.toastr.error("Sorry, there was a problem deleting the file");
          }
        }
      });
  }

  ngOnDestroy() {
    clearTimeout(this.dateListener);
  }

  async refreshPowerBi() {
    try {
      if(!this.refreshingPowerBi) {
        this.refreshingPowerBi = true;
        //const at the moment needs to be dynamic
        const reportName: string = "Epi+"

        let response = await this.powerBi.refreshReport(reportName);
        console.log("status is: "+response);
        switch (response){
          case 202:{
            this.toastr.success("Analytics report succesfully refreshed.");
            break;
          }
          case 409:{
            this.toastr.error("Report currently refreshing. Please try again later");
            break;
          }
          default:{
            this.toastr.error(`Unknown error, ${response}`);
            break;
          }
        }

        this.refreshingPowerBi = false;   
      }  
    } catch(e: any) {
      this.refreshingPowerBi = false;
      
      this.toastr.error(e.message);
    }
    
  }

  navigateTo(item: Opportunity) {
   
    this.router.navigate(['/power-bi',
      {ID:item.ID}]);
    
  }

  private isActiveStage(stageId: number): boolean {
    const position = this.gates.map(el => el.StageNameId).indexOf(stageId);
    return position === this.gates.length - 1;
  }
}

