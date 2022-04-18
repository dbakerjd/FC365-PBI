import { Component, OnInit } from '@angular/core';
import { MatDialog } from '@angular/material/dialog';
import { DomSanitizer, SafeUrl } from '@angular/platform-browser';
import { ActivatedRoute, Router } from '@angular/router';
import { DatepickerOptions } from 'ng2-datepicker';
import { ToastrService } from 'ngx-toastr';
import { take } from 'rxjs/operators';
import { CommentsListComponent } from 'src/app/modals/comments-list/comments-list.component';
import { ConfirmDialogComponent } from 'src/app/modals/confirm-dialog/confirm-dialog.component';
import { CreateOpportunityComponent } from 'src/app/modals/create-opportunity/create-opportunity.component';
import { CreateScenarioComponent } from 'src/app/modals/create-scenario/create-scenario.component';
import { EditFileComponent } from 'src/app/modals/edit-file/edit-file.component';
import { ExternalApproveModelComponent } from 'src/app/modals/external-approve-model/external-approve-model.component';
import { FolderPermissionsComponent } from 'src/app/modals/folder-permissions/folder-permissions.component';
import { RejectModelComponent } from 'src/app/modals/reject-model/reject-model.component';
import { SendForApprovalComponent } from 'src/app/modals/send-for-approval/send-for-approval.component';
import { ShareDocumentComponent } from 'src/app/modals/share-document/share-document.component';
import { StageSettingsComponent } from 'src/app/modals/stage-settings/stage-settings.component';
import { UploadFileComponent } from 'src/app/modals/upload-file/upload-file.component';
import { BreadcrumbsService } from 'src/app/services/breadcrumbs.service';
import { LicensingService } from 'src/app/services/jd-data/licensing.service';
import { NotificationsService } from 'src/app/services/notifications.service';
import { PowerBiService } from 'src/app/services/power-bi.service';
import { WorkInProgressService } from '@services/app/work-in-progress.service';
import { Action, EntityGeography, Indication, Opportunity, Stage } from '@shared/models/entity';
import { FileComments, NPPFile, NPPFolder } from '@shared/models/file-system';
import { User } from '@shared/models/user';
import { FILES_FOLDER, FOLDER_DOCUMENTS } from '@shared/sharepoint/folders';
import { AppDataService } from '@services/app/app-data.service';
import { PermissionsService } from '@services/permissions.service';
import { FilesService } from '@services/files.service';
import { SelectInputList } from '@shared/models/app-config';
import { SelectListsService } from '@services/select-lists.service';

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
  opportunityGeographies: EntityGeography[] = []; // geographies (not soft removed)
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
  defaultProfilePic = '/assets/user.svg';
  ownerProfilePic: SafeUrl | null = null;
  hasAccessToModels = false;

  constructor(
    private route: ActivatedRoute, 
    private router: Router,
    public matDialog: MatDialog,
    private toastr: ToastrService,
    public licensing: LicensingService,
    public jobs: WorkInProgressService,
    public powerBi: PowerBiService,
    private breadcrumbService: BreadcrumbsService,
    public sanitize: DomSanitizer,
    private readonly appData: AppDataService,
    private readonly files: FilesService,
    private readonly permissions: PermissionsService,
    private readonly notifications: NotificationsService,
    private readonly selectLists: SelectListsService
    ) { }

  ngOnInit(): void {
    this.route.params.subscribe(async (params) => {
      if(params.id && params.id != this.opportunityId) {
        this.opportunityId = params.id;
        this.opportunity = await this.appData.getEntity(params.id);
        if (!this.opportunity) {
          this.router.navigate(['notfound']);
        }
        this.currentUser = await this.appData.getCurrentUserInfo();
        this.isOwner = this.currentUser.Id === this.opportunity.EntityOwnerId;
        this.breadcrumbService.addBreadcrumbLevel(this.opportunity.Title);
        this.opportunityGeographies = await this.appData.getEntityGeographies(this.opportunity.ID, false);

        if (this.opportunity.EntityOwner) {
          const profileImgBlob = await this.appData.getUserProfilePic(this.opportunity.EntityOwnerId);
          this.ownerProfilePic = profileImgBlob ? this.sanitize.bypassSecurityTrustUrl(window.URL.createObjectURL(profileImgBlob)) : null;
        }
        this.gates = await this.appData.getEntityStages(params.id);
        this.gates.forEach(async (el, index) => {
          el.actions = await this.appData.getActions(params.id, el.StageNameId);
          el.folders = await this.appData.getStageFolders(el.StageNameId, this.opportunityId, this.opportunity?.BusinessUnitId);
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
            this.nextStage = await this.appData.getNextStage(el.StageNameId);
          }

        });

        this.setUpDateListener();
      }
    });
  }

  getIndications(indications: Indication[]) {
    if(indications) {
      return indications.map(el => el.Title).join(", ");
    }
    return '';
  }

  async openUploadDialog() {
    if (!this.currentGate) return;

    let geographiesList: SelectInputList[] = [];
    const modelFolder = this.currentFolders.find(f => f.containsModels);
    if (this.opportunity) {
      geographiesList = await this.selectLists.getEntityAccessibleGeographiesList(
        this.opportunity,
        this.currentGate.StageNameId
      );
    }
    this.dialogInstance = this.matDialog.open(UploadFileComponent, {
      height: '600px',
      width: '405px',
      data: {
        folderList: this.currentFolders,
        selectedFolder: this.currentSection === 'documents' && this.currentFolder ? this.currentFolder.DepartmentID : null,
        geographies: geographiesList,
        scenarios: await this.selectLists.getScenariosList(),
        masterStageId: this.currentGate?.StageNameId,
        entity: this.opportunity
      }
    });

    this.dialogInstance.afterClosed()
    .pipe(take(1))
    .subscribe(async (result: any) => {
      if (result.success && result.uploaded) {
        this.toastr.success(`The file ${result.name} was uploaded successfully`, "File Uploaded");
        await this.updateCurrentFiles();
      } else if (result.success === false) {
        this.toastr.error("Sorry, there was a problem uploading your file");
      }
    });

    
  }

  openCommentsDetail(file: NPPFile) {
    let comments: FileComments[] = [];
    if (file.ListItemAllFields && file.ListItemAllFields.Comments) {
      try {
        comments = JSON.parse(file.ListItemAllFields.Comments);
      } catch(e) {
        console.log("Error parsing comments for file "+file.ListItemAllFields.ID);
      }

      this.dialogInstance = this.matDialog.open(CommentsListComponent, {
        height: '75vh',
        width: '500px',
        data: {
          comments
        }
      });
    }
  }

  async updateCurrentFiles() {
    if (this.currentFolder?.containsModels) {
      let geoFolders = await this.appData.getSubfolders('/'+this.currentFolderUri);
      geoFolders = geoFolders.filter((gf: any) => this.opportunityGeographies.some((og: EntityGeography) => +gf.Name === og.ID));
      this.currentFiles = [];
      for (const geofolder of geoFolders) {
        this.currentFiles.push(...await this.appData.getFolderFiles(FILES_FOLDER + '/' + this.currentFolderUri + '/' + geofolder.Name+'/0', true));
      }
    } else {
      this.currentFiles = await this.appData.getFolderFiles(FILES_FOLDER + '/' + this.currentFolderUri+'/0/0', true);
    }
    this.initLastComments();
  }

  initLastComments() {
    if (this.currentSection === 'documents' && this.currentFolder?.containsModels) {
      this.currentFiles.forEach(el => {
        el.lastComments = this.getLatestComments(el);
      });
    }
  }

  getLatestComments(file: NPPFile): FileComments[] {
    let comments: FileComments[] = [];
    let numComments = 1;
    let lastComments = [];

    if (file.ListItemAllFields && file.ListItemAllFields.Comments) {
      try {
        comments = JSON.parse(file.ListItemAllFields.Comments);
      } catch(e) {
        console.log("Error parsing comments for file "+file.ListItemAllFields.ID);
      }

      for(let i = (comments.length - 1); i >= 0 && (numComments - ((comments.length - 1 ) - i) > 0); i--) {
        lastComments.push(comments[i]);
      }
    }

    return lastComments;
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
          this.updateCurrentFiles();
          //generate notifications
          this.toastr.success("The model has been sent for approval", "Forecast Model");
          await this.notifications.modelSubmittedNotification(file.Name, this.opportunityId, [
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
    if (!this.opportunity) return;

    this.dialogInstance = this.matDialog.open(ExternalApproveModelComponent, {
      height: '300px',
      width: '405px',
      data: {
        file: file,
        entity: this.opportunity,
        rootFolder: FOLDER_DOCUMENTS,
        departmentID: this.currentFolder?.DepartmentID
      }
    });

    this.dialogInstance.afterClosed()
      .pipe(take(1))
      .subscribe(async (result: any) => {
        if (result.success) {
          // update view
          await this.updateCurrentFiles();
          this.toastr.success("The model has been approved!", "Forecast Model");
          await this.notifications.modelApprovedNotification(file.Name, this.opportunityId, [
            `DU-${this.opportunityId}-${departmentId}-${file.ListItemAllFields?.EntityGeographyId}`,
            `OO-${this.opportunityId}`,
            `SU-${this.opportunityId}-${this.currentGate?.StageNameId}`,
          ]);
        } else if (result.success === false) {
          this.toastr.error("There was a problem approving the forecast model", 'Try again');
        }
      });
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
          await this.updateCurrentFiles();
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

  createScenario(file: NPPFile, departmentId: number) {
    this.dialogInstance = this.matDialog.open(CreateScenarioComponent, {
      height: '450px',
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
          await this.updateCurrentFiles();
          await this.notifications.modelNewScenarioNotification(file.Name, this.opportunityId, [
            `DU-${this.opportunityId}-${departmentId}-${file.ListItemAllFields?.EntityGeographyId}`,
            `OO-${this.opportunityId}`,
            `SU-${this.opportunityId}-${this.currentGate?.StageNameId}`,
          ]);
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
        height: '500px',
        width: '405px',
        data: {
          folderList: this.currentFolders,
          entity: this.opportunity,
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
    if (this.currentFolders.length > 0) this.setFolder(this.currentFolders[0].DepartmentID);
    else this.setFolder(undefined);
  }

  showModels() {
    this.setSection('documents');
    let modelsFolder = this.currentFolders.find(el => el.containsModels === true);
    if (modelsFolder) this.setFolder(modelsFolder.DepartmentID);
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
    if (!this.currentUser) this.currentUser = await this.appData.getCurrentUserInfo(); // no tenim ID user al sharepoint

    if (action.Complete) done = await this.appData.uncompleteAction(action.Id);
    else {
      done = await this.appData.completeAction(action.Id, this.currentUser.Id);
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
    this.appData.setActionDueDate(actionId, value);
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
              'initialize stage ' + result.data.ID
            );
            let opp = await this.appData.getEntity(result.data.EntityNameId);
            const oppGeographies = await this.appData.getEntityGeographies(opp.ID);
            this.alreadyGoingNextStage = true;
            this.permissions.initializeStage(opp, result.data,oppGeographies).then(async r => {
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
      minWidth: "350px",
      height: "200px",
      data: {
        message: `Do you want to complete the opportunity <em>${this.opportunity.Title}</em>?`,
        confirmButtonText: 'Yes, complete',
      }
    });

    dialogRef.afterClosed()
      .pipe(take(1))
      .subscribe(async completeResponse => {
        if (completeResponse) {
          // complete opportunity
          if (!this.opportunity) return;

          if (!await this.appData.isInternalOpportunity(this.opportunity.OpportunityTypeId)) {
            const newPhaseDialog = this.matDialog.open(ConfirmDialogComponent, {
              maxWidth: "400px",
              height: "200px",
              data: {
                message: `You can move to a Product Development Opportunity. Do you want to start it?`,
                confirmButtonText: 'Yes',
                cancelButtonText: 'No, complete only'
              }
            });

            newPhaseDialog.afterClosed()
              .pipe(take(1))
              .subscribe(async newInternalResponse => {
                if (!this.opportunity) return;
                if (newInternalResponse) {

                  // create new
                  this.dialogInstance = this.matDialog.open(CreateOpportunityComponent, {
                    height: '75vh',
                    width: '500px',
                    data: {
                      opportunity: { ...this.opportunity },
                      createFrom: true,
                      forceInternal: true
                    }
                  });

                  this.dialogInstance.afterClosed()
                    .pipe(take(1))
                    .subscribe(async (result: any) => {
                      if (!result || !this.opportunity) return;

                      if (result.success) {
                        // complete current opp
                        await this.appData.setOpportunityStatus(this.opportunity.ID, "Approved");

                        this.toastr.success("A new opportunity was created successfully", result.data.opportunity.Title);
                        let opp = await this.appData.getEntity(result.data.opportunity.ID);
                        opp.progress = 0;
                        let job = this.jobs.startJob(
                          "initialize opportunity " + result.data.opportunity.id
                        );
                        this.permissions.initializeOpportunity(result.data.opportunity, result.data.stage).then(async r => {
                          if (!this.opportunity) return;
                          this.files.copyFilesExternalToInternal(this.opportunity?.ID, opp.ID);
                          // set active
                          await this.appData.setOpportunityStatus(opp.ID, 'Active');
                          this.jobs.finishJob(job.id);
                          this.toastr.success("The opportunity is now active", opp.Title);
                          this.router.navigate(['opportunities', opp.ID, 'files']);
                        }).catch(e => {
                          this.jobs.finishJob(job.id);
                        });
                      } else if (result.success === false) {
                        this.toastr.error("The opportunity couldn't be created", "Try again");
                      }
                    });
                } else if (newInternalResponse === false) {
                  // only complete
                  await this.appData.setOpportunityStatus(this.opportunity.ID, "Approved");
                  this.opportunity.OpportunityStatus = 'Approved';
                  this.toastr.success("The opportunity has been completed", this.opportunity.Title);
                }
              });
          } else { // without possibility of pass to internal => complete
            await this.appData.setOpportunityStatus(this.opportunity.ID, "Approved");
            this.opportunity.OpportunityStatus = 'Approved';
            this.toastr.success("The opportunity has been completed", this.opportunity.Title);
          }
         // complete
        //  await this.appData.setOpportunityStatus(this.opportunity.ID, "Approved");
        //  this.opportunity.OpportunityStatus = 'Approved';
        //  this.toastr.success("The opportunity has been completed", this.opportunity.Title);
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
      this.hasAccessToModels = this.currentFolders.some((f: NPPFolder) => f.containsModels);
      if (this.currentFolders.length) this.setFolder(this.currentFolders[0].DepartmentID);
    }
  }

  async setFolder(folderId: number | undefined) {
    this.currentFiles = [];
    if (folderId || folderId === 0) {
      this.loading = true;
      this.currentFolder = this.currentFolders.find(el => el.DepartmentID === folderId);
      this.currentFolderUri = `${this.opportunity?.BusinessUnitId}/${this.opportunityId}/${this.currentGate?.StageNameId}/`+folderId;
      
      await this.updateCurrentFiles();
  
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

    const response = await this.appData.readFile(fileInfo.ServerRelativeUrl);
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
        this.toastr.error("This file type can't be opened online. Try downloading it instead.");
        return;
      }
      
      let arrUrl = fileInfo.LinkingUri.split("?");
      let url = arrUrl[0];
      const arrFile = url.split(".");
      const extension = arrFile[arrFile.length - 1];

      switch(extension) {
        case "xlsx":
        case "xlsm":
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
    let folderUsersList = await this.appData.getGroupMembers(folderGroup);
    folderUsersList = folderUsersList.concat(
      await this.appData.getGroupMembers('OO-' + this.opportunityId),
      await this.appData.getGroupMembers('SU-' + this.opportunityId + '-' + this.currentGate?.StageNameId)
    );

    // clean users list
    let uniqueFolderUsersList = [...new Map(folderUsersList.map(u => [u.Id, u])).values()];
    // remove own user
    const currentUser = await this.appData.getCurrentUserInfo();
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
      width: "400px",
      height: "300px",
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
          await this.updateCurrentFiles();
        } else if (result.success === false) {
          this.toastr.error("Sorry, there was a problem renaming the file");
        }
      });
  }

  async deleteFile(fileId: number) {
    const fileInfo = this.currentFiles.find(f => f.ListItemAllFields?.ID === fileId);
    if (!fileInfo) return;

    const dialogRef = this.matDialog.open(ConfirmDialogComponent, {
      width: "370px",
      maxWidth: "400px",
      height: "250px",
      maxHeight: "300px",
      data: {
        message: 'Are you sure you want to delete the file <em>' + fileInfo.Name + '</em> ?',
        confirmButtonText: 'Yes, delete'
      }
    });

    dialogRef.afterClosed()
      .pipe(take(1))
      .subscribe(async deleteConfirmed => {
        if (deleteConfirmed) {
          if (await this.files.deleteFile(fileInfo.ServerRelativeUrl, this.currentFolder?.containsModels)) {
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
        const reportName: string = "Epi Report"

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

