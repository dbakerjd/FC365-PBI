import { Component, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MatDialog } from '@angular/material/dialog';
import { ActivatedRoute } from '@angular/router';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { DatepickerOptions } from 'ng2-datepicker';
import { ToastrService } from 'ngx-toastr';
import { take } from 'rxjs/operators';
import { CommentsListComponent } from 'src/app/modals/comments-list/comments-list.component';
import { ConfirmDialogComponent } from 'src/app/modals/confirm-dialog/confirm-dialog.component';
import { CreateForecastCycleComponent } from 'src/app/modals/create-forecast-cycle/create-forecast-cycle.component';
import { CreateScenarioComponent } from 'src/app/modals/create-scenario/create-scenario.component';
import { EntityEditFileComponent } from 'src/app/modals/entity-edit-file/entity-edit-file.component';
import { ExternalApproveModelComponent } from 'src/app/modals/external-approve-model/external-approve-model.component';
import { ExternalUploadFileComponent } from 'src/app/modals/external-upload-file/external-upload-file.component';
import { FolderPermissionsComponent } from 'src/app/modals/folder-permissions/folder-permissions.component';
import { RejectModelComponent } from 'src/app/modals/reject-model/reject-model.component';
import { SendForApprovalComponent } from 'src/app/modals/send-for-approval/send-for-approval.component';
import { ShareDocumentComponent } from 'src/app/modals/share-document/share-document.component';
import { InlineNppDisambiguationService } from 'src/app/services/inline-npp-disambiguation.service';
import { LicensingService } from 'src/app/services/licensing.service';
import { NotificationsService } from 'src/app/services/notifications.service';
import { PowerBiService } from 'src/app/services/power-bi.service';
import { SharepointService, FileComments, Brand, NPPFile, SelectInputList, User, FORECAST_MODELS_FOLDER_NAME, NPPFolder, NPPFileMetadata, ForecastCycle, BrandForecastCycle, Indication, Opportunity, FOLDER_ARCHIVED, FOLDER_APPROVED, FOLDER_WIP, FOLDER_DOCUMENTS, FILES_FOLDER } from 'src/app/services/sharepoint.service';
import { TeamsService } from 'src/app/services/teams.service';

@Component({
  selector: 'app-files-list',
  templateUrl: './files-list.component.html',
  styleUrls: ['./files-list.component.scss']
})
export class FilesListComponent implements OnInit {
  isOwner = false;
  currentDepartmentId: number = 0;
  currentUser: User | undefined = undefined;
  currentFolder: NPPFolder | undefined = undefined;
  selectedFolder: NPPFolder | undefined = undefined;
  selectedDepartmentId: number = 0;
  documentFolders: NPPFolder[] = [];
  cycles: BrandForecastCycle[] = [];
  refreshingPowerBi = false;
  entityId = 0;
  entity: Brand | Opportunity | undefined = undefined;
  dateOptions: DatepickerOptions = {
    format: 'Y-M-d'
  };
  profilePic: string | boolean = '';
  currentSection = 'models';
  currentFiles: NPPFile[] = [];
  uploadDialogInstance: any; 
  modelStatus = ['Work in Progress', 'Approved', 'Archived'];
  currentStatus = this.modelStatus[0];
  dialogInstance: any; 
  formCycleSelect = new FormGroup({});
  formCycleSelectFields: FormlyFieldConfig[] = [];
  currentCycle: number | undefined;
  masterCycles: SelectInputList[] = [];
  updatingFiles = false;
  updateFilesTimeout: any;
  selectedCycle: any = false;

  constructor(
    private sharepoint: SharepointService, 
    private powerBi: PowerBiService, 
    private route: ActivatedRoute, 
    public matDialog: MatDialog,
    private toastr: ToastrService, 
    private teams: TeamsService,
    public licensing: LicensingService,
    public disambiguator: InlineNppDisambiguationService,
    public notifications: NotificationsService) { }

  ngOnInit(): void {
    if(this.teams.initialized) this.init();
    else {
      this.teams.statusSubject.subscribe(async (msg) => {
        setTimeout(async () => {
          this.init();
        }, 500);
      });
    }
  }

  init() {
    this.route.params.subscribe(async (params) => {
      this.currentUser = await this.sharepoint.getCurrentUserInfo();
      this.masterCycles = await this.sharepoint.getForecastCycles();

      if(params.id && params.id != this.entityId) {
        this.entityId = params.id;
        this.entity = await this.disambiguator.getEntity(params.id);
        this.documentFolders = await this.sharepoint.getInternalDepartments(this.entityId, this.entity.BusinessUnitId);
        let owner = this.entity.EntityOwner;
        let ownerId = this.entity.EntityOwnerId;
        this.isOwner = this.currentUser.Id === ownerId;
        if (this.entity && owner) {
          
          this.cycles = await this.disambiguator.getForecastCycles(this.entity);

          let pic = await this.sharepoint.getUserProfilePic(ownerId);
          owner.profilePicUrl = pic ? pic : '/assets/user.svg';
          this.profilePic = owner.profilePicUrl;
        }
        this.setStatus(this.modelStatus[0]);
      }
    });
  }

  onCycleChange() {
    this.currentCycle = this.formCycleSelect.value?.cycle;
    this.updateCurrentFiles();
  }

  getSharepointFolderNameByModelStatus(status: string) {
    switch(status) {
      case 'Archived':
        return FOLDER_ARCHIVED+'/'+this.entity?.BusinessUnitId+'/'+this.entity?.ID+'/0/0';
      case 'Approved':
        return FOLDER_APPROVED+'/'+this.entity?.BusinessUnitId+'/'+this.entity?.ID+'/0/0';
      case 'Work in Progress':
        return FOLDER_WIP+'/'+this.entity?.BusinessUnitId+'/'+this.entity?.ID+'/0/0';
      default:
        return FOLDER_DOCUMENTS+'/'+this.entity?.BusinessUnitId+'/'+this.entity?.ID+'/0/'+this.selectedDepartmentId+'/0/0';
    }
  }

  getCurrentRootFolder() {
    return this.getRootFolder(this.currentStatus);
  }

  getRootFolder(status: string) {
    switch(status) {
      case 'Archived':
        return FOLDER_ARCHIVED;
      case 'Approved':
        return FOLDER_APPROVED;
      case 'Work in Progress':
        return FOLDER_WIP;
      default:
        return FOLDER_DOCUMENTS;
    }
  }

  async setFolder(folder: NPPFolder) {
    this.selectedCycle = false;
    this.currentCycle = undefined;
    this.currentStatus = 'none';
    this.selectedFolder = folder;
    this.selectedDepartmentId = folder.DepartmentID ? folder.DepartmentID : 0;
    this.updateCurrentFiles();
  }

  async updateCurrentFiles() {
    try {
      if(!this.updatingFiles) {
        this.updatingFiles = true;
        let currentFolder = this.getCurrentFolder();
        
        if (this.currentStatus != 'none') {
          const geoFolders = await this.sharepoint.getSubfolders(currentFolder, true);
          this.currentFiles = [];
          for (const geofolder of geoFolders) {
            let folder = currentFolder + '/' + geofolder.Name;
            if(this.currentStatus == 'Archived') {
              folder = folder + '/' + this.currentCycle;
            } else {
              folder = folder + '/0';
            }
            this.currentFiles.push(...await this.disambiguator.readFolderFiles(folder, true));
          }
        } else {
          this.currentFiles = await this.disambiguator.readFolderFiles(currentFolder, true);
        }

        this.initLastComments();

        this.updatingFiles = false;

      } else {
        
        if(this.updateFilesTimeout) {
          clearTimeout(this.updateFilesTimeout);
        }
        
        this.updateFilesTimeout = setTimeout(() => {
          this.updateCurrentFiles();
        }, 500);
        
      }
    } catch(e: any) {
      this.updatingFiles = false;
    }
    
    
  }

  getCurrentFolder() {
    return this.getSharepointFolderNameByModelStatus(this.currentStatus);
  }

  getIndications(indications: Indication[]) {
    if(indications) {
      return indications.map(el => el.Title).join(", ");
    }
    return '';
  }

  getTherapyArea(indications: Indication[]) {
    if(indications && indications.length) {
      return indications[0].TherapyArea;
    }

    return '';
  }

  async openUploadDialog() {
    if(this.entity) {
      let geographiesList = await this.disambiguator.getAccessibleGeographiesList(this.entity);
      let folders = [...this.documentFolders]
      this.dialogInstance = this.matDialog.open(ExternalUploadFileComponent, {
        height: '600px',
        width: '405px',
        data: {
          geographies: geographiesList,
          scenarios: await this.sharepoint.getScenariosList(),
          folderList: folders,
          selectedFolder: this.currentSection == 'none' && this.selectedFolder ?  this.selectedFolder.ID : 'Forecast Models',
          entity: this.entity
        }
      });
  
      this.dialogInstance.afterClosed()
      .pipe(take(1))
      .subscribe(async (result: any) => {
        if (result.success) {
          this.toastr.success(`The file ${result.name} was uploaded successfully`, "File Uploaded");
          this.updateCurrentFiles();
        } else if (result.success === false) {
          this.toastr.error("Sorry, there was a problem uploading your file");
        }
      });
    }
      
  }

  sendForApproval(file: NPPFile) {
    this.dialogInstance = this.matDialog.open(SendForApprovalComponent, {
      height: '300px',
      width: '405px',
      data: {
        file: file,
        rootFolder: this.getCurrentRootFolder(),
        entity: this.entity
      }
    });

    this.dialogInstance.afterClosed()
      .pipe(take(1))
      .subscribe(async (result: any) => {
        if (result.success) {
          // update view
          this.updateCurrentFiles();
          this.toastr.success("The model has been sent for approval", "Forecast Model");
          await this.notifications.modelSubmittedNotification(file.Name, this.entityId, [
            `DU-${this.entityId}-0-${file.ListItemAllFields?.EntityGeographyId}`,
            `OO-${this.entityId}`
          ]);
        } else if (result.success === false) {
          this.toastr.error("The model couldn't be sent for approval");
        }
      });
  }

  openFolderPermissions() {
    if (this.isOwner || this.currentUser?.IsSiteAdmin) { // TODO: open to all stage users when using API
      let folders = [...this.documentFolders]
      this.dialogInstance = this.matDialog.open(FolderPermissionsComponent, {
        height: '500px',
        width: '405px',
        data: {
          entity: this.entity,
          folderList: folders
        }
      });
    }
  }

  async approve(file: NPPFile) {
    
    if (!file.ListItemAllFields) return;
    if (!this.entity) return;

    this.dialogInstance = this.matDialog.open(ExternalApproveModelComponent, {
      height: '300px',
      width: '405px',
      data: {
        file: file,
        entity: this.entity,
        rootFolder: this.getCurrentRootFolder(),
        departmentID: this.currentDepartmentId
      }
    });

    this.dialogInstance.afterClosed()
      .pipe(take(1))
      .subscribe(async (result: any) => {
        if (result.success) {
          // update view
          this.updateCurrentFiles();
          this.toastr.success("The model has been approved!", "Forecast Model");
        } else if (result.success === false) {
          this.toastr.error("There was a problem approving the forecast model", 'Try again');
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
        if (success) {
          this.toastr.success(`The new model scenario has been created successfully`, "New Forecast Model");
          this.updateCurrentFiles();
        } else if (success === false) {
          this.toastr.error('The new model scenario could not be created', 'Try Again');
        }
      });
  }

  setSection(section: string) {
    this.currentSection = section;
    if(section == 'models') {
      this.setStatus(this.modelStatus[0]);
    } else {
      this.setStatus('none');
    }
  }

  async setStatus(status: string) {
    this.selectedCycle = false;
    this.currentCycle = undefined;
    this.currentStatus = status;
    this.selectedFolder = undefined;
    this.selectedDepartmentId = 0;
    this.updateCurrentFiles();
  }

  showFolders() {
    this.setSection('documents');
    this.setStatus('none');
  }

  showModels() {
    this.setSection('models');
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

  async shareFile(fileId: number, geoId: number | null = null, countryId: number | null = null) {
    const file = this.currentFiles.find(f => f.ListItemAllFields?.ID === fileId);
    if (!file) return;
    
    const oppGeo = await this.disambiguator.getEntityGeographies(this.entityId);

    let involvedGeo = null;
    if (geoId) {
      involvedGeo = oppGeo.find(el => el.Master_x0020_GeographyId == geoId);
    } else if (countryId) {
      involvedGeo = oppGeo.find(el => el.CountryId == countryId);
    }
    if (!involvedGeo && (geoId || countryId)) return;

    let folderUsersList: User[] = [];
    if (involvedGeo) {
      let folderGroup = this.disambiguator.getGroupName(`EU-${this.entityId}-${involvedGeo.Id}`);
      folderUsersList = await this.sharepoint.getGroupMembers(folderGroup);
    }
    
    // users with access
    folderUsersList = folderUsersList.concat(
      await this.sharepoint.getGroupMembers( this.disambiguator.getGroupName('EO-' + this.entityId))
    );

    // remove own user
    const currentUser = await this.sharepoint.getCurrentUserInfo();
    folderUsersList = folderUsersList.filter(el => el.Id !== currentUser.Id);

    this.matDialog.open(ShareDocumentComponent, {
      height: '250px',
      width: '405px',
      data: {
        file,
        folderUsersList
      }
    });
  }


  async editFile(fileId: number) {
    const fileInfo = this.currentFiles.find(f => f.ListItemAllFields?.ID === fileId);
    if (!fileInfo) return;

    const dialogRef = this.matDialog.open(EntityEditFileComponent, {
      width: "400px",
      height: "325px",
      data: {
        fileInfo,
        entity: this.entity
      }
    });

    dialogRef.afterClosed()
      .pipe(take(1))
      .subscribe(async result => {
        let res = result.success;
        let str = '';
        let error = false;
        if(res.needsRename && res.renameWorked) {
          str = `The file has been renamed.`;
        }
        if(res.needsRename && !res.renameWorked) {
          error = true;
          str = `Sorry there was a problem renaming the file.`
        }

        if(res.needsIndicationsUpdate && res.indicationsUpdateWorked) {
          str += ` Indications have been updated.`
        }

        if(res.needsIndicationsUpdate && !res.indicationsUpdateWorked) {
          error = true;
          str += ` There was an error updating model indications.`
        }

        if (!error) {
          fileInfo.Name = result.filename;
          this.toastr.success(str, "File Update");
          this.updateCurrentFiles();
        } else {
          this.toastr.error(str);
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
  
  async refreshPowerBi() {
    try {
      if(!this.refreshingPowerBi) {
        this.refreshingPowerBi = true;
        const reportName: string = "Epi+";

        let response = await this.powerBi.refreshReport(reportName);
        this.refreshingPowerBi = false;   
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
      }  
    } catch(e: any) {
      this.refreshingPowerBi = false;
      this.toastr.error(e.message);
    }
    
  }

  createForecast() {
    this.dialogInstance = this.matDialog.open(CreateForecastCycleComponent, {
      height: '400px',
      width: '405px',
      data: {
        entity: this.entity
      }
    });

    this.dialogInstance.afterClosed()
      .pipe(take(1))
      .subscribe(async (success: any) => {
        if (success) {
          this.toastr.success(`The new forecast cycle has been created successfully`, "New Forecast Cycle");
          if(this.entity) this.cycles = await this.disambiguator.getForecastCycles(this.entity);
          this.entity = Object.assign(this.entity, {
            ForecastCycleId: success.ForecastCycleId,
            ForecastCycle: { 
              Title: this.masterCycles.find(el => el.value == success.ForecastCycleId)?.label,
              ID: success.ForecastCycleId
            },
            Year: success.Year
          });
          this.updateCurrentFiles();
        } else if (success === false) {
          this.toastr.error('The new forecast cycle could not be created', 'Try Again');
        }
      });
  }

  selectCycle(cycle: ForecastCycle) {
    if(!cycle) {
      this.selectedCycle = false;
      this.currentCycle = undefined;  
    } else {
      this.selectedCycle = cycle;
      this.currentCycle = cycle.ID;
      this.updateCurrentFiles();
    }
    
  }

  async rejectModel(file: NPPFile) {
    if (!file.ListItemAllFields) return;
    this.dialogInstance = this.matDialog.open(RejectModelComponent, {
      height: '300px',
      width: '405px',
      data: {
        file: file,
        rootFolder: this.getCurrentRootFolder(),
        entity: this.entity
      }
    });

    this.dialogInstance.afterClosed()
      .pipe(take(1))
      .subscribe(async (result: any) => {
        if (result.success) {
          // update view
          await this.updateCurrentFiles();
          this.toastr.warning("The model " + file.Name + " has been rejected", "Forecast Model");
          await this.notifications.modelRejectedNotification(file.Name, this.entityId, [
            `DU-${this.entityId}-0-${file.ListItemAllFields?.EntityGeographyId}`,
            `OO-${this.entityId}`
          ]);
        } else if (result.success === false) {
          this.toastr.error("There were a problem rejecting the forecast model", 'Try again');
        }
      });
  }

  initLastComments() {
    this.currentFiles.forEach(el => {
      el.lastComments = this.getLatestComments(el);
    });
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

}
