import { Component, OnInit } from '@angular/core';
import { MatDialog } from '@angular/material/dialog';
import { AppDataService } from '@services/app/app-data.service';
import { FilesService } from '@services/files.service';
import { NPPFile, NPPFolder, SystemFolder } from '@shared/models/file-system';
import { User } from '@shared/models/user';
import { GLOBAL_DOCUMENTS_FOLDER } from '@shared/sharepoint/folders';
import { ToastrService } from 'ngx-toastr';
import { take } from 'rxjs/operators';
import { ConfirmDialogComponent } from '../modals/confirm-dialog/confirm-dialog.component';
import { SimpleUploadComponent } from '../modals/simple-upload/simple-upload.component';

@Component({
  selector: 'app-general-area',
  templateUrl: './general-area.component.html',
  styleUrls: ['./general-area.component.scss']
})
export class GeneralAreaComponent implements OnInit {

  currentUser: User | undefined = undefined;
  currentFolder: NPPFolder | undefined = undefined;
  currentFiles: NPPFile[] = [];
  documentFolders: SystemFolder[] = [];
  selectedFolder: SystemFolder | undefined = undefined;
  selectedFolderWritable = false;
  loading = true;
  updatingFiles = false;
  updateFilesTimeout: any;
  dialogInstance: any; 

  constructor(
    private readonly appData: AppDataService,
    private readonly files: FilesService,
    public matDialog: MatDialog,
    private readonly toastr: ToastrService
    ) { }

  async ngOnInit(): Promise<void> {
    this.currentUser = await this.appData.getCurrentUserInfo();
    this.documentFolders = await this.appData.getSubfolders(GLOBAL_DOCUMENTS_FOLDER, true);
    this.documentFolders = this.documentFolders.filter(f => f.Name !== 'Forms'); // remove hidden system sharepoint folder
    if (this.documentFolders.length > 0) this.setFolder(this.documentFolders[0]);
  }

  async openUploadDialog() {
    this.dialogInstance = this.matDialog.open(SimpleUploadComponent, {
      height: '600px',
      width: '405px',
      data: {
        folder: this.selectedFolder,
      }
    });

    this.dialogInstance.afterClosed()
    .pipe(take(1))
    .subscribe(async (result: any) => {
      if (result.success && result.uploaded) {
        this.toastr.success(`The file ${result.name} was uploaded successfully`, "File Uploaded");
        this.updateCurrentFiles();
      } else if (result.success === false) {
        this.toastr.error("Sorry, there was a problem uploading your file");
      }
    });
  }

  async setFolder(folder: SystemFolder) {
    this.selectedFolderWritable = folder.Name === 'General' ? true : false;
    this.selectedFolder = folder;
    await this.updateCurrentFiles();
  }

  private async updateCurrentFiles() {
    this.loading = true;
    try {
      if(!this.updatingFiles) {
        this.updatingFiles = true;
        let currentFolder = GLOBAL_DOCUMENTS_FOLDER + '/' + this.selectedFolder?.Name;
        this.currentFiles = await this.appData.getFolderFiles(currentFolder, true);
        this.updatingFiles = false;
      } else {
        if(this.updateFilesTimeout) {
          clearTimeout(this.updateFilesTimeout);
        }
        
        this.updateFilesTimeout = setTimeout(() => {
          this.updateCurrentFiles();
        }, 500);
      }
      this.loading = false;
    } catch(e: any) {
      this.updatingFiles = false;
      this.loading = false;
    }
  }

  canUpload() {
    return this.currentUser?.IsSiteAdmin && this.selectedFolderWritable;
  }

  async downloadFile(fileId: number) {
    const fileInfo = this.currentFiles.find(f => f.ListItemAllFields?.ID === fileId);
    if (!fileInfo) return;

    const response = await this.appData.readFile(fileInfo.ServerRelativeUrl);
    var newBlob = new Blob([response]);

    var link = document.createElement('a');
    link.href = window.URL.createObjectURL(newBlob);
    link.download = fileInfo.Name;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    this.toastr.success("File downloaded to your Downloads folder.");
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
          if (await this.files.deleteFile(fileInfo.ServerRelativeUrl)) {
            // remove file for the current files list
            this.currentFiles = this.currentFiles.filter(f => f.ListItemAllFields?.ID !== fileId);
            this.toastr.success(`The file ${fileInfo.Name} has been deleted`, "File Removed");
          } else {
            this.toastr.error("Sorry, there was a problem deleting the file");
          }
        }
      });
  }
}
