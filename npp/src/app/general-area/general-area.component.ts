import { Component, OnInit } from '@angular/core';
import { AppDataService } from '@services/app/app-data.service';
import { NPPFile, NPPFolder } from '@shared/models/file-system';
import { User } from '@shared/models/user';

@Component({
  selector: 'app-general-area',
  templateUrl: './general-area.component.html',
  styleUrls: ['./general-area.component.scss']
})
export class GeneralAreaComponent implements OnInit {

  currentUser: User | undefined = undefined;
  currentFolder: NPPFolder | undefined = undefined;
  currentFiles: NPPFile[] = [];
  documentFolders: NPPFolder[] = [];
  loading = false;

  constructor(
    private readonly appData: AppDataService
  ) { }

  async ngOnInit(): Promise<void> {
    this.currentUser = await this.appData.getCurrentUserInfo();
  }

  async openUploadDialog() {
    
  }
}
