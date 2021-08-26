import { Inject, Component, OnInit } from '@angular/core';
import { UploadFileConfig } from 'src/app/shared/forms/upload-file.config';
import { MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { FormGroup } from '@angular/forms';
import { NPPFolder, SharepointService } from 'src/app/services/sharepoint.service';
import { Observable } from 'rxjs';

@Component({
  selector: 'app-upload-file',
  templateUrl: './upload-file.component.html',
  styleUrls: ['./upload-file.component.scss']
})
export class UploadFileComponent implements OnInit {
  formConfig: UploadFileConfig = new UploadFileConfig();
  fields: FormlyFieldConfig[] = [];
  form: FormGroup = new FormGroup({});
  model: any = { };
  folders: NPPFolder[] = [];

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    private readonly sharepoint: SharepointService
  ) { 
    
  }

  ngOnInit(): void {
    this.formConfig = new UploadFileConfig();
    this.fields = this.formConfig.fields(
      this.data.opportunityId, 
      this.data.masterStageId, 
      this.data.folderList,
      this.data.countries,
      this.data.scenarios);
    this.form = new FormGroup({});
  }

  onSubmit() {
    let fileData = {
      StageNameId: this.model.StageNameId,
      OpportunityNameId: this.model.OpportunityNameId,
    };

    if (this.data.folderList.find((f: NPPFolder) => f.ID == this.model.category).containsModels) {
      // forecast model file
      Object.assign(fileData, {
        CountryId: this.model.country,
        ModelScenarioId: this.model.scenario,
        ModelApprovalComments: this.model.description
      });
    } else {
      // regular file
      Object.assign(fileData, {
        ModelApprovalComments: this.model.description
      });
    }

    // upload file to correspondent folder
    let folder = this.sharepoint.getBaseFilesFolder() + '/' + this.model.OpportunityNameId + '/' + this.model.StageNameId + '/' + this.model.category
    this.readFileDataAsText(this.model.file[0]).subscribe(
      data => {
        this.sharepoint.uploadFile(data, folder, this.model.file[0].name, fileData).then(
          r => { console.log('upload response', r); }
        )
      }
    )
  }

  private readFileDataAsText(file: any): Observable<string> {
    return new Observable(obs => {
      const reader = new FileReader();
      reader.onloadend = () => {
        obs.next(reader.result as string);
        obs.complete();
      }
      reader.readAsArrayBuffer(file);
    });
  }
}
