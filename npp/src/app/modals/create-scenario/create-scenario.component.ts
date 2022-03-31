import { Component, Inject, OnInit } from '@angular/core';
import { FormControl, FormGroup } from '@angular/forms';
import { MatDialog, MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { InlineNppDisambiguationService } from 'src/app/services/inline-npp-disambiguation.service';
import { SelectInputList, SharepointService } from 'src/app/services/sharepoint.service';
import { NPPFile } from '@shared/models/file-system';

@Component({
  selector: 'app-create-scenario',
  templateUrl: './create-scenario.component.html',
  styleUrls: ['./create-scenario.component.scss']
})
export class CreateScenarioComponent implements OnInit {

  fields: FormlyFieldConfig[] = [];
  form: FormGroup = new FormGroup({});
  model: any = {};

  scenarios: SelectInputList[] = [];
  file: NPPFile | null = null;

  // flow control
  updating = false;

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<CreateScenarioComponent>,
    public matDialog: MatDialog,
    private readonly sharepoint: SharepointService,
    private readonly disambiguator: InlineNppDisambiguationService
  ) { }

  async ngOnInit(): Promise<void> {

    this.file = this.data.file;
    this.scenarios = await this.sharepoint.getScenariosList();

    this.fields = [{
      fieldGroup: [{
        key: 'scenario',
        type: 'ngsearchable',
        templateOptions: {
            label: 'New Scenario',
            options: this.scenarios,
            required: true,
            multiple: true
        }
      },
      {
        key: 'multipleFiles',
        type: 'checkbox',
        templateOptions: {
            label: 'Create one copy per scenario',
        },
        hideExpression: (model: any) => {
          return model.scenario ? model.scenario.length < 2 : true;
        },
        defaultValue: false
      },
      {
        key: 'comments',
        type: 'textarea',
        templateOptions: {
            label: 'Comments:',
            placeholder: 'Please enter comments',
            rows: 3
        }
      }]
    }];

  }

  async onSubmit() {
    if (this.form.invalid || !this.file) {
      this.validateAllFormFields(this.form);
      return;
    }

    this.updating = this.dialogRef.disableClose = true;
    let success = false;

    const destinationFolder = this.file.ServerRelativeUrl.replace('/' + this.file.Name, '/');
    if (this.model.multipleFiles) {
      success = true;
      for (const scenId of this.model.scenario) {
        const newFileName = await this.sharepoint.addScenarioSufixToFilename(this.file.Name, scenId);
        if (newFileName) {
          success = success && await this.createScenario(newFileName, destinationFolder, [scenId]);
        }
      }
    } else {
      // clone in one file
      let newFileName = this.file.Name;
      for (const scenId of this.model.scenario) {
        const filenameSuffixed = await this.sharepoint.addScenarioSufixToFilename(newFileName, scenId);
        if (filenameSuffixed) newFileName = filenameSuffixed;
      }
      success = await this.createScenario(newFileName, destinationFolder, this.model.scenario);
    }
    this.updating = this.dialogRef.disableClose = false;
    this.dialogRef.close(success);
  }

  private async createScenario(fileName: string, destinationFolder: string, scenarios: number[]): Promise<boolean> {
    let success = false;
    let attemps = 0;
    const extension = fileName.split('.').pop();
    if (!extension) return false;
    const baseFileName = fileName.substring(0, fileName.length - (extension.length + 1));

    while (await this.sharepoint.existsFile(fileName, destinationFolder) && ++attemps < 11) {
      fileName = baseFileName + '-copy-' + attemps + '.' + extension;
    }

    if (attemps > 10) {
      success = false;
    } else {
      if (this.file) {
        let commentsStr = '';
        if(this.model.comments) {
          commentsStr = await this.sharepoint.addComment(this.file, this.model.comments);
        } else {
          commentsStr = this.file.ListItemAllFields?.Comments ? this.file.ListItemAllFields?.Comments : '';
        }
        success = await this.sharepoint.cloneEntityForecastModel(this.file, fileName, scenarios, (await this.sharepoint.getCurrentUserInfo()).Id, commentsStr);
      }
    }
    return success;
  }

  private validateAllFormFields(formGroup: FormGroup) {
    Object.keys(formGroup.controls).forEach(field => {
      const control = formGroup.get(field);
      if (control instanceof FormControl) {
        control.markAsTouched({ onlySelf: true });
        control.markAsDirty({ onlySelf: true });
      } else if (control instanceof FormGroup) {
        this.validateAllFormFields(control);
      }
    });
  }

}
