import { Component, Inject, OnInit } from '@angular/core';
import { FormControl, FormGroup } from '@angular/forms';
import { MatDialog, MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { NPPFile } from '@shared/models/file-system';
import { AppDataService } from '@services/app/app-data.service';
import { FilesService } from 'src/app/services/files.service';
import { SelectInputList } from '@shared/models/app-config';
import { SelectListsService } from '@services/select-lists.service';
import { ConfirmDialogComponent } from '../confirm-dialog/confirm-dialog.component';
import { take } from 'rxjs/operators';
import { forkJoin, Observable } from 'rxjs';

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
    private readonly appData: AppDataService,
    private readonly files: FilesService,
    private readonly selectLists: SelectListsService
  ) { }

  async ngOnInit(): Promise<void> {

    this.file = this.data.file;
    this.scenarios = await this.selectLists.getScenariosList();

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
    const scenarios = await this.appData.getMasterScenarios();
    
    if (this.model.multipleFiles) {
      success = true;
      let dialogsObs: Observable<any>[] = [];

      for (const scenId of this.model.scenario) {
        const existsFileWithSameScenarios = await this.files.getFileByScenarios(destinationFolder, [scenId]);
        if (existsFileWithSameScenarios) {
          const currentScenario = scenarios.find(s => s.ID == scenId);
          const dialogRef = this.matDialog.open(ConfirmDialogComponent, {
            maxWidth: "400px",
            height: "250px",
            data: {
              message: `A model with the scenario ${currentScenario?.Title} already exists. Do you want to overwrite it?`,
              confirmButtonText: 'Yes, overwrite',
              cancelButtonText: 'No, keep the original',
              reference: scenId
            }
          });
          dialogsObs.push(dialogRef.afterClosed());
        }
      }

      forkJoin(dialogsObs)
        .subscribe(async confirmations => {
          for (const scenId of this.model.scenario) {
            const confirmation = confirmations.find(c => c.reference == scenId);
            if (!confirmation || (confirmation && confirmation.result)) {
              const newFileName = await this.files.addScenarioSufixToFilename(this.file!.Name, scenId);
              if (newFileName) {
                success = success && await this.createScenario(newFileName, destinationFolder, [scenId]);
              }
            }
          }
          this.updating = this.dialogRef.disableClose = false;
          this.dialogRef.close(success);
        });

    } else {
      success = true;
      const existsFileWithSameScenarios = await this.files.getFileByScenarios(destinationFolder, this.model.scenario);
      
      if (existsFileWithSameScenarios) {
        const scenariosNames = this.model.scenario.map((scenId: number) => {
          const scenario = scenarios.find(s => s.ID == scenId);
          if (scenario) return scenario.Title;
          else return "";
        });
        let modalMessage = '';
        if (scenariosNames.length > 1) {
          modalMessage = `A model with the same scenarios (${scenariosNames.join(', ')}) already exists. Do you want to overwrite it?`
        } else {
          modalMessage = `A model with the same scenario (${scenariosNames}) already exists. Do you want to overwrite it?`
        }
        const dialogRef = this.matDialog.open(ConfirmDialogComponent, {
          maxWidth: "400px",
          height: "250px",
          data: {
            message: `A model with the same scenarios (${scenariosNames}) already exists. Do you want to overwrite it?`,
            confirmButtonText: 'Yes, overwrite',
            cancelButtonText: 'No, keep the original'
          }
        });
      
        dialogRef.afterClosed()
          .pipe(take(1))
          .subscribe(async overwrite => {
            if (overwrite) {
              await this.cloneInOneFile(this.file!.Name, destinationFolder);
            } else {
              // do nothing and close
              this.updating = this.dialogRef.disableClose = false;
              this.dialogRef.close(true);
            }
          });
      } else {
        await this.cloneInOneFile(this.file.Name, destinationFolder);
      }
    }
  }

  private async cloneInOneFile(filename: string, destinationFolder: string) {
    for (const scenId of this.model.scenario) {
      const filenameSuffixed = await this.files.addScenarioSufixToFilename(filename, scenId);
      if (filenameSuffixed) filename = filenameSuffixed;
    }
    const success = await this.createScenario(filename, destinationFolder, this.model.scenario);
    this.updating = this.dialogRef.disableClose = false;
    this.dialogRef.close(success);
  }

  private async createScenario(fileName: string, destinationFolder: string, scenarios: number[]): Promise<boolean> {
    let success = false;
    let attemps = 0;
    const extension = fileName.split('.').pop();
    if (!extension) return false;
    const baseFileName = fileName.substring(0, fileName.length - (extension.length + 1));

    while (await this.appData.existsFile(fileName, destinationFolder) && ++attemps < 11) {
      fileName = baseFileName + '-copy-' + attemps + '.' + extension;
    }

    if (attemps > 10) {
      success = false;
    } else {
      if (this.file) {
        let commentsStr = '';
        if(this.model.comments) {
          commentsStr = await this.files.firstCommentString(this.model.comments);
        }
        success = await this.files.cloneForecastModel(this.file, fileName, scenarios, (await this.appData.getCurrentUserInfo()).Id, commentsStr);
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
