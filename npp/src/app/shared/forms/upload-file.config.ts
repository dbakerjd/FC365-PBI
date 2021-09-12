import {FormlyFieldConfig} from '@ngx-formly/core';
import { NPPFolder, SelectInputList } from 'src/app/services/sharepoint.service';

export class UploadFileConfig {
  
  constructor() {

  }

  fields(
    opportunityId: number, 
    stageId: number, 
    folders: NPPFolder[], 
    selectedFolder: number | null,
    countriesList: SelectInputList[], 
    scenariosList: SelectInputList[]): FormlyFieldConfig[] {
    let {categories, countries, scenarios} = this;

    let config = [
      {
        fieldGroup: [
          {
            key: 'StageNameId',
            defaultValue: stageId
          },
          {
            key: 'OpportunityNameId',
            defaultValue: opportunityId
          },
          {
            key: 'file',
            type: 'file-input',
            templateOptions: {
                label: 'File',
                placeholder: 'File',
                required: true
            },
          },
          categories(folders, selectedFolder),
          countries(countriesList, folders),
          scenarios(scenariosList, folders),
          {
            key: 'description',
            type: 'textarea',
            placeholder: 'Description',
            templateOptions: {
                label: 'Description:',
                rows: 2
            }
          },
          
        ]
      }
    ];

    return config;
  }

  categories(folders: NPPFolder[], defaultFolder: number | null) {
    return {
        key: 'category',
        type: 'select',
        templateOptions: {
            label: 'Categories:',
            options: folders.map((f: NPPFolder) => {
                return {
                    'name': f.Title,
                    'value': f.ID,
                };
            }),
            valueProp: 'value',
            labelProp: 'name',
            required: true,
        },
        defaultValue: defaultFolder
    }
  }

  countries(options: SelectInputList[], folders: NPPFolder[]) {
    return {
        key: 'country',
        type: 'ngsearchable',
        templateOptions: {
            label: 'Countries:',
            filterLocally: true,
            options: options,
            multiple: true,
        },
        "hideExpression": (model: any) => {
          return !folders.find(f => f.ID === model.category)?.containsModels;
        },
    }
  }


  scenarios(options: SelectInputList[], folders: NPPFolder[]) {
    return {
        key: 'scenario',
        type: 'ngsearchable',
        templateOptions: {
            label: 'Scenarios:',
            options: options,
            multiple: true,
            placeholder: 'Choose scenarios',
            required: true
        },
        "hideExpression": (model: any) => {
          return !folders.find(f => f.ID === model.category)?.containsModels;
        },
    }
  }

}
