import {FormlyFieldConfig} from '@ngx-formly/core';
import { SelectInputList } from '@shared/models/app-config';
import { NPPFolder } from '@shared/models/file-system';

export class UploadFileConfig {
  
  constructor() {

  }

  fields(
    opportunityId: number, 
    stageId: number, 
    folders: NPPFolder[], 
    selectedFolder: number | null,
    geographiesList: SelectInputList[], 
    scenariosList: SelectInputList[],
    indicationsList: SelectInputList[]): FormlyFieldConfig[] {
    let {categories, geographies, scenarios, indications} = this;

    let config = [
      {
        fieldGroup: [
          {
            key: 'StageNameId',
            defaultValue: stageId
          },
          {
            key: 'EntityNameId',
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
          indications(indicationsList, folders),
          geographies(geographiesList, folders),
          scenarios(scenariosList, folders),
          {
            key: 'description',
            type: 'textarea',
            placeholder: 'Description',
            templateOptions: {
                label: 'Description:',
                rows: 3
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
                    'value': f.DepartmentID,
                };
            }),
            valueProp: 'value',
            labelProp: 'name',
            required: true,
        },
        defaultValue: defaultFolder
    }
  }

  geographies(options: SelectInputList[], folders: NPPFolder[]) {
    if (options.length === 1) {
      return {
        key: 'geography',
        type: 'input',
        defaultValue: options[0].value,
        "hideExpression": true,
      };
    }
    return {
        key: 'geography',
        type: 'ngsearchable',
        templateOptions: {
            label: 'Geography:',
            filterLocally: true,
            options: options,
            multiple: false,
            required: true
        },
        "hideExpression": (model: any) => {
          return !folders.find(f => f.DepartmentID === model.category)?.containsModels;
        },
    };
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
          return !folders.find(f => f.DepartmentID === model.category)?.containsModels;
        },
    }
  }


  indications(options: SelectInputList[], folders: NPPFolder[]) {
    return {
      key: 'IndicationId',
      type: 'ngsearchable',
      templateOptions: {
        label: 'Indication Name:',
        options,
        multiple: true,
        required: true,
      },
      "hideExpression": (model: any) => {
        return !folders.find(f => f.DepartmentID === model.category)?.containsModels;
      }
    }
  }

}
