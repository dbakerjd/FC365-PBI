import {FormlyFieldConfig} from '@ngx-formly/core';
import { Country, NPPFolder, SelectInputList } from 'src/app/services/sharepoint.service';
import { CountryList } from '../countries';

export class UploadFileConfig {
  
  constructor() {

  }

  fields(
    opportunityId: number, 
    stageId: number, 
    folders: NPPFolder[], 
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
            },
          },
          categories(folders),
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

  categories(folders: NPPFolder[]) {
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
        }
    }
  }

  countries(options: SelectInputList[], folders: NPPFolder[]) {
    return {
        key: 'country',
        type: 'select',
        templateOptions: {
            label: 'Countries:',
            options: options,multiple: true,
        },
        "hideExpression": (model: any) => {
          return !folders.find(f => f.ID === model.category)?.containsModels;
        },
    }
  }

  /*
  scenarios(options: SelectInputList[], folders: NPPFolder[]) {
    return {
        key: 'scenario',
        type: 'select',
        templateOptions: {
            label: 'Scenarios:',
            options: options,
            multiple: true,
        },
        "hideExpression": (model: any) => {
          return !folders.find(f => f.ID === model.category)?.containsModels;
        },
    }
  }
  */

  scenarios(options: SelectInputList[], folders: NPPFolder[]) {
    return {
        key: 'scenario',
        type: 'ngsearchable',
        templateOptions: {
            label: 'Scenarios:',
            options: options,
            multiple: true,
        },
        "hideExpression": (model: any) => {
          return !folders.find(f => f.ID === model.category)?.containsModels;
        },
    }
  }

}
