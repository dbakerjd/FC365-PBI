import {FormlyFieldConfig} from '@ngx-formly/core';
import { NPPFolder } from 'src/app/services/sharepoint.service';
import { CountryList } from '../countries';

export class UploadFileConfig {
  
  constructor() {

  }

  fields(opportunityId: number, stageId: number, folders: NPPFolder[]): FormlyFieldConfig[] {
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
          countries(folders),
          scenarios(folders),
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

  countries(folders: NPPFolder[]) {
    return {
        key: 'country',
        type: 'select',
        templateOptions: {
            label: 'Countries:',
            options: Object.keys(CountryList).map((key: string) => {
                return {
                    'name': (CountryList as unknown as  {[key:string]: string;})[key],
                    'value': key,
                };
            }),
            valueProp: 'value',
            labelProp: 'name'
        },
        "hideExpression": (model: any) => {
          return folders.find(f => f.ID === model.category)?.Title !== 'Forecast Models';
        },
    }
  }

  scenarios(folders: NPPFolder[]) {
    return {
        key: 'scenario',
        type: 'select',
        templateOptions: {
            label: 'Scenarios:',
            options: [{
                name: 'Base Case',
                value: '1'
            },{
                name: 'Upside',
                value: '2'
            },{
                name: 'Downside',
                value: '3'
            }],
            valueProp: 'value',
            labelProp: 'name'
        },
        "hideExpression": (model: any) => {
          return folders.find(f => f.ID === model.category)?.Title !== 'Forecast Models';
        },
    }
  }

}
