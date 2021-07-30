import {FormlyFieldConfig} from '@ngx-formly/core';
import { NPPFolder } from 'src/app/services/sharepoint.service';
import { CountryList } from '../countries';

export class UploadFileConfig {
  
  constructor() {

  }

  fields(folders: NPPFolder[]): FormlyFieldConfig[] {
    let {categories, countries, scenarios} = this;

    let config = [
      {
        fieldGroup: [
          {
            key: 'name',
            type: 'input',
            templateOptions: {
                label: 'File Name:',
                placeholder: 'File Name'
            }
          },{
            key: 'file',
            type: 'file-input',
            templateOptions: {
                label: 'File',
                placeholder: 'File'
            }
          },
          categories(folders),
          countries(),
          scenarios(),
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
                    'name': f.name,
                    'value': f.id,
                };
            }),
            valueProp: 'value',
            labelProp: 'name'
        }
    }
  }

  countries() {
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
        "hideExpression": (model:any) => model.category  !== 6
    }
  }

  scenarios() {
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
        "hideExpression": (model:any) => model.category  !== 6
    }
  }
}
