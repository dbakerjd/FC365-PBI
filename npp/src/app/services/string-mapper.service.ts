import { Injectable } from '@angular/core';
import { StringMapping } from '@shared/models/app-config';
import { AppDataService } from './app/app-data.service';


type capitalLetterParam =  'l' | 'lower' | 'u' | 'upper' | 'f' | 'first' | 't' | 'titlecase'; 
@Injectable({
  providedIn: 'root'
})
export class StringMapperService {

  private mappingList: StringMapping[] | undefined = undefined;

  constructor(private readonly appData: AppDataService) { 
    this.loadList(); 
  }

  private async loadList() {
    this.mappingList = await this.appData.getStringMappingItems();
  }

  getString(key: string, capitalLetters: capitalLetterParam = 'titlecase'): string {
    if (!this.mappingList) return key;
    const lowerCaseKey = key.toLocaleLowerCase();
    const mappedString = this.mappingList.find(e => e.Key.toLocaleLowerCase() === lowerCaseKey);
    if (mappedString) {
      switch (capitalLetters) {
        case 'first':
        case 'f':
          return mappedString.Title[0].toLocaleUpperCase() + mappedString.Title.slice(1);
        case 'upper':
        case 'u':
          return mappedString.Title.toLocaleUpperCase();
        case 'lower':
        case 'l':
          return mappedString.Title.toLocaleLowerCase();
        case 'titlecase':
        case 't':
          return this.titleCase(mappedString.Title);
      }
    }
    return key;
  }

  private titleCase(str: string) {
    return str.split(' ')
      .map(w => w[0].toLocaleUpperCase() + w.substring(1).toLocaleLowerCase())
      .join(' ');
  }
}
