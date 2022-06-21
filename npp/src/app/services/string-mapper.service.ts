import { Injectable } from '@angular/core';
import { StringMapping } from '@shared/models/app-config';
import { AppDataService } from './app/app-data.service';

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

  getString(key: string): string {
    if (!this.mappingList) return key;
    const mappedString = this.mappingList.find(e => e.Key === key);
    return mappedString ? mappedString.Title : key;
  }
}
