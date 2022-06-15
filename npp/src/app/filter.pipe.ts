import { Pipe, PipeTransform } from '@angular/core';
import { Opportunity } from './shared/models/entity';

@Pipe({
  name: 'filter'
})
export class FilterPipe implements PipeTransform {

  transform(list: Opportunity[], titleFilter?: string, statusFilter?: string, typeFilter?: string, indicationFilter?: any): any {
    if (titleFilter) {
      list = [...list.filter(e => e.Title.search(new RegExp(titleFilter, 'i')) > -1)];
    }
    
    if (statusFilter) {
      list = [...list.filter(e => e.OpportunityStatus.toLocaleLowerCase().indexOf(statusFilter.toLocaleLowerCase()) > -1)];
    }
    
    if (typeFilter) {
      list = [...list.filter(e => e.OpportunityTypeId === +typeFilter)];
    }

    if (indicationFilter && indicationFilter.length > 0) {
      list = [...list.filter(e => indicationFilter.some((currentFilter: number | string) => { 
        // if currentFilter is a string, is the 'Therapy name' including all 'Indications'
        // if its a number is an individual 'Indication'
        if (typeof currentFilter === 'string') return e.Indication.some(i => i.TherapyArea === currentFilter);
        else return e.IndicationId.includes(currentFilter)
      }))];
    }
    
    return list;
  }


}
