import { Pipe, PipeTransform } from '@angular/core';
import { Opportunity } from './services/sharepoint.service';

@Pipe({
  name: 'filter'
})
export class FilterPipe implements PipeTransform {

  transform(list: Opportunity[], titleFilter?: string, statusFilter?: string, typeFilter?: string, indicationFilter?: number): any {
    if (titleFilter) {
      list = [...list.filter(e => e.Title.search(new RegExp(titleFilter, 'i')) > -1)];
    }
    
    if (statusFilter) {
      list = [...list.filter(e => e.OpportunityStatus.toLocaleLowerCase().indexOf(statusFilter.toLocaleLowerCase()) > -1)];
    }
    
    if (typeFilter) {
      list = [...list.filter(e => e.OpportunityTypeId === +typeFilter)];
    }

    if (indicationFilter) {
      list = [...list.filter(e => (e.IndicationId.indexOf(indicationFilter) > -1))];
    }
    
    return list;
  }


}
