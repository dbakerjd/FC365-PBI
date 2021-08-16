import { Pipe, PipeTransform } from '@angular/core';

@Pipe({
  name: 'filter'
})
export class FilterPipe implements PipeTransform {

  transform(list: any[], titleFilter?: string, statusFilter?: string, typeFilter?: string, indicationFilter?: number): any {
    if (titleFilter) {
      list = [...list.filter(e => e.Title.search(new RegExp(titleFilter, 'i')) > -1)];
    }
    
    if (statusFilter) {
      statusFilter = statusFilter.toLocaleLowerCase();
      list = [...list.filter(e => e.OpportunityStatus.toLocaleLowerCase().indexOf(statusFilter) > -1)];
    }
    
    if (typeFilter) {
      list = [...list.filter(e => e.OpportunityTypeId === typeFilter)];
    }

    if (indicationFilter) {
      list = [...list.filter(e => e.IndicationId === indicationFilter)];
    }
    
    return list;
  }


}
