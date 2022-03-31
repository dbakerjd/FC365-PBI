import { Pipe, PipeTransform } from '@angular/core';
import { Opportunity } from './shared/models/entity';

@Pipe({
  name: 'sort'
})
export class SortPipe implements PipeTransform {

  transform(list: Opportunity[], sortField: string): any {
    switch(sortField) {
      case 'title': 
        list.sort((a: Opportunity, b: Opportunity) => { return a.Title.toLocaleLowerCase() < b.Title.toLocaleLowerCase() ? -1 : 1 });
        break;
      case 'molecule': 
        list.sort((a: Opportunity, b: Opportunity) => { return a.MoleculeName.toLocaleLowerCase() < b.MoleculeName.toLocaleLowerCase() ? -1 : 1 });
        break;
      case 'indication': 
        list.sort((a: Opportunity, b: Opportunity) => { return a.Indication[0]?.Title.toLocaleLowerCase() < b.Indication[0]?.Title.toLocaleLowerCase() ? -1 : 1 });
        break;
      case 'projectStart':
        list.sort((a: Opportunity, b: Opportunity) => { 
          let aDate = new Date(a.ProjectStartDate).getTime();
          let bDate = new Date(b.ProjectStartDate).getTime();
          return aDate < bDate ? -1 : 1 
        });
        break;
      case 'projectEnd':
        list.sort((a: Opportunity, b: Opportunity) => { 
          let aDate = new Date(a.ProjectEndDate).getTime();
          let bDate = new Date(b.ProjectEndDate).getTime();
          return aDate < bDate ? -1 : 1 
        });
        break;
      case 'opportunityType':
        list.sort((a: Opportunity, b: Opportunity) => { 
          if (a.OpportunityType && b.OpportunityType) {
            return a.OpportunityType.Title < b.OpportunityType.Title ? -1 : 1 
          }
          return 0;
        });
        break;
      default:
        list.sort((a: Opportunity, b: Opportunity) => { 
          return a.ID < b.ID ? -1 : 1 
        });
    }
    return list;
  }

}
