import {Component} from "@angular/core";
import {catchError, debounceTime, distinctUntilChanged, filter, map, switchMap, takeUntil, tap} from "rxjs/operators";
import {concat, Observable, of, ReplaySubject, Subject} from "rxjs";
import {FormControl} from "@angular/forms";
import { HttpParams } from "@angular/common/http";
import { FieldType } from "@ngx-formly/core";
import { ErrorService } from "@services/app/error.service";
import { AppDataService } from "@services/app/app-data.service";

/*
  templateOptions: {
    query?: string, //url to hit for values, default ''
    filterLocally?: boolean, //should query all and filter locally, default false,
    filterField?: string, //field name to filter by, default title
  }
*/

@Component({
  selector: 'app-formly-field-ng-select-by-groups',
  template: `
    <ng-select [items]="to.options"
      [bindLabel]="labelProp"
      [bindValue]="valueProp"
      [groupBy]="groupProp"
      [multiple]="true"
      [compareWith]="compareIndications"
      [placeholder]="to.placeholder"
      [dropdownPosition]="'bottom'"
      [(ngModel)]="model.indication"
      appendTo="app-npp-header"
      [closeOnSelect] = "false"
      [selectableGroup]="true"
      [selectableGroupAsModel]="true">
      <ng-template ng-optgroup-tmp let-item="item">
        {{item.group || 'Unnamed group'}}
      </ng-template>
    </ng-select>
  `
})
export class FormlyFieldNgSelectByGroups extends FieldType {

  filterControl: FormControl = new FormControl();

  constructor(private readonly appData: AppDataService, private readonly error: ErrorService) {
    super();
  }

  ngOnInit() {

    const { options} = this.to;
  }

  // trackByFn(item: any) {
    // return item.Id;
  // }

  get labelProp(): string { return this.to.labelProp || 'label'; }
  get valueProp(): string { return this.to.valueProp || 'value'; }
  get groupProp(): string { return this.to.groupProp || 'group'; }

  compareIndications = (item: any, selected: any) => {
    if (selected.group && item.group) {
        return item.group === selected.group;
    }
    if (item.name && selected.name) {
        return item.name === selected.name;
    }
    return false;
};


}
