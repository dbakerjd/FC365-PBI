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
  selector: 'app-formly-field-users-ng-select',
  template: `
    <ng-select [items]="to.options"
      [bindLabel]="labelProp"
      [bindValue]="valueProp"
      [multiple]="to.multiple"
      [placeholder]="to.placeholder"
      [formControl]="formControl" *ngIf="filterLocally">
        <ng-template ng-option-tmp let-item="item">
            {{item.name}} <br/>
            <small>{{item.gender}}</small>
        </ng-template>
    </ng-select>
    <ng-select [items]="to.options | async"
      [bindLabel]="labelProp"
      [bindValue]="valueProp"
      [multiple]="to.multiple"
      [placeholder]="to.placeholder"
      [formControl]="formControl"
      [trackByFn]="trackByFn"
      [minTermLength]="2"
      [loading]="searching"
      typeToSearchText="Please enter 2 or more characters"
      [typeahead]="textInput$" *ngIf="!filterLocally">
        <ng-template ng-option-tmp let-item="item">
            {{item.label}} <br/>
            <small>{{item.email}}</small>
        </ng-template>

        <ng-template ng-label-tmp let-item="item" let-clear="clear">
          <span class="ng-value-label">
            <!--<img [src]="item.avatar_url" width="20px" height="20px">--> {{item.label}}
          </span>
          <span class="ng-value-icon right" (click)="clear(item)" aria-hidden="true">×</span>
        </ng-template>

    </ng-select>
  `
})
export class FormlyFieldUsersNgSelect extends FieldType {

  textInput$ = new Subject<string>();
  searching: boolean = false;
  query = '';
  filterLocally = true;
  filterField = 'Title';

  filterControl: FormControl = new FormControl();

  constructor(private readonly appData: AppDataService, private readonly error: ErrorService) {
    super();
  }

  ngOnInit() {

    const { filterField, filterLocally, query, options} = this.to;

    this.filterField = filterField ? filterField : this.filterField;
    this.filterLocally = filterLocally === undefined ? this.filterLocally : filterLocally;
    this.query = query ? query : this.query;

    if (!this.filterLocally && this.query) {

      this.to.options = concat(
        of(this.to.options ? this.to.options : []), // default items
        this.textInput$.pipe(
          distinctUntilChanged(),
          debounceTime(500),
          tap(() => this.searching = true),
          switchMap(term => this.appData.searchByTermInputList(this.query, this.filterField, term).pipe(
            catchError(() => of([])), // empty list on error
            tap(() => this.searching = false)
          ))
        )
      ) as Observable<any>;

    }
  
  }

  trackByFn(item: any) {
    return item.Id;
  }

  get labelProp(): string { return this.to.labelProp || 'label'; }
  get valueProp(): string { return this.to.valueProp || 'value'; }
  get groupProp(): string { return this.to.groupProp || 'group'; }

}
