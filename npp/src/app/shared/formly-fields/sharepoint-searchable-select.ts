import {Component} from "@angular/core";
import {debounceTime, distinctUntilChanged, filter, map, takeUntil, tap} from "rxjs/operators";
import {ReplaySubject, Subject} from "rxjs";
import {FormControl} from "@angular/forms";
import { HttpParams } from "@angular/common/http";
import { SharepointService } from "src/app/services/sharepoint.service";
import { FieldType } from "@ngx-formly/core";
import { ErrorService } from "src/app/services/error.service";

/*
  templateOptions: {
    modelParser: method that receives sharepoint response and returns object {id, value}, value will display on screen
    returnObjects?: boolean, //should return the whole object instead of just IDs, default true
    query?: string, //url to hit for values, default siteusers
    filterLocally?: boolean, //should query all and filter locally, default false,
    filterField?: string, //field name to filter by, default title
  }
*/

@Component({
  selector: 'app-formly-field-searchable-select-api',
  template: `

    <input type="hidden" [formControl]="formControl" [formlyAttributes]="field">
  `
})
export class FormlyFieldSearchableSelectApi extends FieldType {

  options$: ReplaySubject<any[]> = new ReplaySubject<any[]>(1);
  searching: boolean = false;
  items: any[] = [];
  parsedItems: {id:number, value: string}[] = [];
  selectedItems: {id:number, value: string}[] = [];
  returnObjects = true;
  query = 'siteusers?'
  filterLocally = false;
  filterField = 'title';
  modelsParser: any;

  //title:string("much nothing",
  
  filterControl: FormControl = new FormControl();
  /** Subject that emits when the component has been destroyed. */
  protected _onDestroy = new Subject<void>();

  constructor(private readonly api: SharepointService, private readonly error: ErrorService) {
    super();
  }

  ngOnInit() {

    const { filterField, filterLocally, returnObjects, modelsParser, query } = this.to;
    this.filterField = filterField ? filterField : this.filterField;
    this.filterLocally = filterLocally ? filterLocally : this.filterLocally;
    this.returnObjects = (returnObjects === undefined || returnObjects === null) ? this.returnObjects : returnObjects;
    this.modelsParser = modelsParser ? modelsParser : this.defaultModelParser;
    this.query = query ? query : this.query;

    this.filterOptions('', true);

    // Listen for search field value changes
    this.filterControl.valueChanges
      .pipe(
        filter(q => [1,2].indexOf(q.length) < 0), // Only filter with minimum 3 chars search term or empty string
        tap(() => this.searching = true),
        takeUntil(this._onDestroy),
        debounceTime(500),
        distinctUntilChanged()
      )
      .subscribe(q => {
        this.filterOptions(q);
      });

    // Listen for selection changes
    this.formControl.valueChanges
      .subscribe(value => {
        if (this.to.onChange) {
          this.to.onChange(value);
        }
      });

  }

  defaultModelParser(el: any) {
    return {
      id: el.id,
      value: el.title
    }
  }

  async filterOptions(q: string, firstLoad?: boolean) {

    try {
      const { filterField, filterLocally, returnObjects, modelsParser, query, parsedItems } = this;
      
      let res;

      if (firstLoad || !filterLocally) {
        let partial = query;
        if(q) {
          partial+='$filter='+q;
        } 

        s = modelsParser(await this.api.query(partial));

      } else {
        res = 
      }

      
      
      this.searching = false;

      if (firstLoad && this.formControl.value) {
        const initialOption = res.find(o => o.value === this.formControl.value.toString());
        this.formControl.setValue(initialOption.value);
      }

      this.options$.next(res);

    } catch(error) {
        this.searching = false;
        this.notificationService.handleError(error);
    };

  }

  ngOnDestroy() {
    this._onDestroy.next();
    this._onDestroy.complete();
  }

}
