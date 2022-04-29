import {Component} from "@angular/core";
import {debounceTime, distinctUntilChanged, filter, map, takeUntil, tap} from "rxjs/operators";
import {ReplaySubject, Subject} from "rxjs";
import {FormControl} from "@angular/forms";
import { HttpParams } from "@angular/common/http";
import { SharepointService } from "src/app/services/microsoft-data/sharepoint.service";
import { FieldType } from "@ngx-formly/core";
import { ErrorService } from "@services/app/error.service";

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
    <div class="searchable-field-wrapper">
      <div class="searchable-field-selected-value" *ngFor="let item of selectedItems">
        {{item.value}}
        <span (click)="unselect(item)">X</span>
      </div>
      <input class="searchable-field-input" [formControl]="filterControl">
      <div class="searchable-field-options" *ngIf="displayOptions">
        <div class="searchable-field-option" (click)="select(option)" *ngFor="let option of parsedItems">{{option.value}}</div>
      </div>
    </div>
    <input type="hidden" [formControl]="formControl" [formlyAttributes]="field">
  `
})
export class FormlyFieldSearchableSelectApi extends FieldType {

  searching: boolean = false;
  items: any[] = [];
  parsedItems: {id:number, value: string}[] = [];
  filteredItems: {id:number, value: string}[] = [];
  selectedItems: {id:number, value: string}[] = [];
  idField = 'Id';
  returnObjects = true;
  query = 'siteusers'
  filterLocally = false;
  filterField = 'Title';
  modelsParser: any;
  displayOptions = false;

  //title:string("much nothing",
  
  filterControl: FormControl = new FormControl();
  /** Subject that emits when the component has been destroyed. */
  protected _onDestroy = new Subject<void>();

  constructor(private readonly api: SharepointService, private readonly error: ErrorService) {
    super();
  }

  ngOnInit() {

    const { idField, filterField, filterLocally, returnObjects, modelsParser, query } = this.to;
    this.idField = idField ? idField : this.idField;
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

  select(item: {id: number, value: string}) {
    this.displayOptions = false;
    let newItem =  this.parsedItems.find(i => i.id == item.id);
    let existingItem = this.selectedItems.find(i => i.id == item.id);

    if(!existingItem && newItem) {
      this.selectedItems.push(newItem);
      this.triggerValueChanged();
    }
    
  }

  unselect(item: {id: number, value: string}) {
    this.selectedItems = this.selectedItems.filter(el => el.id != item.id);
    this.triggerValueChanged();
  }

  triggerValueChanged() {
    let selectedItems = this.selectedItems.map(item => this.items.find(el => el[this.idField] == item.id));
    this.formControl.setValue(selectedItems);
  }

  defaultModelParser(el: any) {
    return {
      id: el[this.idField],
      value: el.Title
    }
  }

  initSelected() {
    this.selectedItems = [];
    this.formControl.value && this.formControl.value.forEach && this.formControl.value.forEach((el:any) => {
      let item =  this.items.find(i => (this.returnObjects && i[this.idField] == el[this.idField]) || (!this.returnObjects && i == el[this.idField]));
      if(item) this.selectedItems.push(this.modelsParser(item));
    })
  }

  async filterOptions(q: string, firstLoad?: boolean) {

    try {
      const { filterField, filterLocally, returnObjects, query } = this;
      
      let res;

      if (firstLoad || !filterLocally) {
        
        let partial = query;
        //this.items = await this.api.query(partial, '$filter=Title eq '+q);
        this.items = await this.api.query(partial, '').toPromise();
        this.parsedItems = this.items.map(el => this.modelsParser(el));
        this.filteredItems = this.items;
        if(!this.selectedItems) this.initSelected();

      } else {
        this.filteredItems = this.items.filter(el => el[this.filterField].indexOf(q) != -1);
      }
      
      if(!firstLoad) this.displayOptions = true;
      this.searching = false;

    } catch(error) {
      this.searching = false;
      this.error.handleError(error);
    };

  }

  ngOnDestroy() {
    this._onDestroy.next();
    this._onDestroy.complete();
  }

}
