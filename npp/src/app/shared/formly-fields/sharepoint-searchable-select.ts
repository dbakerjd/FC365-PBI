import {Component} from "@angular/core";
import {debounceTime, distinctUntilChanged, filter, map, takeUntil, tap} from "rxjs/operators";
import {ReplaySubject, Subject} from "rxjs";
import {FormControl} from "@angular/forms";
import { HttpParams } from "@angular/common/http";
import { SharepointService } from "src/app/services/sharepoint.service";
import { FieldType } from "@ngx-formly/core";

@Component({
  selector: 'app-formly-field-searchable-select-api',
  template: `
    <mat-form-field>
      <mat-select [id]="id"
                  [formControl]="formControl"
                  [formlyAttributes]="field"
                  [multiple]="to.multiple"
                  [placeholder]="to.placeholder"
                  [errorStateMatcher]="errorStateMatcher"
                  [aria-labelledby]="formField?._labelId">
        <mat-option>
          <ngx-mat-select-search [formControl]="filterControl"
                                 [placeholderLabel]="(to.placeholderLabel || 'Search for items') | translate"
                                 [searching]="searching"
                                 [noEntriesFoundLabel]="(to.noEntriesFoundLabel || 'No matching items found') | translate">
          </ngx-mat-select-search>
        </mat-option>
        <ng-container *ngFor="let item of options$ | async">
          <mat-option [value]="item.value" [disabled]="item.disabled">{{ item.label }}</mat-option>
        </ng-container>
      </mat-select>
    </mat-form-field>
  `
})
export class FormlyFieldSearchableSelectApi extends FieldType {

  options$: ReplaySubject<any[]> = new ReplaySubject<any[]>(1);
  searching: boolean = false;

  filterControl: FormControl = new FormControl();
  /** Subject that emits when the component has been destroyed. */
  protected _onDestroy = new Subject<void>();

  constructor(private readonly api: SharepointService) {
    super();
  }

  ngOnInit() {

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

  async filterOptions(q: string, firstLoad?: boolean) {

    try {
      const { queryParams, modelsParser, partial, useCache, method } = this.to;

      let variables = {q};

      if (queryParams) {
        variables = {...variables, ...queryParams};
      }

      let url = partial;
      let opts = new HttpParams();

      switch(method) {
        case 'post':
          variables && Object.keys(variables).forEach((key) => {
            opts = opts.append(key, variables[key]);
          });
          break;
        case 'get':
        default:
          variables && Object.keys(variables).forEach((key) => {
            url = url.replace(':'+key, variables[key]);
          });
      }
      

      let res = await this.api[method](url, opts,  useCache);
      res = modelsParser ? modelsParser(res) : res;
      
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
