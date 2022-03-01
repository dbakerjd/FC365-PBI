import { Component, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { MatDialog } from '@angular/material/dialog';
import { Router } from '@angular/router';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { ToastrService } from 'ngx-toastr';
import { Subject } from 'rxjs';
import { debounceTime, take, takeUntil, tap } from 'rxjs/operators';
import { CreateBrandComponent } from 'src/app/modals/create-brand/create-brand.component';
import { CreateForecastCycleComponent } from 'src/app/modals/create-forecast-cycle/create-forecast-cycle.component';
import { InlineNppDisambiguationService } from 'src/app/services/inline-npp-disambiguation.service';
import { Indication, Opportunity, SelectInputList, SharepointService, User } from 'src/app/services/sharepoint.service';
import { TeamsService } from 'src/app/services/teams.service';

@Component({
  selector: 'app-brand-list',
  templateUrl: './brand-list.component.html',
  styleUrls: ['./brand-list.component.scss']
})
export class BrandListComponent implements OnInit {
  private readonly _destroying$ = new Subject<void>();
  currentUser: User | undefined = undefined;
  brands: Opportunity[] = [];
  filteredBrands: Opportunity[] = [];
  form = new FormGroup({});
  model: any = { };
  fields: FormlyFieldConfig[] = [];
  dialogInstance: any;
  masterCycles: SelectInputList[] = [];
  updateSearchTimeout: any; 
  loading = true;

  constructor(private sharepoint: SharepointService, private teams: TeamsService, private router: Router, public matDialog: MatDialog, private toastr: ToastrService, public disambiguator: InlineNppDisambiguationService) { }

  async ngOnInit() {
    if(this.teams.initialized) this.init();
    else {
      this.teams.statusSubject.subscribe(async (msg) => {
        setTimeout(async () => {
          this.init();
        }, 500);
      });
    }
    
  }

  async init() {
    this.currentUser = await this.sharepoint.getCurrentUserInfo();
    let indicationsList: SelectInputList[] = [];
    let forecastCycles = await this.sharepoint.getForecastCycles();
    let businessUnits = await this.sharepoint.getBusinessUnitsList();
    let brandFields = await this.sharepoint.getBrandFields();
    let therapies = await this.sharepoint.getTherapiesList();
    this.masterCycles = await this.sharepoint.getForecastCycles();
    
    this.fields = [{
        key: 'search',
        type: 'input',
        templateOptions: {
          placeholder: 'Search all brands',
        },
        hooks: {
          onInit: (field) => {
            if(field && field.formControl) {
              field.formControl.valueChanges.pipe(
                debounceTime(500),
                takeUntil(this._destroying$),
                tap(th => {
                  this.onSubmit();
                }),
              ).subscribe();
            }
          }
        }
      },{
        key: 'therapy',
        type: 'select',
        templateOptions: {
          placeholder: 'Therapy Area',
          options: therapies,
        }
      },{
        key: 'indication',
        type: 'select',
        templateOptions: {
          placeholder: 'Indication Name',
          options: indicationsList
        },
        hooks: {
          onInit: (field) => {
            if (!field?.parent?.fieldGroup) return;
            const therapySelect = field.parent.fieldGroup.find(f => f.key === 'therapy');
            if (!therapySelect?.formControl) return;
            therapySelect.formControl.valueChanges.pipe(
              takeUntil(this._destroying$),
              tap(th => {
                this.sharepoint.getIndicationsList(th).then(r => {
                  if (field.templateOptions) field.templateOptions.options = r;
                });
              }),
            ).subscribe();
          }
        }
      },{
        key: 'businessUnit',
        type: 'select',
        templateOptions: {
          placeholder: 'Business Unit',
          options: businessUnits
        }
      },{
        key: 'forecastCycle',
        type: 'select',
        templateOptions: {
          placeholder: 'Forecast Cycle',
          options: forecastCycles
        }
      },{
        key: 'sort_by',
        type: 'select',
        templateOptions: {
          placeholder: 'Sort by',
          options: brandFields
        }
      }
    ];

    this.brands = await this.disambiguator.getEntities() as Opportunity[];
    this.loading = false;

    //console.log(this.brands);
    this.onSubmit();
  }

  createBrand() {
    this.dialogInstance = this.matDialog.open(CreateBrandComponent, {
      height: '75vh',
      width: '500px'
    });

    this.dialogInstance.afterClosed().subscribe(async (result:Opportunity) => {
      this.brands = await this.disambiguator.getEntities() as Opportunity[];
      this.onSubmit();
    });
    
  }

  async editBrand(brand: Opportunity) {
    this.dialogInstance = this.matDialog.open(CreateBrandComponent, {
      height: '75vh',
      width: '500px',
      data: {
        brand
      }
    });

    this.dialogInstance.afterClosed()
    .pipe(take(1))
    .subscribe(async (result: any) => {
      if (result.success) {
        Object.assign(brand, await this.sharepoint.getBrand(brand.ID));
      } else if (result.success === false) {
        this.toastr.error("The brand couldn't be updated", "Try again");
      }
    });
  }

  getIndications(indications: Indication[]) {
    if(indications && indications.length) {
      return indications.map(e => e.Title).join(", ")
    } else return "";
  }

  onSubmit() {
    let list = [...this.brands];
    if (this.model.search) {
      list = list.filter(e => e.Title.search(new RegExp(this.model.search, 'i')) > -1);
    }
    
    if (this.model.businessUnit) {
      list = list.filter(e => e.BusinessUnitId === this.model.businessUnit);
    }
    
    if (this.model.forecastCycle) {
      list = list.filter(e => e.ForecastCycleId === this.model.forecastCycle);
    }

    if (this.model.therapy) {
      list = list.filter(e => (e.Indication && e.Indication.length && e.Indication[0].TherapyArea == this.model.therapy));
    }

    if (this.model.indication) {
      list = list.filter(e => (e.IndicationId.indexOf(this.model.indication) > -1));
    }

    if(this.model.sort_by) {
      list = list.sort((a: any,b:any) => {
        let fields = this.model.sort_by.split(".");
        if(fields.length == 1) {
          a = a[fields[0]].toLocaleLowerCase();
          b = b[fields[0]].toLocaleLowerCase();
        //length 2
        } else {
          
          if(a[fields[0]][fields[1]]) {
            a = a[fields[0]][fields[1]].toLocaleLowerCase();
          } else {
            if(a[fields[0]][0] && a[fields[0]][0][fields[1]]) {
              a = a[fields[0]][0][fields[1]].toLocaleLowerCase();
            } else {
              a = '';
            }
          }

          if(b[fields[0]][fields[1]]) {
            b = b[fields[0]][fields[1]].toLocaleLowerCase();
          } else {
            if(b[fields[0]][0] && b[fields[0]][0][fields[1]]) {
              b = b[fields[0]][0][fields[1]].toLocaleLowerCase();
            } else {
              b = '';
            }
          }
          
        }
        
        if(a < b) return -1;
        if(a > b) return 1;
        return 0;
      })
    }

    this.filteredBrands = list;
  }

  createForecast(brand: Opportunity) {
    this.dialogInstance = this.matDialog.open(CreateForecastCycleComponent, {
      height: '400px',
      width: '405px',
      data: {
        entity: brand
      }
    });

    this.dialogInstance.afterClosed()
      .pipe(take(1))
      .subscribe(async (success: any) => {
        if (success) {
          this.toastr.success(`The new forecast cycle has been created successfully`, "New Forecast Cycle");
          brand = Object.assign(brand, {
            ForecastCycleId: success.ForecastCycleId,
            ForecastCycle: { 
              Title: this.masterCycles.find(el => el.value == success.ForecastCycleId)?.label,
              ID: success.ForecastCycleId
            },
            Year: success.Year
        });
        } else if (success === false) {
          this.toastr.error('The new forecast cycle could not be created', 'Try Again');
        }
      });
  }

  navigateTo(item: Opportunity) {
    this.router.navigate(['brands', item.ID, 'files']);
  }

  ngOnDestroy(): void {
    this._destroying$.next();
    this._destroying$.complete();
  }
}
