import { Component, Inject, OnInit } from '@angular/core';
import { FormGroup } from '@angular/forms';
import { FormlyFieldConfig } from '@ngx-formly/core';
import { takeUntil, tap } from 'rxjs/operators';
import { Subject } from 'rxjs';
import { MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { WorkInProgressService } from '@services/app/work-in-progress.service';
import { ToastrService } from 'ngx-toastr';
import { EntityGeography, Opportunity } from '@shared/models/entity';
import { AppDataService } from '@services/app/app-data.service';
import { EntitiesService } from 'src/app/services/entities.service';
import { PermissionsService } from 'src/app/services/permissions.service';
import { SelectInputList } from '@shared/models/app-config';
import { SelectListsService } from '@services/select-lists.service';

@Component({
  selector: 'app-create-brand',
  templateUrl: './create-brand.component.html',
  styleUrls: ['./create-brand.component.scss']
})
export class CreateBrandComponent implements OnInit {

  private readonly _destroying$ = new Subject<void>();
  
  form = new FormGroup({});
  model: any = { };
  fields: FormlyFieldConfig[] = [];
  indications: any[] = [];
  dialogInstance: any;
  isEdit = false;
  brand: Opportunity | null = null;
  loading = true;
  updating = false;
  geographies: EntityGeography[] = [];

  constructor(
    @Inject(MAT_DIALOG_DATA) public data: any,
    public dialogRef: MatDialogRef<CreateBrandComponent>, 
    public jobs: WorkInProgressService, 
    public toastr: ToastrService,
    private readonly appData: AppDataService,
    private readonly entities: EntitiesService, 
    private readonly permissions: PermissionsService, 
    private readonly selectLists: SelectListsService,
    ) { }

  async ngOnInit() {
    this.loading = true;
    this.brand = this.data?.brand;
    //this.model.Brand = this.brand;

    let therapies = await this.selectLists.getTherapiesList();
    let forecastCycles = await this.selectLists.getForecastCyclesList();
    let businessUnits = await this.selectLists.getBusinessUnitsList();
    const geo = (await this.selectLists.getGeographiesList()).map(el => { return { label: el.label, value: 'G-' + el.value } });
    const countries = (await this.selectLists.getCountriesList()).map(el => { return { label: el.label, value: 'C-' + el.value } });;
    const locationsList = geo.concat(countries);
    let indicationsList: SelectInputList[] = [];
    let defaultUsersList: SelectInputList[] = await this.selectLists.getSiteOwnersList();

    if (this.brand) {
      //this.brand.FCDueDate = new Date(this.brand.FCDueDate);
      this.isEdit = true;
      this.geographies = await this.appData.getEntityGeographies(this.brand?.ID);
      this.model.geographies = this.geographies.map(el => el.CountryId ? 'C-'+el.CountryId : 'G-' + el.GeographyId);
    
      
      // default indications for the therapy selected
      if (this.brand && this.brand.Indication && this.brand.Indication.length) {
        indicationsList = await this.selectLists.getIndicationsList(this.brand.Indication[0].TherapyArea);
      }
    }

    const currentYear = new Date().getFullYear();
    let year = currentYear;
    let elegibleYears = [currentYear];
    for(let i=1; i<6; i++) {
      elegibleYears.push(++year);
    }
    

    this.fields = [
      {
        fieldGroup: [{
          key: 'Brand.Title',
          type: 'input',
          templateOptions: {
            label: 'Brand Name:',
            placeholder: 'Brand Name',
            required: true,
          },
          defaultValue: this.brand?.Title
        }, {
          key: 'Brand.EntityOwnerId',
          type: 'ngsearchable',
          templateOptions: {
            label: 'Brand Owner:',
            placeholder: 'Brand Owner',
            required: true,
            options: defaultUsersList
            /*filterLocally: false,
            query: 'siteusers'*/
          },
          defaultValue: this.brand?.EntityOwnerId
        }, {
          key: 'therapy',
          type: 'select',
          templateOptions: {
            label: 'Therapy Area:',
            options: therapies,
            required: true,
          },
          defaultValue: this.brand && this.brand.Indication && this.brand.Indication.length ? this.brand.Indication[0].TherapyArea : null
        },{
          key: 'Brand.IndicationId',
          type: 'ngsearchable',
          templateOptions: {
            label: 'Indication Name:',
            options: indicationsList,
            multiple: true,
            required: true,
          },
          defaultValue: this.brand?.IndicationId,
          hooks: {
            onInit: (field) => {
              if (!field?.parent?.fieldGroup) return;
              const therapySelect = field.parent.fieldGroup.find(f => f.key === 'therapy');
              if (!therapySelect?.formControl) return;
              therapySelect.formControl.valueChanges.pipe(
                takeUntil(this._destroying$),
                tap(th => {
                  this.selectLists.getIndicationsList(th).then(r => {
                    if (r.length > 0) field.formControl?.setValue(r[0].value);
                    if (field.templateOptions) field.templateOptions.options = r;
                  });
                }),
              ).subscribe();
            }
          }
        }, {
          key: 'Brand.BusinessUnitId',
          type: 'select',
          templateOptions: {
            label: 'Business Unit:',
            options: businessUnits,
            required: true,
          },
          defaultValue: this.brand?.BusinessUnitId
        }, {
          key: 'Brand.ForecastCycleId',
          type: 'select',
          templateOptions: {
            label: 'Forecast Cycle:',
            options: forecastCycles,
            required: true,
          },
          defaultValue: this.brand?.ForecastCycleId
        }, /*{
          key: 'Brand.FCDueDate',
          type: 'datepicker',
          className: 'date-input',
          templateOptions: {
            label: 'Forecast Cycle Due Date:',
            required: true,
          },
          defaultValue: this.brand?.FCDueDate ? new Date(this.brand?.FCDueDate) : null
        },*/{
          key: 'Brand.Year',
          type: 'select',
          templateOptions: {
            label: 'Year:',
            options: elegibleYears.map(el => {
              return {
                label: el,
                value: el
              }
            }),
            required: true,
          },
          defaultValue: this.brand?.Year || currentYear
        }, 
        {
          key: 'Brand.ForecastCycleDescriptor',
          type: 'input',
          templateOptions: {
            label: 'Forecast Cycle Descriptor',
            required: false
          },
          defaultValue: this.brand?.ForecastCycleDescriptor
        },
        {
          key: 'geographies',
          type: 'ngsearchable',
          templateOptions: {
            label: 'Geographies:',
            placeholder: 'Related geographies and countries',
            required: true,
            multiple: true,
            options: locationsList
          }
        }]
      }
    ];

    this.loading = false;
  }

  async onSubmit() {
    let job = this.jobs.startJob(
      "Creating Brand"
      );
    try {
      if (this.isEdit) {
        this.updating = this.dialogRef.disableClose = true;
        await this.permissions.updateEntityGeographies(this.data.brand, this.model.geographies);
        const success = await this.entities.updateEntity(this.data.brand.ID, this.model.Brand);
        this.updating = this.dialogRef.disableClose = false;
        this.jobs.finishJob(job.id);
        this.toastr.success("The brand has been updated", this.model.Brand.Title);
        this.dialogRef.close({
          success: success,
          data: this.model.Brand
        });
      } else {
        // force opportunity type
        const internalType = (await this.appData.getOpportunityTypes()).find(el => el.IsInternal);
        this.model.Brand.OpportunityTypeId = internalType?.ID;
        let brand = await this.entities.createBrand(
          this.model.Brand,
          this.model.geographies.filter((el: string) => el.startsWith('G-')).map((el: string) => +el.substring(2)),
          this.model.geographies.filter((el: string) => el.startsWith('C-')).map((el: string) => +el.substring(2)));
        this.jobs.finishJob(job.id);
        this.toastr.success("The brand is now active", brand?.Title);
        this.dialogRef.close(brand);
        

        this.dialogRef.close({
          success: brand ? true : false,
          data: brand
        });
      }    
    } catch(e: any) {
      this.updating = false;
      this.jobs.finishJob(job.id);
      this.toastr.error(e.message);
    }

    
  }
 
  ngOnDestroy(): void {
    this._destroying$.next();
    this._destroying$.complete();
  }
}


