import { CommonModule } from "@angular/common";
import { NgModule } from "@angular/core";
import { FormsModule, ReactiveFormsModule } from "@angular/forms";
import { NgSelectModule } from "@ng-select/ng-select";
import { FormlyModule } from "@ngx-formly/core";
import { DatepickerModule } from "ng2-datepicker";
import { FormlyFieldDatePicker } from "./date-picker";
import { FormlyFieldFile } from "./file-input";
import { FileValueAccessor } from "./file-value-accessor";
import { FormlyFieldNgSelect } from "./ng-select-input";
import { FormlyFieldSearchableSelectApi } from "./sharepoint-searchable-select";

export const FORMLY_CONFIG = {
  types: [
    { name: 'file-input', component: FormlyFieldFile, wrappers: ['form-field'] },
    { name: 'datepicker', component: FormlyFieldDatePicker, wrappers: ['form-field'] },
    { name: 'searchable', component: FormlyFieldSearchableSelectApi, wrappers: ['form-field'] },
    { name: 'ngsearchable', component: FormlyFieldNgSelect, wrappers: ['form-field'] },
  ],
  validationMessages: [
    { name: 'required', message: 'This field is required' },
  ]
};

@NgModule({
  imports: [
    FormlyModule,
    FormsModule,
    ReactiveFormsModule,
    CommonModule,
    DatepickerModule,
    NgSelectModule
  ],
  declarations: [
    FileValueAccessor,
    FormlyFieldFile,
    FormlyFieldDatePicker,
    FormlyFieldSearchableSelectApi,
    FormlyFieldNgSelect
  ]
})
export class FormlyTypesModule { }
