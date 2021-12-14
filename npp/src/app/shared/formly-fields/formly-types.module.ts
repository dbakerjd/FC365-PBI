import { CommonModule } from "@angular/common";
import { NgModule } from "@angular/core";
import { AbstractControl, FormsModule, ReactiveFormsModule } from "@angular/forms";
import { NgSelectModule } from "@ng-select/ng-select";
import { FormlyModule } from "@ngx-formly/core";
import { DatepickerModule } from "ng2-datepicker";
import { FormlyFieldDatePicker } from "./date-picker";
import { FormlyFieldFile } from "./file-input";
import { FileValueAccessor } from "./file-value-accessor";
import { FormlyFieldNgSelect } from "./ng-select-input";
import { FormlyFieldSearchableSelectApi } from "./sharepoint-searchable-select";

export function afterDateValidator(control: AbstractControl) {
  const { ProjectStartDate, ProjectEndDate } = control.value.Opportunity;

  // avoid displaying the message error when values are empty
  if (!ProjectStartDate || !ProjectEndDate) {
    return null;
  }

  if (ProjectEndDate >= ProjectStartDate) {
    return null;
  }

  return { fieldMatch: { message: 'The end date cannot be earlier than the start date' } };
}

export const FORMLY_CONFIG = {
  types: [
    { name: 'file-input', component: FormlyFieldFile, wrappers: ['form-field'] },
    { name: 'datepicker', component: FormlyFieldDatePicker, wrappers: ['form-field'] },
    { name: 'searchable', component: FormlyFieldSearchableSelectApi, wrappers: ['form-field'] },
    { name: 'ngsearchable', component: FormlyFieldNgSelect, wrappers: ['form-field'] },
  ],
  validators: [
    { name: 'afterDate', validation: afterDateValidator },
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
