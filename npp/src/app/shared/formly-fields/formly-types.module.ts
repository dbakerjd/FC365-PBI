import { CommonModule } from "@angular/common";
import { NgModule } from "@angular/core";
import { FormsModule, ReactiveFormsModule } from "@angular/forms";
import { FormlyModule } from "@ngx-formly/core";
import { DatepickerModule } from "ng2-datepicker";
import { FormlyFieldDatePicker } from "./date-picker";
import { FormlyFieldFile } from "./file-input";

export const FORMLY_CONFIG = {
  types: [
    { name: 'file-input', component: FormlyFieldFile, wrappers: ['form-field']  },
    { name: 'datepicker', component: FormlyFieldDatePicker, wrappers: ['form-field']  },
  ],
};

@NgModule({
  imports: [
    FormlyModule,
    FormsModule,
    ReactiveFormsModule,
    CommonModule,
    DatepickerModule
  ],
  declarations: [
    FormlyFieldFile,
    FormlyFieldDatePicker
  ]
})
export class FormlyTypesModule { }
