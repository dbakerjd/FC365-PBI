import { Component } from '@angular/core';
import { FieldType } from '@ngx-formly/core';
import { DatepickerOptions } from 'ng2-datepicker';

@Component({
  selector: 'formly-field-date-picker',
  template: `
    <ngx-datepicker [(ngModel)]="formData" [options]="dateOptions"></ngx-datepicker>  
  `,
})
export class FormlyFieldDatePicker extends FieldType {
  _formData: string = '';
  oldValue: string = '';
  dateOptions: DatepickerOptions = {
    format: 'Y-M-d'
  };

  ngOnInit() {
    this._formData = this.formControl.value;
  }

  set formData(value: string) {
    this._formData = value;
    this.formControl.setValue(value);
  }

  get formData() {
    return this._formData;
  }

}