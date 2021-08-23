import { Directive } from '@angular/core';
import { NG_VALUE_ACCESSOR, ControlValueAccessor } from '@angular/forms';

@Directive({
  selector: 'input[type=file]',
  host: {
    '(change)': 'onChange($event.target.files)',
    '(blur)': 'onTouched()',
  },
  providers: [
    { provide: NG_VALUE_ACCESSOR, useExisting: FileValueAccessor, multi: true },
  ],
})
// https://github.com/angular/angular/issues/7341
export class FileValueAccessor implements ControlValueAccessor {
  value: any;
  onChange = (_: any) => { };
  onTouched = () => { };

  writeValue(value: any) { }
  registerOnChange(fn: any) { this.onChange = fn; console.log('chan', fn) }
  registerOnTouched(fn: any) { this.onTouched = fn; console.log('touc', fn)}
}