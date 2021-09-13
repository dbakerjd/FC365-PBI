import { Component, Input } from '@angular/core';

@Component({
  selector: 'app-progress-spinner',
  templateUrl: './progress-spinner.component.html',
})
export class ProgressSpinnerComponent {

  @Input() size: 'normal' | 'small' = 'normal';
  
 }
