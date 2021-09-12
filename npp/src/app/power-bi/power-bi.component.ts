import { Component, OnInit } from '@angular/core';
import { LicensingService } from '../services/licensing.service';

@Component({
  selector: 'app-power-bi',
  templateUrl: './power-bi.component.html',
  styleUrls: ['./power-bi.component.scss']
})
export class PowerBiComponent implements OnInit {

  constructor(public licensing: LicensingService) { }

  ngOnInit(): void {
  }

}
