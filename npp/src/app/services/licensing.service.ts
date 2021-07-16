import { Injectable } from '@angular/core';

@Injectable({
  providedIn: 'root'
})
export class LicensingService {
  siteUrl: string = 'https://betasoftwaresl.sharepoint.com/_api/';
  constructor() { }
}
