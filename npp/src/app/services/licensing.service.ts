import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Router } from '@angular/router';
import { MsalService } from '@azure/msal-angular';
import { TeamsService } from './teams.service';

export interface JDLicense {
  Tier: string;
  Expiration: Date;
  SharePointUri: string;
  PowerBi?:  any;
}

@Injectable({
  providedIn: 'root'
})
export class LicensingService {
  siteUrl: string = 'https://betasoftwaresl.sharepoint.com/sites/JDNPPApp/';
  licensingApiUrl: string = ' https://jdlicensingfunctions.azurewebsites.net/api/license?code=0R6EUPw28eUEVmBU9gNfi1yEwEpX28kOUWXZtEIjxavv5qV6VacwDw==';

  public license: JDLicense | null = null;

  constructor(private http: HttpClient, private teams: TeamsService, private router: Router, private authService: MsalService) { }

  async askLicensingApi(token: string): Promise<JDLicense> {

     /** OK */
      let headers = new HttpHeaders({
        'token': token,
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET',
      });
      return await this.http.get(this.licensingApiUrl, { 
        headers: headers
      }).toPromise() as JDLicense;

      return {
        "Tier": "silver",
        "Expiration": new Date("2021-12-29T00:00:00"),
        "SharePointUri": 'https://betasoftwaresl.sharepoint.com/sites/JDNPPApp/_api/web/'
      };
  }

  async setJDLicense(token: string) {
    this.license = await this.askLicensingApi(token);
    console.log("license", this.license);
  }

  isValidJDLicense() {
    if (!this.license) return false;
    /*this.license.Tier == "silver" && */
    return (new Date(this.license.Expiration)).getTime() >= new Date().getTime();
  }

  async validateLicense() {
    let activeAccount = this.authService.instance.getActiveAccount();

    if(activeAccount) {
      this.teams.getStorageToken();

      if(this.teams.token) {
        if(!this.license) {
          await this.setJDLicense(this.teams.token);
        }
        if(!this.isValidJDLicense()) {
          this.router.navigate(['expired-license']);
        }
      }
    }
    
    return true;
  }

  getSharepointUri() {
    return this.siteUrl; // temporal
    return this.license?.SharePointUri;
  }

  getSharepointApiUri() {
    return this.getSharepointUri() + '/_api/web/';
  }

  getSharepointDomain() {
    return this.getSharepointUri()?.split('/')[2];
  }
}
