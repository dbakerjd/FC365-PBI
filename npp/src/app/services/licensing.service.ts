import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Injectable } from '@angular/core';

export interface JDLicense {
  Tier: string;
  Expiration: Date;
  SharePointUri: string;
}

@Injectable({
  providedIn: 'root'
})
export class LicensingService {
  siteUrl: string = 'https://betasoftwaresl.sharepoint.com/sites/JDNPPApp/';
  licensingApiUrl: string = ' https://jdlicensingfunctions.azurewebsites.net/api/license?code=0R6EUPw28eUEVmBU9gNfi1yEwEpX28kOUWXZtEIjxavv5qV6VacwDw==';

  private license: JDLicense | null = null;

  constructor(private http: HttpClient) { }

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
    return /*this.license.Tier == "silver" && */this.license.Expiration.getTime() >= new Date().getTime();
  }

  getSharepointUri() {
    return this.siteUrl; // temporal
    return this.license?.SharePointUri;
  }

  getSharepointApiUri() {
    return this.getSharepointUri() + '/_api/web/';
  }
}
