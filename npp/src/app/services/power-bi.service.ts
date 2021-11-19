import { Injectable } from '@angular/core';
import { HttpClient, HttpHeaders, HttpRequest } from '@angular/common/http';
import { LicensingService } from './licensing.service';
import { ErrorService } from './error.service';
import { TeamsService } from './teams.service';

export interface PageDetails {
  ReportSection: string;
  DisplayName: string;
}

@Injectable({
  providedIn: 'root'
})

export class PowerBiService {

  constructor(private http: HttpClient, private error: ErrorService, private licensing: LicensingService, private teams: TeamsService) { }

  async refreshReport() {
    try {
      let url = this.licensing.license?.PowerBi?.Refresh;
      if (url) {
        let res = await this.http.post(url, {}).toPromise();

        this.teams.hackyConsole += "******* POWER BI REFRESH ********      " + JSON.stringify(res) + "       ************************";
        return true;
      } else {
        this.error.handleError(new Error("Licensing information missing, try again in a few seconds."));
        return false;
      }
    } catch (e: any) {
      this.error.handleError(e);
      return false;
    }

  }

  async getPBIToken() {
    let token: string = "";
    
    try {
      let domain = "api.powerbi.com";
      console.log("trying to obtain token for domain: " + domain);
      let scopes = this.teams.getResourceByDomain(domain);
      console.log(scopes);

      if (scopes) {
        let request = {
          scopes
        }

        let tokenResponse = await this.teams.msalInstance.acquireTokenSilent(request);
        token = tokenResponse.accessToken;
        return token;

      }
    } catch (e: any) {
      this.error.handleError(e);
      return token;
    }
    return token;

  }
}
