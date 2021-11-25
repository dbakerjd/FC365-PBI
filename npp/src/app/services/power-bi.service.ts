import { Injectable } from '@angular/core';
import { HttpClient, HttpHeaders, HttpRequest } from '@angular/common/http';
import { LicensingService } from './licensing.service';
import { ErrorService } from './error.service';
import { TeamsService } from './teams.service';
import { PBIReport } from './sharepoint.service';

export interface PageDetails {
  ReportSection: string;
  DisplayName: string;
}

export interface PBIResult{
  'odata.metadata': string;
  value: any;
}

export interface PBIObject {
  id: string;
  name: string;
  webUrl: string;
  embedUrl: string;
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
  async getObjects(groupId: string, objectType: string): Promise<PBIObject[]> {

    const url: string = `https://api.powerbi.com/v1.0/myorg/groups/${groupId}/${objectType}`

    let objects = await this.http.get(url).toPromise() as PBIResult;

    if (objects.value && objects.value.length >0){
      return objects.value
    }
    return[];

  }

  async getReportId(groupId: string, objectType: string, reportName: string): Promise<PBIObject> {
    
    let objects:PBIObject[] = await this.getObjects(groupId,objectType);
    
    console.log(objects);
    
    let returnObject!:PBIObject;
    
    objects.forEach(async (object)=>{
      
      if(object.name == reportName){
        var objectDetails = {id:object.id, name:object.name, webUrl: object.webUrl, embedUrl: object.embedUrl}
        
        returnObject = objectDetails;  
      }
      
    })
    console.log(returnObject);
    return returnObject;


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
        console.log(token);
        return token;

      }
    } catch (e: any) {
      this.error.handleError(e);
      return token;
    }
    return token;

  }
}
