import { Injectable } from '@angular/core';
import { HttpClient, HttpHeaders, HttpRequest, HttpResponse } from '@angular/common/http';
import { ErrorService } from './app/error.service';
import { environment } from 'src/environments/environment';
import { PBIRefreshComponent, PBIReport } from '../shared/models/pbi';
import { AppDataService } from './app/app-data.service';
import { TeamsService } from './microsoft-data/teams.service';

export interface PageDetails {
  ReportSection: string;
  DisplayName: string;
}

export interface PBIResult {
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

  report: PBIReport| undefined = undefined;
  reportComponents: PBIRefreshComponent[] = [];

  constructor(
    private http: HttpClient, 
    private error: ErrorService, 
    private teams: TeamsService, 
    private readonly appData: AppDataService) { }

  async refreshReport(reportName: string): Promise<number> {
    try {

      this.report = await this.appData.getReportByName(encodeURIComponent(reportName));
      this.reportComponents = await this.appData.getComponents(this.report);

      const token = await this.getPBIToken();
      const userObjectId = this.teams.context.userObjectId;

      const body = {
        reportType: reportName,
        token: token,
        userObjectId: userObjectId,
        entityId: this.teams.context.entityId,
        teamSiteDomain: this.teams.context.teamSiteDomain,
        reportComponents: this.reportComponents
      }

      const url = environment.functionAppUrl;
      
      return new Promise((resolve) => {
        this.http.post(url, body, { observe: 'response' }).subscribe(response => {
          resolve(response.status);
        }, error => {
          resolve(error.status);
        })
      })

      //this.teams.hackyConsole += "******* POWER BI REFRESH ********      " + JSON.stringify(res) + "       ************************";


    } catch (e: any) {

      this.error.handleError(e);

      return 500;
    }
  }

  async getObjects(groupId: string, objectType: string): Promise<PBIObject[]> {

    const url: string = `https://api.powerbi.com/v1.0/myorg/groups/${groupId}/${objectType}`

    let objects = await this.http.get(url).toPromise() as PBIResult;

    if (objects.value && objects.value.length > 0) {
      return objects.value
    }
    return [];

  }

  async getReportId(groupId: string, objectType: string, reportName: string): Promise<PBIObject> {

    let objects: PBIObject[] = await this.getObjects(groupId, objectType);

    console.log(objects);

    let returnObject!: PBIObject;

    objects.forEach(async (object) => {

      if (object.name == reportName) {
        var objectDetails = { id: object.id, name: object.name, webUrl: object.webUrl, embedUrl: object.embedUrl }

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
