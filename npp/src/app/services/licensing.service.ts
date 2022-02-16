import { HttpClient, HttpErrorResponse, HttpHeaders } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Router } from '@angular/router';
import { ErrorService } from './error.service';

export interface JDLicense {
  Tier: string;
  Expiration: Date;
  SharePointUri: string;
  HasPowerBi: boolean;
  TenantId: string;
  TotalSeats: number;
  AssignedSeats: number;
  AvailableSeats: number;
}

interface JDLicenseContext {
  appId: number;
  teamSiteDomain: string;
}

@Injectable({
  providedIn: 'root'
})
export class LicensingService {
  //siteUrl: string = 'https://betasoftwaresl.sharepoint.com/sites/JDNPPApp/';
  siteUrl: string = 'https://janddconsulting.sharepoint.com/sites/NPPBetaV1/';
  //licensingApiUrl: string = 'https://jdlicensingfunctions.azurewebsites.net/api/license?code=tFs/KoE40qeTvQlsYUTA6GmgF88G3QF9RXxX51kasNV2Z8nzr2Y/hA==';
  licensingApiUrl: string = 'https://jdlicensingfunctions.azurewebsites.net/api';
  
  public license: JDLicense | null = null;
  private licenseContext: JDLicenseContext | null = null;

  constructor(
    private error: ErrorService, 
    private http: HttpClient, 
    private router: Router
  ) { 
    let license = localStorage.getItem("JDLicense");
    if(license) {
      this.license = JSON.parse(license);
    }
  }

  async askLicensingApi(context: any): Promise<JDLicense> {

      let headers = new HttpHeaders({
        'x-functions-key': 'Gyzm5Htg4Er8UJTzlfAI2a0Vsg3bVubLTRak7xVIeMLTO9HzgW4e1Q==',
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'POST',
      });
      return await this.http.post(this.licensingApiUrl + '/license', {
        "appId" : context.entityId,
        "teamSiteDomain" : context.teamSiteDomain
      }, { 
        headers: headers
      }).toPromise() as JDLicense;
  }

  async setJDLicense(context: any) {
    this.license = await this.askLicensingApi(context);
    localStorage.setItem("JDLicense", JSON.stringify(this.license));
    this.licenseContext = {
      appId : context.entityId,
      teamSiteDomain : context.teamSiteDomain
    };
    console.log('SEATS: License', this.license);
  }

  isValidJDLicense() {
    if (!this.license) return false;
    return (new Date(this.license.Expiration)).getTime() >= new Date().getTime();
  }

  async validateLicense(context: any) {
    try {
      await this.setJDLicense(context);
      if(!this.isValidJDLicense()) {
        this.error.toastr.error("License not valid: "+JSON.stringify(this.license));
        this.router.navigate(['splash/expired']);
      }
      return true;
    } catch(e) {
      this.router.navigate(['splash/expired']);
      this.error.handleError(e as Error);
      return false;
    }
    
  }

  async addSeat(email: string) {
    let headers = new HttpHeaders({
      'x-functions-key': 'Gyzm5Htg4Er8UJTzlfAI2a0Vsg3bVubLTRak7xVIeMLTO9HzgW4e1Q==',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST',
    });
    try {
      return await this.http.post(this.licensingApiUrl + '/seats', {
        applicationIdentity: this.licenseContext,
        userEmail: email
      }, {
        headers: headers
      }).toPromise();
    } catch(e: any) {
      if (e.status === 422) {
        throw e;
      }
      return false;
    }
  }

  async removeSeat(email: string) {
    let headers = new HttpHeaders({
      'x-functions-key': 'Gyzm5Htg4Er8UJTzlfAI2a0Vsg3bVubLTRak7xVIeMLTO9HzgW4e1Q==',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'DELETE',
    });
    try {
      return await this.http.request(
        'delete',
        this.licensingApiUrl + '/seats',
        {
          headers: headers,
          body: {
            applicationIdentity: this.licenseContext,
            userEmail: email
          },
        }).toPromise();
    } catch(e: any) {
      if (e.status === 422) {
        throw e;
      }
      return false;
    }
  }

  async getSeats(email: string) {
    let headers = new HttpHeaders({
      'x-functions-key': 'Gyzm5Htg4Er8UJTzlfAI2a0Vsg3bVubLTRak7xVIeMLTO9HzgW4e1Q==',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET',
    });
    try {
      // return await this.http.get(this.licensingApiUrl + '/seats',{
      //   applicationIdentity: this.licenseContext,
      //   userEmail: email
      // }).toPromise();
      if (this.licenseContext)
      // return await this.http.get(this.licensingApiUrl + '/seats', {
      //   headers: headers,
      //   params: {
      //     userEmail: email,
      //     teamSiteDomain: this.licenseContext?.teamSiteDomain,
      //     appId: this.licenseContext?.appId
      //   }
      // }).toPromise();
      return await this.http.request(
        'get',
        this.licensingApiUrl + '/seats', 
        { 
          headers: headers,
          body: {
            applicationIdentity: this.licenseContext,
            userEmail: email
          },
        }).toPromise();

      return 0;
    } catch(e: any) {
      if (e.status === 422) {
        throw e;
      }
      return false;
    }
  }

  getSharepointUri() {
    //return this.siteUrl; // temporal
    return this.license?.SharePointUri;
  }

  getSharepointApiUri() {
    return this.getSharepointUri() + '/_api/web/';
  }

  getSharepointDomain() {
    return this.getSharepointUri()?.split('/')[2];
  }
  
}
