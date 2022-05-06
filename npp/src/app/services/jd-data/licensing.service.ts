import { HttpBackend, HttpClient, HttpErrorResponse, HttpHeaders } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Router } from '@angular/router';
import { LicenseContext } from '@shared/models/app-config';
import { Md5 } from 'ts-md5';
import { ErrorService } from '../app/error.service';

export interface JDLicense {
  Tier: string;
  Expiration: Date;
  SharePointUri: string;
  HasPowerBi: boolean;
  AppId: string;
  TenantId: string;
  TotalSeats: number;
  AssignedSeats: number;
  AvailableSeats: number;
  ContactName: string;
  ContactEmail: string;
}

interface JDLicenseContext {
  appId?: string;
  installationAddress?: string;
}

interface SeatsResponse {
  TotalSeats: number;
  AssignedSeats: number;
  AvailableSeats: number;
  UserGroupsCount: number;
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
  private httpClient: HttpClient;

  constructor(
    private error: ErrorService, 
    private handler: HttpBackend, 
    private router: Router
  ) { 
    this.httpClient = new HttpClient(handler);
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
      let dataRequest;
      if (context.entityId) {
        dataRequest = {
          "appId" : context.entityId,
        }
      } else {
        dataRequest = {
          "InstallationAddress" : context.host
        }
      }
      return await this.httpClient.post(this.licensingApiUrl + '/license', dataRequest, { 
        headers: headers
      }).toPromise() as JDLicense;
  }

  async setJDLicense(context: LicenseContext) {
    this.license = await this.askLicensingApi(context);
    localStorage.setItem("JDLicense", JSON.stringify(this.license));
    if (context.entityId) {
      this.licenseContext = { appId: context.entityId }
    } else {
      this.licenseContext = { installationAddress: context.host }
    }
  }

  getJDLicense(): JDLicense | null {
    return this.license;
  }

  isValidJDLicense() {
    if (!this.license) return false;
    return (new Date(this.license.Expiration)).getTime() >= new Date().getTime();
  }

  async validateLicense(context: any) {
    try {
      await this.setJDLicense(context);
      if(!this.isValidJDLicense()) {
        // this.error.toastr.error("License not valid: "+JSON.stringify(this.license));
        this.router.navigate(['splash/expired']);
      }
      return true;
    } catch(e) {
      this.router.navigate(['splash/non-license']);
      this.error.handleError(e as Error);
      return false;
    }
    
  }

  async addSeat(email: string): Promise<SeatsResponse | null> {
    if (email.trim() == '') return null;
    let headers = new HttpHeaders({
      'x-functions-key': 'Gyzm5Htg4Er8UJTzlfAI2a0Vsg3bVubLTRak7xVIeMLTO9HzgW4e1Q==',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST',
    });
    try {
      if (this.licenseContext) {
        return await this.httpClient.post(this.licensingApiUrl + '/seats', {
          applicationIdentity: this.licenseContext,
          userEmail: Md5.hashStr(email)
        }, {
          headers: headers
        }).toPromise() as SeatsResponse;
      }
      return null;
    } catch(e: any) {
      if (e.status === 422) {
        throw e;
      }
      return null;
    }
  }

  async removeSeat(email: string): Promise<SeatsResponse | null> {
    if (email.trim() == '') return null;
    let headers = new HttpHeaders({
      'x-functions-key': 'Gyzm5Htg4Er8UJTzlfAI2a0Vsg3bVubLTRak7xVIeMLTO9HzgW4e1Q==',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'DELETE',
    });
    try {
      if (this.licenseContext) {
        return await this.httpClient.request(
          'delete',
          this.licensingApiUrl + '/seats',
          {
            headers: headers,
            body: {
              applicationIdentity: this.licenseContext,
              userEmail: Md5.hashStr(email)
            },
          }).toPromise() as SeatsResponse;
        }
        return null;
    } catch(e: any) {
      if (e.status === 422) {
        throw e;
      }
      return null;
    }
  }

  async getSeats(email: string): Promise<SeatsResponse | null> {
    let headers = new HttpHeaders({
      'x-functions-key': 'Gyzm5Htg4Er8UJTzlfAI2a0Vsg3bVubLTRak7xVIeMLTO9HzgW4e1Q==',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST',
    });
    try {
      if (this.licenseContext) {
        return await this.httpClient.post(this.licensingApiUrl + '/userseats', {
          applicationIdentity: this.licenseContext,
          userEmail: Md5.hashStr(email)
        }, {
          headers: headers
        }).toPromise() as SeatsResponse;
      }
      return null;
    } catch(e: any) {
      if (e.status === 422) {
        throw e;
      }
      return null;
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

  getContactInfo(): { contactName: string; contactEmail: string; } | null {
    if (this.license) {
      return {
        contactName: this.license.ContactName,
        contactEmail: this.license.ContactEmail
      };
    }
    return null;
  }
  
}
