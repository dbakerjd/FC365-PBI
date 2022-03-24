import { HttpClient } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { ErrorService } from './error.service';
import { TeamsService } from './teams.service';

export interface MSGraphQueryParams {
  count?: boolean;
  expand?: string;
  filter?: string;
  orderBy?: string;
  search?: string;
  select?: string;
  top?: number;
}

export interface MSGraphAPIResult {
  'odata.context': string;
  'odata.nextLink'?: string;
  value?: any;
}

export interface AzureGroup {
  id: string;
  createdDateTime: Date;
  displayName: string;
  renewedDateTime: Date;
}

export interface MSGraphUser {
  id: string;
  displayName: string;
  givenName: string;
  jobTitle: string;
  mail: string;
  surName: string;
  userPrincipalName: string;
}

@Injectable({
  providedIn: 'root'
})
export class GraphService {

  private baseGraphUrl = 'https://graph.microsoft.com/v1.0/';

  constructor(
    private readonly http: HttpClient,
    private error: ErrorService,
    private readonly teams: TeamsService
  ) { }

  async getMSGraphToken() {
    let token: string = "";

    try {
      let domain = "graph.microsoft.com";
      let scopes = this.teams.getResourceByDomain(domain);

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

  /** Executes a GET Request using a MS Graph Token and return the result of the query */
  private async getRequest(endpoint: string, params: MSGraphQueryParams | null = null): Promise<any> {
    try {
      const graphToken = await this.getMSGraphToken();
      endpoint = this.baseGraphUrl + endpoint + this.generateParamsQueryString(params);
      const result = await this.http.get(
        endpoint,
        {
          headers: {
            token: graphToken
          }
        }
      ).toPromise() as MSGraphAPIResult;
      if (result.value) return result.value;
      else if (result) return result;
      return null;
    } catch (e: any) {
      this.error.handleError(e);
      return null;
    }
  }

  /** Executes a POST Request using a MS Graph Token and return the result of the query */
  private async postRequest(endpoint: string, body: any): Promise<any> {
    try {
      const graphToken = await this.getMSGraphToken();
      endpoint = this.baseGraphUrl + endpoint;
      const result = await this.http.post(
        endpoint,
        body,
        {
          headers: {
            token: graphToken
          }
        }
      ).toPromise() as MSGraphAPIResult;
      return true;
    } catch (e: any) {
      this.error.handleError(e);
      return false;
    }
  }

  /** Executes a DELETE Request using a MS Graph Token */
  private async deleteRequest(endpoint: string): Promise<boolean> {
    try {
      const graphToken = await this.getMSGraphToken();
      endpoint = this.baseGraphUrl + endpoint;
      const result = await this.http.delete(
        endpoint,
        {
          headers: {
            token: graphToken
          }
        }
      ).toPromise() as MSGraphAPIResult;
      return true;
    } catch (e: any) {
      this.error.handleError(e);
      return false;
    }
  }

  /** Generates the params string for filtering requests to the MS Graph API */
  private generateParamsQueryString(params: MSGraphQueryParams | null = null): string {
    if (!params) return '';
    let queryArray: string[] = [];
    if (params.filter) {
      queryArray.push('$filter='+params.filter);
    }
    return '?' + queryArray.join('&');
  }

  /** List all the Azure Groups */
  async getAllGroups(): Promise<AzureGroup[]> {
    return await this.getRequest('groups');
  }

  /** Returns the Azure Group with the id requested */
  async getAzureGroupById(id: string): Promise<AzureGroup | null> {
    const group: AzureGroup = await this.getRequest('groups/' + id);
    return group ? group : null;
  }

  /** Returns the Azure Group named as "name" */
  async getAzureGroupByName(name: string): Promise<AzureGroup | null> {
    const result = await this.getRequest('groups', { filter: `displayName eq '${name}'` });
    return result[0] ? result[0] as AzureGroup : null;
  }

  /** Adds the MS User as a member of the Azure Group */
  async addUserToAzureGroup(userId: string, groupId: string): Promise<boolean> {
    return await this.postRequest(`groups/${groupId}/members/$ref`, {
      '@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/'+ userId
    });
  }

  /** Removes the MS User as a member of the Azure Group */
  async removeUserToAzureGroup(userId: string, groupId: string): Promise<boolean> {
    return await this.deleteRequest(`groups/${groupId}/members/${userId}/$ref`);
  }

  /** Current user info */
  async getCurrentMSGraphUser(): Promise<MSGraphUser> {
    return await this.getRequest('me');
  }

  /** Find a Microsoft Graph User by Principal Name (email) */
  async getUserByPrincipalName(name: string): Promise<MSGraphUser | null> {
    return await this.getRequest(`users/${name}`);
  }

  /** Adds the user to the Azure Group controling Power BI RLS Access */
  async addUserToPowerBI_RLSGroup(email: string, groupName: string): Promise<boolean> {
    const group = await this.getAzureGroupByName(groupName);
    const user = await this.getUserByPrincipalName(email);
    if (group && user) return this.addUserToAzureGroup(user.id, group.id);
    return false;
  }

  /** Removes the user of the Azure Group controling Power BI RLS Access */
  async removeUserToPowerBI_RLSGroup(email: string, groupName: string): Promise<boolean> {
    const group = await this.getAzureGroupByName(groupName);
    const user = await this.getUserByPrincipalName(email);
    if (group && user) return this.removeUserToAzureGroup(user.id, group.id);
    return false;
  }

  /** User profile pic from Microsoft Graph */
  async getProfilePic(email: string): Promise<Blob | null> {
    try {
      const graphToken = await this.getMSGraphToken();
      const endpoint = this.baseGraphUrl + 'users/' + email + '/photo/$value';
      const result = await this.http.get(
        endpoint,
        {
          headers: {
            token: graphToken
          },
          responseType: 'arraybuffer',
        }
      ).toPromise();
      if (result) {
        return new Blob([result]);
      }
      return null;
    } catch (e: any) {
      this.error.handleError(e);
      return null;
    }
  }
}
