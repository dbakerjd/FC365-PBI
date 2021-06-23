import { Injectable } from '@angular/core';
import * as microsoftTeams from "@microsoft/teams-js";

@Injectable({
  providedIn: 'root'
})
export class TeamsService {
  public user: any;
  public token: any;
  public context: any;

  constructor() { 

    microsoftTeams.initialize();
    
    microsoftTeams.getContext((context) => {
      this.context = context;
      console.log(context);
    });
    
    microsoftTeams.authentication.getAuthToken({ 
      successCallback: (result) => {
        this.token = result;
        console.log(result);
      },
      failureCallback: (error) => {
        this.token = error;
        console.log(error);
      }
    });
  }

}
