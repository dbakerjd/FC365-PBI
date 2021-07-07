import { Injectable } from '@angular/core';
import * as microsoftTeams from "@microsoft/teams-js";
import { ErrorService } from './error.service';

@Injectable({
  providedIn: 'root'
})
export class TeamsService {
  public user: any;
  public token: any;
  public context: any;

  constructor(private errorService: ErrorService) { 

    microsoftTeams.initialize();
    
    microsoftTeams.getContext((context) => {
      this.context = context;
      console.log(context);
    });
    
    microsoftTeams.authentication.authenticate({ 
      url: window.location.href + '/auth-start.html',
      width: 600,
      height: 535,
      successCallback: (result) => {
        this.token = result;
        console.log(result);
      },
      failureCallback: (error) => {
        this.token = false;
        this.errorService.handleError(new Error(error));
      }
    });
  }

}
