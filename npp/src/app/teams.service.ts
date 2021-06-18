import { Injectable } from '@angular/core';
import { TeamsUserCredential } from "@microsoft/teamsfx";

@Injectable({
  providedIn: 'root'
})
export class TeamsService {
  public user: any;

  constructor() { 

    let cred = new TeamsUserCredential();
    this.user = cred.getUserInfo();
    
  }

}
