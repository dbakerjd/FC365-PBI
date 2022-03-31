import { Injectable } from '@angular/core';
import { Opportunity } from '@shared/models/entity';
import { SharepointService } from './sharepoint.service';
import { InlineNppDisambiguationService } from './inline-npp-disambiguation.service';
import { ENTITIES_LIST_NAME } from '@shared/sharepoint/list-names';

@Injectable({
  providedIn: 'root'
})
export class EntitiesService {

  constructor(private readonly sharepoint: SharepointService, private readonly appService: InlineNppDisambiguationService) { }

  async getAll(expand = true, onlyActive = false): Promise<Opportunity[]> {
    let filter = undefined;
    if (expand) {
      //TODO check why OpportunityType/isInternal is failing
      filter = "$select=*,ClinicalTrialPhase/Title,OpportunityType/Title,Indication/TherapyArea,Indication/Title,EntityOwner/FirstName,EntityOwner/LastName,EntityOwner/ID,EntityOwner/EMail&$expand=OpportunityType,Indication,EntityOwner,ClinicalTrialPhase";
    }
    if (onlyActive) {
      if (!filter) filter = "$filter=AppTypeId eq '"+this.appService.getAppType()?.ID+"' and OpportunityStatus eq 'Active'";
      else filter += "&$filter=AppTypeId eq '"+this.appService.getAppType()?.ID+"' and OpportunityStatus eq 'Active'";
    } else {
      if (!filter) filter = "$filter=AppTypeId eq '"+this.appService.getAppType()?.ID+"'";
      else filter += "&$filter=AppTypeId eq '"+this.appService.getAppType()?.ID+"'";
    }

    return await this.sharepoint.getAllItems(ENTITIES_LIST_NAME, filter);
  }
}
