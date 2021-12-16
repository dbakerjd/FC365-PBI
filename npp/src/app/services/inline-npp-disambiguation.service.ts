import { Injectable } from '@angular/core';
import { environment } from 'src/environments/environment';
import { Brand, Opportunity, SharepointService } from './sharepoint.service';

@Injectable({
  providedIn: 'root'
})
export class InlineNppDisambiguationService {
  
  isInline: boolean = false;
  
  constructor(private readonly sharepoint: SharepointService) { 
    this.isInline = environment.isInlineApp;
  }

  getEntity(id: number) {
    if(this.isInline) {
      return this.sharepoint.getBrand(id);
    } else {
      return this.sharepoint.getOpportunity(id);
    }
  }

  getOwnerId(entity: Brand | Opportunity) {
    if(this.isInline) {
      return (entity as Brand).BrandOwnerId;
    } else {
      return (entity as Opportunity).OpportunityOwnerId;
    }
  }

  getOwner(entity: Brand | Opportunity) {
    if(this.isInline) {
      return (entity as Brand).BrandOwner;
    } else {
      return (entity as Opportunity).OpportunityOwner;
    }
  }

  getForecastCycles(entity: Brand | Opportunity) {
    if(this.isInline) {
      return this.sharepoint.getBrandForecastCycles(entity as Brand);
    } else {
      return this.sharepoint.getOpportunityForecastCycles(entity as Opportunity);
    }
  }

  
}
