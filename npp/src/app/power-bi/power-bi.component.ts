import { Component, OnInit } from '@angular/core';
import { LicensingService } from '../services/licensing.service';
import * as microsoftTeams from "@microsoft/teams-js";
import { ErrorService } from '../services/error.service';
import { ToastrService } from 'ngx-toastr';

@Component({
  selector: 'app-power-bi',
  templateUrl: './power-bi.component.html',
  styleUrls: ['./power-bi.component.scss']
})
export class PowerBiComponent implements OnInit {
  constructor(
    public licensing: LicensingService,
    private error: ErrorService,
    private toastr: ToastrService,) {

  }

  ngOnInit(): void {
    
    let reportFromLicencingAPI:string = this.licensing.license?.PowerBi?.Report!
    var encodedReportFromLicencingAPI = encodeURIComponent(reportFromLicencingAPI);
    var opportunity = 'Acquisition of Nucala for COPD';
    var filter = encodeURIComponent(`?filter=Opportunities/Opportunity eq ${opportunity}`)
    var entity = encodeURI('1c4340de-2a85-40e5-8eb0-4f295368978b/Home');
    var deepLink = `https://teams.microsoft.com/l/entity/${entity}?context={"subEntityId":"${encodedReportFromLicencingAPI}%3Faction%3DOpenReport%26pbi_source%3DMSTeams"}`;

    try {
      this.toastr.success("Done");
      microsoftTeams.executeDeepLink(deepLink);
    } catch (e: any) {
      this.error.handleError(e);
    }
  }
}
