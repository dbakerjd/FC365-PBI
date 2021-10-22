import { Component, OnInit } from '@angular/core';
import { LicensingService } from '../services/licensing.service';
import * as microsoftTeams from "@microsoft/teams-js";
import { ErrorService } from '../services/error.service';

//const PBIReportFromLicencingAPI = "https://app.powerbi.com/groups/b76b03e1-cdd6-4233-b682-a1d81f25ba04/reports/531cfc4a-9c47-4163-a152-37002ad84d6d/ReportSectionf6c5d750370200b00708"

@Component({
  selector: 'app-power-bi',
  templateUrl: './power-bi.component.html',
  styleUrls: ['./power-bi.component.scss']
})
export class PowerBiComponent implements OnInit {
  constructor(public licensing: LicensingService,
    private error: ErrorService,) {

  }

  ngOnInit(): void {
    const reportFromLicencingAPI = this.licensing.license?.PowerBi?.Report
    const entity = encodeURI("1c4340de-2a85-40e5-8eb0-4f295368978b/Home");
    const deepLink = `https://teams.microsoft.com/l/entity/${entity}?context={"subEntityId":"${reportFromLicencingAPI}"?action=OpenReport&pbi_source=MSTeams"}`
    try {
      microsoftTeams.executeDeepLink(deepLink);
    } catch (e: any) {
      this.error.handleError(e);
    }
  }
}
