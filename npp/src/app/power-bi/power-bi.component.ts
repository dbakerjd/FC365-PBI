import { Component, OnInit } from '@angular/core';
import { LicensingService } from '../services/licensing.service';
import * as microsoftTeams from "@microsoft/teams-js";
import { ErrorService } from '../services/error.service';
import { ToastrService } from 'ngx-toastr';
import { connectableObservableDescriptor } from 'rxjs/internal/observable/ConnectableObservable';

//const PBIReportFromLicencingAPI = "https://app.powerbi.com/groups/b76b03e1-cdd6-4233-b682-a1d81f25ba04/reports/531cfc4a-9c47-4163-a152-37002ad84d6d/ReportSectionf6c5d750370200b00708"
const BigPBIReport ='https://teams.microsoft.com/l/entity/1c4340de-2a85-40e5-8eb0-4f295368978b/Home?context={"subEntityId":"https%3A%2F%2Fapp.powerbi.com%2Fgroups%2Fb76b03e1-cdd6-4233-b682-a1d81f25ba04%2Freports%2F531cfc4a-9c47-4163-a152-37002ad84d6d%2FReportSectionf6c5d750370200b00708%3Faction%3DOpenReport%26pbi_source%3DMSTeams"}';
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
    const reportFromLicencingAPI = this.licensing.license?.PowerBi?.Report
    const entity = encodeURI("1c4340de-2a85-40e5-8eb0-4f295368978b/Home");
    const deepLink = `https://teams.microsoft.com/l/entity/${entity}?context={"subEntityId":"${reportFromLicencingAPI}"?action=OpenReport&pbi_source=MSTeams"}`
    console.log(deepLink);
    console.log(BigPBIReport);
    try {
      this.toastr.success("Done");
      //microsoftTeams.executeDeepLink(deepLink);
    } catch (e: any) {
      this.error.handleError(e);
    }
  }
}
