import { Component, OnInit, ViewChild } from '@angular/core';
import { TeamsService } from '../services/teams.service';
import { HttpClient } from '@angular/common/http';
import { ActivatedRoute, Router } from '@angular/router';
import { IReportEmbedConfiguration, models, Report } from 'powerbi-client';
import { PowerBIReportEmbedComponent } from 'powerbi-client-angular';
import { LicensingService } from '../services/licensing.service';
import * as microsoftTeams from "@microsoft/teams-js";
import { ErrorService } from '../services/error.service';
import { ToastrService } from 'ngx-toastr';
import { PBIReport, SharepointService } from '../services/sharepoint.service';
import { PageDetails, PowerBiService } from '../services/power-bi.service';

@Component({
  selector: 'app-power-bi',
  templateUrl: './power-bi.component.html',
  styleUrls: ['./power-bi.component.scss']
})

export class PowerBiComponent implements OnInit {

  @ViewChild(PowerBIReportEmbedComponent) reportObj!: PowerBIReportEmbedComponent;

  pagesList: PageDetails[] = [];
  page!: PageDetails;
  
  pbireports: PBIReport[] = [];
  pbireport: PBIReport | undefined = undefined;
  
  oppID: number[] = []
  ID!: number;
  
  DisplayName!: string;

  displayMessage!: string;
  filters!: models.ReportLevelFilters;

  pageName!: string;

  reportClass = 'report-container desktop-view'

  showPages = false;
  settingsHidden = true;

  actionbarVisible: boolean = false;
  actionbarOption: string = "Show";

  filtersVisible: boolean = false;
  filterPaneOption: string = "Show";

  bookmarksVisible: boolean = false;
  bookmarksOption: string = "Show";

  highContrastMode: models.ContrastMode = models.ContrastMode.None;
  highContrastOption: string = "On";



  reportConfig: IReportEmbedConfiguration = {
    type: 'report',
    embedUrl: undefined,
    tokenType: undefined,
    accessToken: undefined,
    id: undefined,
    settings: undefined,
    filters: undefined,
    pageName: undefined
  };

  constructor(
    public licensing: LicensingService,
    private sharepoint: SharepointService,
    private pbi: PowerBiService,
    public teams: TeamsService,
    private error: ErrorService,
    private toastr: ToastrService,
    private router: Router,
    private route: ActivatedRoute,

  ) {

    this.route.params.subscribe(params => {
      this.oppID.push(+params['ID']);

    })

  }

  ngOnInit(): void {

    this.getReportNames().then(getReports => this.embedReport(this.highContrastMode, getReports[0].ID))
    
  }

  async getReportNames(): Promise<PBIReport[]> {
    return this.pbireports = await this.sharepoint.getReports();
  }

  async setEmbed(ID: number) {
    
    if (ID && ID != this.ID) {
      console.log("change report");
      await this.embedReport(this.highContrastMode, ID);
     
    } else if (ID && ID == this.ID) {
      
      if (this.showPages) {
        this.showPages = false;
      }
      else {
        await this.getPages();
        this.showPages = true;
        
      }
    }



  }

  async getPages(){

    const report: Report = this.reportObj.getReport();
    let pages = await report.getPages();
    let pagesList: PageDetails[] = [];
    
    console.log(this.pageName);
    
    let DisplayName!: string;
    let currentpageName: string = this.pageName;
    let newPageName!: string;

    pages.forEach(function (page) {
      if (page.visibility == 0) {
        let pageItem: PageDetails = { ReportSection: page.name, DisplayName: page.displayName };
        pagesList.push(pageItem)

        if (page.name == currentpageName) {
          DisplayName = page.displayName;
          newPageName = page.name;
        }
      }

    })

    this.DisplayName = DisplayName;
    this.pageName = newPageName;
    this.pagesList = pagesList;

  }

  async pageNavigate(ReportSection: string, DisplayName: string) {
    const report: Report = this.reportObj.getReport();
    console.log(ReportSection);
    report.page(ReportSection).setActive();
    this.DisplayName = DisplayName;
    this.pageName = ReportSection;
  }

  async embedReport(highContrastMode: models.ContrastMode, ID: number): Promise<void> {
    //set pbi report
    
    this.pbireport = await this.sharepoint.getReport(ID);
    //get token
    const token = await this.pbi.getPBIToken();
    
    //set embedUrl
    let embedUrl: string = `https://app.powerbi.com/reportEmbed?reportId=${this.pbireport.ReportId}groupId=${this.pbireport.GroupId}`;

    //set required filters. Based on OppID
    this.filters = {
      $schema: "http://powerbi.com/product/schema#basic",
      target: {
        table: "Opportunities",
        column: "OpportunityID"
      },
      operator: "In",
      values: this.oppID,
      filterType: models.FilterType.Basic
    }
    //set report config
    this.reportConfig = {
      type: 'report',
      tokenType: models.TokenType.Aad,
      accessToken: token,
      embedUrl: embedUrl,
      id: this.pbireport.ReportId,
      pageName: this.pbireport.pageName,
      settings: {
        navContentPaneEnabled: false,
        filterPaneEnabled: this.filtersVisible,
        background: models.BackgroundType.Transparent
      },
      contrastMode: highContrastMode
    }

    //if navigating from an opportunity then this.OppID is a number and filters are applied.
    if (!Number.isNaN(this.oppID[0])) {
      this.reportConfig.filters = [this.filters]
    };
    
    this.pageName = this.pbireport.pageName;
    this.ID = ID;
  }
  
  async removeFilters() {
    const report: Report = this.reportObj.getReport();

    await report.updateFilters(models.FiltersOperations.RemoveAll);
  }

  async reloadReport() {
    const report: Report = this.reportObj.getReport();

    await report.reload();
  }
  async actionbar(): Promise<undefined> {
    // Get report from the wrapper component
    //const report: Report = this.reportObj.getReport();
    const report: Report = this.reportObj.getReport();

    if (!report) {
      // Prepare status message for Error

      this.displayMessage = 'Report not available.';
      console.log(this.displayMessage);
      return;
    }
    this.actionbarVisible = !this.actionbarVisible;
    // New settings to hide filter pane
    const settings = {
      bars: {
        actionBar: {
          visible: this.actionbarVisible,
        },
      },
    };

    try {
      const response = await report.updateSettings(settings);

      // Prepare status message for success
      this.displayMessage = 'Action bar altered.';
      console.log(this.displayMessage);
      console.log(response);
      if (this.actionbarVisible) this.actionbarOption = "Hide"; else this.actionbarOption = "Show";
      return;
    } catch (error) {
      console.error(error);
      return;
    }
  }
  
  async filterPane(): Promise<undefined> {
    // Get report from the wrapper component
    const report: Report = this.reportObj.getReport();

    if (!report) {
      // Prepare status message for Error

      this.displayMessage = 'Report not available.';
      console.log(this.displayMessage);
      return;
    }
    this.filtersVisible = !this.filtersVisible;
    // New settings to hide filter pane
    const settings = {
      panes: {
        filters: {
          expanded: false,
          visible: this.filtersVisible,
        },
      },
    };

    try {
      const response = await report.updateSettings(settings);

      // Prepare status message for success
      this.displayMessage = 'Filter pane is hidden.';
      console.log(this.displayMessage);
      console.log(response);
      if (this.filtersVisible) this.filterPaneOption = "Hide"; else this.filterPaneOption = "Show";
      return;
    } catch (error) {
      console.error(error);
      return;
    }
  }

  async bookmarks(): Promise<undefined> {
    // Get report from the wrapper component
    //const report: Report = this.reportObj.getReport();
    const report: Report = this.reportObj.getReport();

    if (!report) {
      // Prepare status message for Error

      this.displayMessage = 'Report not available.';
      console.log(this.displayMessage);
      return;
    }
    this.bookmarksVisible = !this.bookmarksVisible;
    // New settings to hide filter pane
    const settings = {
      panes: {
        bookmarks: {
          visible: this.bookmarksVisible,
        },
      },
    };

    try {
      const response = await report.updateSettings(settings);

      // Prepare status message for success
      this.displayMessage = 'Bookmarks altered.';
      console.log(this.displayMessage);
      console.log(response);
      if (this.bookmarksVisible) this.bookmarksOption = "Hide"; else this.bookmarksOption = "Show";
      return;
    } catch (error) {
      console.error(error);
      return;
    }
  }

  

  async highContrast(ID: number): Promise<undefined> {
    // Get report from the wrapper component
    //const report: Report = this.reportObj.getReport();
    const report: Report = this.reportObj.getReport();

    if (!report) {
      // Prepare status message for Error

      this.displayMessage = 'Report not available.';
      console.log(this.displayMessage);
      return;
    }

    if (this.highContrastMode == models.ContrastMode.None) this.highContrastMode = models.ContrastMode.HighContrastBlack; else this.highContrastMode = models.ContrastMode.None

    try {
      await this.embedReport(this.highContrastMode, ID);
      report.populateConfig(this.reportConfig, false)

      report.reload();
      // Prepare status message for success
      this.displayMessage = 'High contrast altered.';
      console.log(this.displayMessage);

      if (this.highContrastMode == models.ContrastMode.None) this.highContrastOption = "On"; else this.highContrastOption = "Off";
      return;
    } catch (error) {
      console.error(error);
      return;
    }
  }



  showSettings() {
    this.settingsHidden = !this.settingsHidden
  }



}