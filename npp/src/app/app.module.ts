import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { DashboardComponent } from './dashboard/dashboard.component';
import { ErrorService } from './services/error.service';
import { TeamsService } from './services/teams.service';
import { HTTP_INTERCEPTORS, HttpClientModule } from '@angular/common/http';
import { IPublicClientApplication, PublicClientApplication, InteractionType, BrowserCacheLocation, LogLevel } from '@azure/msal-browser';
import { MsalGuard, MsalInterceptor, MsalBroadcastService, MsalInterceptorConfiguration, MsalModule, MsalService, MSAL_GUARD_CONFIG, MSAL_INSTANCE, MSAL_INTERCEPTOR_CONFIG, MsalGuardConfiguration, MsalRedirectComponent } from '@azure/msal-angular';
import { OpportunityListComponent } from './opportunity/opportunity-list/opportunity-list.component';
import { OpportunityDetailComponent } from './opportunity/opportunity-detail/opportunity-detail.component';
import { ActionsListComponent } from './actions/actions-list/actions-list.component';
import { HeaderComponent } from './shared/header/header.component';
import { NotificationsListComponent } from './shared/notifications-list/notifications-list.component';
import { UserProfilePicComponent } from './shared/user-profile-pic/user-profile-pic.component';
import { SummaryComponent } from './summary/summary.component';
import { PowerBiComponent } from './power-bi/power-bi.component';
import { SharepointService } from './services/sharepoint.service';
import { ReactiveFormsModule } from '@angular/forms';
import { FormlyModule } from '@ngx-formly/core';
import { FormlyBootstrapModule } from '@ngx-formly/bootstrap';
import { ProgressBarComponent } from './shared/progress-bar/progress-bar.component';
import { environment } from 'src/environments/environment';
import { LicensingService } from './services/licensing.service';
import { DatepickerModule } from 'ng2-datepicker';
import { UploadFileComponent } from './modals/upload-file/upload-file.component';
import { MatButtonModule } from '@angular/material/button';
import { MatDialogModule } from '@angular/material/dialog';
import { MatProgressSpinnerModule } from '@angular/material/progress-spinner';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { DialogHeaderComponent } from './modals/dialog-header/dialog-header.component';
import { FormlyFieldFile } from './shared/formly-fields/file-input';
import { FileValueAccessor } from './shared/formly-fields/file-value-accessor';
import { SendForApprovalComponent } from './modals/send-for-approval/send-for-approval.component';
import { CreateScenarioComponent } from './modals/create-scenario/create-scenario.component';
import { CreateOpportunityComponent } from './modals/create-opportunity/create-opportunity.component';
import { FormlyTypesModule, FORMLY_CONFIG } from './shared/formly-fields/formly-types.module';
import { StageSettingsComponent } from './modals/stage-settings/stage-settings.component';
import { FilterPipe } from './filter.pipe';
import { SortPipe } from './sort.pipe';
import { ConfirmDialogComponent } from './modals/confirm-dialog/confirm-dialog.component';
import { ProgressSpinnerComponent } from './shared/progress-spinner/progress-spinner.component';
import { NotFoundComponent } from './not-found/not-found.component';
import { ShareDocumentComponent } from './modals/share-document/share-document.component';
import { ToastrModule } from 'ngx-toastr';
import { ChartModule, HIGHCHARTS_MODULES } from 'angular-highcharts';
import * as xrange from 'highcharts/modules/xrange.src';
import { ExpiredLicenseComponent } from './expired-license/expired-license.component';
import { SafePipe } from './shared/safe.pipe';
import { WorkInProgressService } from './services/work-in-progress.service';
import { FolderPermissionsComponent } from './modals/folder-permissions/folder-permissions.component';
import { AuthStartComponent } from './auth/auth-start/auth-start.component';
import { AuthEndComponent } from './auth/auth-end/auth-end.component';

const isIE = window.navigator.userAgent.indexOf("MSIE ") > -1 || window.navigator.userAgent.indexOf("Trident/") > -1; // Remove this line to use Angular Universal

export function loggerCallback(logLevel: LogLevel, message: string) {
  console.log(message);
}

export function MSALInstanceFactory(): IPublicClientApplication {
  return new PublicClientApplication({
    auth: {
      // clientId: '6226576d-37e9-49eb-b201-ec1eeb0029b6', // Prod enviroment. Uncomment to use. 
      clientId: '17534ca2-f4f8-43c0-8612-72bdd29a9ee8', // PPE testing environment
      authority: 'https://login.microsoftonline.com/common', // Prod environment. Uncomment to use.
      //authority: 'https://login.windows-ppe.net/common', // PPE testing environment.
      redirectUri: environment.ssoRedirectUrl,
      postLogoutRedirectUri: environment.ssoRedirectUrl
    },
    cache: {
      cacheLocation: BrowserCacheLocation.LocalStorage,
      storeAuthStateInCookie: isIE, // set to true for IE 11. Remove this line to use Angular Universal
    },
    system: {
      loggerOptions: {
        loggerCallback,
        logLevel: LogLevel.Info,
        piiLoggingEnabled: false
      }
    }
  });
}

export function MSALInterceptorConfigFactory(): MsalInterceptorConfiguration {
  console.log('intercepted request');
  const protectedResourceMap = new Map<string, Array<string>>();
  // protectedResourceMap.set('https://betasoftwaresl.sharepoint.com', ['AllSites.FullControl', 'AllSites.Manage', 'Sites.Search.All']);
  protectedResourceMap.set('https://betasoftwaresl.sharepoint.com', ['https://betasoftwaresl.sharepoint.com/.default']);

  return {
    interactionType: InteractionType.Redirect,
    protectedResourceMap
  };
}

export function MSALGuardConfigFactory(): MsalGuardConfiguration {
  return { 
    interactionType: InteractionType.Redirect,
    authRequest: {
      scopes: ['https://betasoftwaresl.sharepoint.com/.default']
    },
    loginFailedRoute: '/auth-end'
  };
}

@NgModule({
  imports: [
    BrowserModule,
    BrowserAnimationsModule,
    AppRoutingModule,
    HttpClientModule,
    MsalModule,
    ReactiveFormsModule,
    DatepickerModule,
    FormlyTypesModule,
    ChartModule,
    FormlyModule.forRoot(FORMLY_CONFIG),
    FormlyBootstrapModule,
    MatButtonModule,
    MatDialogModule,
    MatProgressSpinnerModule,
    ToastrModule.forRoot({
      positionClass: 'toast-bottom-right',
    }),
  ],
  declarations: [
    AppComponent,
    DashboardComponent,
    OpportunityListComponent,
    OpportunityDetailComponent,
    ActionsListComponent,
    HeaderComponent,
    NotificationsListComponent,
    UserProfilePicComponent,
    SummaryComponent,
    PowerBiComponent,
    ProgressBarComponent,
    UploadFileComponent,
    DialogHeaderComponent,
    SendForApprovalComponent,
    CreateScenarioComponent,
    CreateOpportunityComponent,
    StageSettingsComponent,
    FolderPermissionsComponent,
    ShareDocumentComponent,
    FilterPipe,
    SortPipe,
    ConfirmDialogComponent,
    ProgressSpinnerComponent,
    NotFoundComponent,
    ExpiredLicenseComponent,
    SafePipe,
    AuthStartComponent,
    AuthEndComponent
  ],
  providers: [
    {
      provide: HTTP_INTERCEPTORS,
      useClass: MsalInterceptor,
      multi: true
    },
    {
      provide: MSAL_INSTANCE,
      useFactory: MSALInstanceFactory
    },
    {
      provide: MSAL_GUARD_CONFIG,
      useFactory: MSALGuardConfigFactory
    },
    {
      provide: MSAL_INTERCEPTOR_CONFIG,
      useFactory: MSALInterceptorConfigFactory
    },
    { provide: HIGHCHARTS_MODULES, useFactory: () => [ xrange ] }, // add as factory to your providers
    MsalService,
    MsalGuard,
    MsalBroadcastService,
    TeamsService,
    ErrorService,
    SharepointService,
    LicensingService,
    WorkInProgressService
  ],
  bootstrap: [AppComponent, MsalRedirectComponent]
})
export class AppModule { }
