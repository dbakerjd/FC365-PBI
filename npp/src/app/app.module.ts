import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { DashboardComponent } from './dashboard/dashboard.component';
import { ErrorService } from './services/app/error.service';
import { HttpClientModule, HTTP_INTERCEPTORS } from '@angular/common/http';
import { OpportunityListComponent } from './opportunity/opportunity-list/opportunity-list.component';
import { ActionsListComponent } from './actions/actions-list/actions-list.component';
import { HeaderComponent } from './shared/header/header.component';
import { NotificationsListComponent } from './shared/notifications-list/notifications-list.component';
import { UserProfilePicComponent } from './shared/user-profile-pic/user-profile-pic.component';
import { SummaryComponent } from './summary/summary.component';
import { PowerBiComponent } from './power-bi/power-bi.component';
import { SharepointService } from './services/microsoft-data/sharepoint.service';
import { ReactiveFormsModule } from '@angular/forms';
import { FormlyModule } from '@ngx-formly/core';
import { FormlyBootstrapModule } from '@ngx-formly/bootstrap';
import { ProgressBarComponent } from './shared/progress-bar/progress-bar.component';
import { LicensingService } from './services/jd-data/licensing.service';
import { DatepickerModule } from 'ng2-datepicker';
import { UploadFileComponent } from './modals/upload-file/upload-file.component';
import { MatButtonModule } from '@angular/material/button';
import { MatDialogModule } from '@angular/material/dialog';
import { MatProgressSpinnerModule } from '@angular/material/progress-spinner';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { DialogHeaderComponent } from './modals/dialog-header/dialog-header.component';
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
import { SplashScreenComponent } from './splash-screen/splash-screen.component';
import { SafePipe } from './shared/safe.pipe';
import { WorkInProgressService } from './services/app/work-in-progress.service';
import { FolderPermissionsComponent } from './modals/folder-permissions/folder-permissions.component';
import { AuthStartComponent } from './auth/auth-start/auth-start.component';
import { AuthEndComponent } from './auth/auth-end/auth-end.component';
import { AuthInterceptor } from './auth/auth.interceptor';
import { BlockDialogComponent } from './modals/block-dialog/block-dialog.component';
import { EditFileComponent } from './modals/edit-file/edit-file.component';
import { RejectModelComponent } from './modals/reject-model/reject-model.component';
import { PowerBIEmbedModule } from 'powerbi-client-angular';
import { FilesListComponent } from './files/files-list/files-list.component';
import { CreateForecastCycleComponent } from './modals/create-forecast-cycle/create-forecast-cycle.component';
import { CommentsListComponent } from './modals/comments-list/comments-list.component';
import { ApproveModelComponent } from './modals/approve-model/approve-model.component';
import { EntityEditFileComponent } from './modals/entity-edit-file/entity-edit-file.component';
import { CreateBrandComponent } from './modals/create-brand/create-brand.component';
import { BrandListComponent } from './brand/brand-list/brand-list.component';
import { BrandSummaryComponent } from './brand/brand-summary/brand-summary.component';
import { MatTooltipModule } from '@angular/material/tooltip';
import { BreadcrumbsComponent } from './breadcrumbs/breadcrumbs.component';
import { EntitiesService } from './services/entities.service';
import { TeamsService } from '@services/microsoft-data/teams.service';
import { AppControlService } from '@services/app/app-control.service';
import { SeatsTableComponent } from './seats-table/seats-table.component';
import { PowerBiService } from '@services/microsoft-data/power-bi.service';
import { GeneralAreaComponent } from './general-area/general-area.component';

@NgModule({
  imports: [
    BrowserModule,
    BrowserAnimationsModule,
    AppRoutingModule,
    HttpClientModule,
    ReactiveFormsModule,
    DatepickerModule,
    FormlyTypesModule,
    ChartModule,
    FormlyModule.forRoot(FORMLY_CONFIG),
    FormlyBootstrapModule,
    MatButtonModule,
    MatDialogModule,
    MatTooltipModule,
    MatProgressSpinnerModule,
    PowerBIEmbedModule,  
    ToastrModule.forRoot({
      positionClass: 'toast-bottom-right',
      timeOut: 7000,
    }),
  ],
  declarations: [
    AppComponent,
    DashboardComponent,
    OpportunityListComponent,
    FilesListComponent,
    ActionsListComponent,
    HeaderComponent,
    NotificationsListComponent,
    UserProfilePicComponent,
    SummaryComponent,
    PowerBiComponent,
    ProgressBarComponent,
    UploadFileComponent,
    EditFileComponent,
    DialogHeaderComponent,
    SendForApprovalComponent,
    RejectModelComponent,
    CreateScenarioComponent,
    CreateOpportunityComponent,
    StageSettingsComponent,
    FolderPermissionsComponent,
    ShareDocumentComponent,
    FilterPipe,
    SortPipe,
    ConfirmDialogComponent,
    BlockDialogComponent,
    ProgressSpinnerComponent,
    NotFoundComponent,
    SplashScreenComponent,
    SafePipe,
    AuthStartComponent,
    AuthEndComponent,
    CreateForecastCycleComponent,
    CommentsListComponent,
    ApproveModelComponent,
    EntityEditFileComponent,
    CreateBrandComponent,
    BrandListComponent,
    BrandSummaryComponent,
    BreadcrumbsComponent,
    SeatsTableComponent,
    GeneralAreaComponent,
  ],
  providers: [
    { provide: HTTP_INTERCEPTORS, useClass: AuthInterceptor, multi: true },
    { provide: HIGHCHARTS_MODULES, useFactory: () => [ xrange ] }, // add as factory to your providers
    TeamsService,
    ErrorService,
    SharepointService,
    LicensingService,
    WorkInProgressService,
    PowerBiService,
    AppControlService,
    EntitiesService
  ],
  bootstrap: [AppComponent]
})
export class AppModule { }
