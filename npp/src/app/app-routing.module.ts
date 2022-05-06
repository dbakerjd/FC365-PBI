import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { ActionsListComponent } from './actions/actions-list/actions-list.component';
import { AuthEndComponent } from './auth/auth-end/auth-end.component';
import { AuthStartComponent } from './auth/auth-start/auth-start.component';
import { BrandListComponent } from './brand/brand-list/brand-list.component';
import { BrandSummaryComponent } from './brand/brand-summary/brand-summary.component';
import { DashboardComponent } from './dashboard/dashboard.component';
import { SplashScreenComponent } from './splash-screen/splash-screen.component';
import { FilesListComponent } from './files/files-list/files-list.component';
import { NotFoundComponent } from './not-found/not-found.component';
import { OpportunityListComponent } from './opportunity/opportunity-list/opportunity-list.component';
import { PowerBiComponent } from './power-bi/power-bi.component';
import { SummaryComponent } from './summary/summary.component';
import { GeneralAreaComponent } from './general-area/general-area.component';

const routes: Routes = [
  { path: '', component: DashboardComponent, data: { breadcrumb: 'Home' } },
  { path: 'summary', component: SummaryComponent, data: { breadcrumb: 'NPP Summary' } },
  { path: 'brands-summary', component: BrandSummaryComponent },
  { path: 'general-area', component: GeneralAreaComponent, data: { breadcrumb: 'General Area' } },
  { path: 'opportunities', component: OpportunityListComponent, data: { breadcrumb: 'NPP Opportunity Assessment'} },
  { path: 'opportunities/:id/actions', component: ActionsListComponent, data: { breadcrumb: { label: 'NPP Opportunity Assessment', url: 'opportunities' } } },
  { path: 'opportunities/:id/files', component: FilesListComponent, data: { breadcrumb: { label: 'NPP Opportunity Assessment', url: 'opportunities' } } },
  { path: 'brands', component: BrandListComponent, data: { breadcrumb: 'Brands' } },
  { path: 'brands/:id/files', component: FilesListComponent, data: { breadcrumb: { label: 'Brands', url: 'brands' } } },
  { path: 'power-bi', component: PowerBiComponent, data: { breadcrumb: 'Power BI' } },
  { path: 'splash/:message', component: SplashScreenComponent },
  { path: 'auth-start', component: AuthStartComponent },
  { path: 'auth-end', component: AuthEndComponent },
  { path: '**', pathMatch: 'full', component: NotFoundComponent },
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { }
