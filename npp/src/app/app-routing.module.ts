import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { ActionsListComponent } from './actions/actions-list/actions-list.component';
import { AuthEndComponent } from './auth/auth-end/auth-end.component';
import { AuthStartComponent } from './auth/auth-start/auth-start.component';
import { BrandListComponent } from './brand/brand-list/brand-list.component';
import { DashboardComponent } from './dashboard/dashboard.component';
import { ExpiredLicenseComponent } from './expired-license/expired-license.component';
import { FilesListComponent } from './files/files-list/files-list.component';
import { NotFoundComponent } from './not-found/not-found.component';
import { OpportunityListComponent } from './opportunity/opportunity-list/opportunity-list.component';
import { PowerBiComponent } from './power-bi/power-bi.component';
import { SummaryComponent } from './summary/summary.component';

const routes: Routes = [
  { path: '', component: DashboardComponent, data: { breadcrumb: 'Home' } },
  { path: 'summary', component: SummaryComponent, data: { breadcrumb: 'NPP Summary' } },
  { path: 'opportunities', component: OpportunityListComponent, data: { breadcrumb: 'NPP Opportunity Assessment'} },
  { path: 'opportunities/:id/actions', component: ActionsListComponent, data: { breadcrumb: { alias: 'opportunityName' } } },
  { path: 'opportunities/:id/files', component: FilesListComponent},
  { path: 'brands', component: BrandListComponent, data: { breadcrumb: 'Brands' } },
  { path: 'brands/:id/files', component: FilesListComponent},
  { path: 'power-bi', component: PowerBiComponent, data: { breadcrumb: 'Power BI' } },
  { path: 'expired-license', component: ExpiredLicenseComponent, data: { breadcrumb: { skip: true } } },
  { path: 'auth-start', component: AuthStartComponent },
  { path: 'auth-end', component: AuthEndComponent },
  { path: '**', pathMatch: 'full', component: NotFoundComponent, data: { breadcrumb: { skip: true } } },
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { }
