import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { ActionsListComponent } from './actions/actions-list/actions-list.component';
import { AuthEndComponent } from './auth/auth-end/auth-end.component';
import { AuthStartComponent } from './auth/auth-start/auth-start.component';
import { DashboardComponent } from './dashboard/dashboard.component';
import { ExpiredLicenseComponent } from './expired-license/expired-license.component';
import { FilesListComponent } from './files/files-list/files-list.component';
import { NotFoundComponent } from './not-found/not-found.component';
import { OpportunityListComponent } from './opportunity/opportunity-list/opportunity-list.component';
import { PowerBiComponent } from './power-bi/power-bi.component';
import { SummaryComponent } from './summary/summary.component';

const routes: Routes = [
  { path: '', component: DashboardComponent },
  { path: 'summary', component: SummaryComponent },
  { path: 'opportunities', component: OpportunityListComponent },
  { path: 'opportunities/:id/actions', component: ActionsListComponent },
  { path: 'opportunities/:id/files', component: FilesListComponent},
  { path: 'power-bi', component: PowerBiComponent },
  { path: 'expired-license', component: ExpiredLicenseComponent },
  { path: 'auth-start', component: AuthStartComponent },
  { path: 'auth-end', component: AuthEndComponent },
  { path: '**', pathMatch: 'full', component: NotFoundComponent },
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { }
