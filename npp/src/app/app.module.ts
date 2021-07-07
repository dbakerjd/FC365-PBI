import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { DashboardComponent } from './dashboard/dashboard.component';
import { ErrorService } from './services/error.service';
import { TeamsService } from './services/teams.service';
import { AuthStartComponent } from './auth/auth-start/auth-start.component';
import { AuthEndComponent } from './auth/auth-end/auth-end.component';

@NgModule({
  declarations: [
    AppComponent,
    DashboardComponent,
    AuthStartComponent,
    AuthEndComponent
  ],
  imports: [
    BrowserModule,
    AppRoutingModule
  ],
  providers: [TeamsService, ErrorService],
  bootstrap: [AppComponent]
})
export class AppModule { }
