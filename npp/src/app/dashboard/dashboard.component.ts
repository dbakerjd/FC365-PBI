import { Component, OnInit } from '@angular/core';
import { TeamsService } from '../teams.service';

@Component({
  selector: 'app-dashboard',
  templateUrl: './dashboard.component.html',
  styleUrls: ['./dashboard.component.scss']
})
export class DashboardComponent implements OnInit {

  constructor(private readonly teams: TeamsService) { }

  ngOnInit(): void {
  }

  getUser() {
    return this.teams.user;
  }

  getContext()  {
    return this.teams.context;
  }

  getToken()  {
    return this.teams.token;
  }
}

