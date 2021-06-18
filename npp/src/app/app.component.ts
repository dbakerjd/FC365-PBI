import { Component } from '@angular/core';
import { TeamsService } from './teams.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  constructor(private readonly teams: TeamsService) {

  }
  ngOnInit(): void {
  }
}
