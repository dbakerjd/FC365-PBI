import { Component, Input, OnInit } from '@angular/core';
import { WorkInProgressService } from 'src/app/services/work-in-progress.service';

@Component({
  selector: 'app-npp-header',
  templateUrl: './header.component.html',
  styleUrls: ['./header.component.scss']
})
export class HeaderComponent implements OnInit {
  @Input() isHome = false;
  constructor(public jobs: WorkInProgressService) { }

  ngOnInit(): void {
  }

  goBack() {
    window.history.back();
  }

}
