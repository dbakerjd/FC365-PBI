import { Component, Input, OnInit } from '@angular/core';
import { WorkInProgressService } from 'src/app/services/work-in-progress.service';

@Component({
  selector: 'app-npp-header',
  templateUrl: './header.component.html',
  styleUrls: ['./header.component.scss']
})
export class HeaderComponent implements OnInit {
  @Input() isHome = false;
  workInProgress = false;
  constructor(public jobs: WorkInProgressService) { }

  ngOnInit(): void {
    this.jobs.getWrokingSubject().subscribe(() => {
      this.workInProgress = true;
    });

    this.jobs.getIdleSubject().subscribe(() => {
      this.workInProgress = false;
    });
  }

  goBack() {
    window.history.back();
  }

}
