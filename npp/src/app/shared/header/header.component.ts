import { Component, Input, OnInit } from '@angular/core';
import { SharepointService } from 'src/app/services/sharepoint.service';
import { WorkInProgressService } from 'src/app/services/work-in-progress.service';

@Component({
  selector: 'app-npp-header',
  templateUrl: './header.component.html',
  styleUrls: ['./header.component.scss']
})
export class HeaderComponent implements OnInit {
  @Input() isHome = false;

  public notificationsCounter = 0;

  constructor(
    public jobs: WorkInProgressService, 
    private readonly sharepoint: SharepointService
  ) { }

  async ngOnInit() {
    this.notificationsCounter = await (await this.sharepoint.getUserNotifications((await this.sharepoint.getCurrentUserInfo()).Id)).length;
  }

  goBack() {
    window.history.back();
  }

}
