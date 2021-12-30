import { Component, Input, OnInit } from '@angular/core';
import { NotificationsService } from 'src/app/services/notifications.service';
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
    private readonly notifications: NotificationsService
  ) { }

  async ngOnInit() {
    this.notificationsCounter = await this.notifications.getUnreadNotifications();
  }

  goBack() {
    window.history.back();
  }

}
