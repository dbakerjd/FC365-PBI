import { Component, Input, OnInit } from '@angular/core';
import { Router } from '@angular/router';
import { NotificationsService } from 'src/app/services/notifications.service';
import { WorkInProgressService } from '@services/app/work-in-progress.service';
import { environment } from 'src/environments/environment';

@Component({
  selector: 'app-npp-header',
  templateUrl: './header.component.html',
  styleUrls: ['./header.component.scss']
})
export class HeaderComponent implements OnInit {
  @Input() currentSection = 'none';
  isInline: boolean = false;

  public notificationsCounter = 0;

  constructor(
    public jobs: WorkInProgressService, 
    private readonly notifications: NotificationsService,
    private router: Router
  ) { }

  async ngOnInit() {
    this.isInline = environment.isInlineApp;
    if (this.router.url != '/summary') { // si summary, continuar a 0
      this.notificationsCounter = await this.notifications.getUnreadNotifications();
    }
    setInterval(async () => this.notificationsCounter = await this.notifications.getUnreadNotifications(), 60000);
  }

  clearNotifications() {
    this.notificationsCounter = 0;
  }

  goBack() {
    window.history.back();
  }

}
