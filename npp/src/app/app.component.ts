import { Component } from '@angular/core';
import { MsalService } from '@azure/msal-angular';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  constructor(private authService: MsalService) {

  }
  ngOnInit(): void {
    this.authService.handleRedirectObservable().subscribe();
  }
}
