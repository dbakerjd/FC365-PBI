import { Component, OnInit } from '@angular/core';
import { ActivatedRoute } from '@angular/router';
import { environment } from 'src/environments/environment';

@Component({
  selector: 'app-splash-screen',
  templateUrl: './splash-screen.component.html',
  styleUrls: ['./splash-screen.component.scss']
})
export class SplashScreenComponent implements OnInit {

  version = environment.version;
  messageToShow = '';
  appTitle = 'NPP';
  client: any;
  
  constructor(private readonly route: ActivatedRoute) { }

  ngOnInit(): void {
    this.client = environment.contact;
    this.route.params.subscribe(async (params) => {
      if (params.message && typeof params.message === 'string') {
        this.messageToShow = params.message;
        if(environment.isInlineApp) { 
          this.appTitle = 'Inline';
        }
      }
    });
  }

}
