import { Injectable } from '@angular/core';
import { HttpEvent, HttpHandler, HttpInterceptor, HttpRequest } from '@angular/common/http';
import { from, Observable } from 'rxjs';
import { TeamsService } from '../services/teams.service';
import { InteractionRequiredAuthError } from '@azure/msal-common';
@Injectable()
export class AuthInterceptor implements HttpInterceptor {

    constructor(public teams: TeamsService) {}

    intercept(req: HttpRequest<any>, next: HttpHandler): Observable<HttpEvent<any>> {
        return from(this.handle(req, next));
    }

    async handle(req: HttpRequest<any>, next: HttpHandler) {
        
        try {
            let pathArr = req.url.split("/");
            let domain = pathArr[2];
            console.log("trying to obtain token for domain: "+domain);
            let scopes = this.teams.getResourceByDomain(domain);
            console.log(scopes);
            this.teams.hackyConsole += "********** INTERCEPTOR ************  DOMAIN: "+ domain +"      --------------          ";
            this.teams.hackyConsole += "********** INTERCEPTOR ************  SCOPES: "+JSON.stringify(scopes) +"      --------------          ";
            if(scopes) {
                let request = {
                    scopes
                }
        
                let tokenResponse = await this.teams.msalInstance.acquireTokenSilent(request);
                this.teams.hackyConsole += "********** INTERCEPTOR ************  TOKEN: "+JSON.stringify(tokenResponse) +"      --------------          ";
                const token = tokenResponse.accessToken;
                
                if (!token) {
                    return next.handle(req).toPromise();
                }
                const headers = req.clone({
                    headers: req.headers.set('Authorization', `Bearer ${token}`)
                });

                return next.handle(headers).toPromise();
                
            } else {
                return next.handle(req).toPromise();
            }
        } catch(e) {
            this.teams.hackyConsole += "********** INTERCEPTOR ************  ERROR: "+JSON.stringify(e) +"      --------------          ";
            if (e instanceof InteractionRequiredAuthError) {
                // fallback to interaction when silent call fails
                this.teams.login();
            }
            return next.handle(req).toPromise();
        }
    }
}