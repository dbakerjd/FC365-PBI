import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Injectable } from '@angular/core';

@Injectable({
  providedIn: 'root'
})
export class LicensingService {
  siteUrl: string = 'https://betasoftwaresl.sharepoint.com/sites/JDNPPApp/_api/web/';
  licensingApiUrl: string = ' https://jdlicensingfunctions.azurewebsites.net/api/license?code=0R6EUPw28eUEVmBU9gNfi1yEwEpX28kOUWXZtEIjxavv5qV6VacwDw==';
  
  constructor(private http: HttpClient) { }

  async hasJplusDLicense(token: string) {
    // new RequestOptions({ headers:null, withCredentials: 
    //   true });
  
    /*
    let headers = new HttpHeaders({
      // token:'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyIsImtpZCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyJ9.eyJhdWQiOiJodHRwczovL2JldGFzb2Z0d2FyZXNsLnNoYXJlcG9pbnQuY29tIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvZTc4ZmFiM2UtMjQ4ZS00NDJjLWIyZTEtMTVmMzA5ZTlkMjc2LyIsImlhdCI6MTYyOTU0MzcwNCwibmJmIjoxNjI5NTQzNzA0LCJleHAiOjE2Mjk1NDc2MDQsImFjciI6IjEiLCJhaW8iOiJFMlpnWUFqbXQrRW9pYk5JRnZhV1pIYXN2eHg3aTJPZVJSM0htK2pmMWY4ZTFzWUh5cTF6NUFzTWxTbDhKQ25JKzdDazlEQVRBQT09IiwiYW1yIjpbInB3ZCJdLCJhcHBfZGlzcGxheW5hbWUiOiJOUFAgRGVtbyIsImFwcGlkIjoiMTc1MzRjYTItZjRmOC00M2MwLTg2MTItNzJiZGQyOWE5ZWU4IiwiYXBwaWRhY3IiOiIwIiwiZmFtaWx5X25hbWUiOiJNYcOxw6kiLCJnaXZlbl9uYW1lIjoiQWxiZXJ0IiwiaWR0eXAiOiJ1c2VyIiwiaXBhZGRyIjoiNjIuNTcuMjkuMzEiLCJuYW1lIjoiQWxiZXJ0IE1hw7HDqSIsIm9pZCI6IjkzMmI2Y2QwLTc4NjgtNDgxYS05Mzc0LWE5NDY2ODgzYzRmMyIsInB1aWQiOiIxMDAzMjAwMTZCMUMwMEEwIiwicmgiOiIwLkFZRUFQcXVQNTQ0a0xFU3k0Ulh6Q2VuU2RxSk1VeGY0OU1CRGhoSnl2ZEthbnVpQkFDNC4iLCJzY3AiOiJBbGxTaXRlcy5GdWxsQ29udHJvbCBVc2VyLlJlYWQiLCJzaWQiOiIxNDM2YzE0ZS1hOTE5LTRjNTQtOGEzMi1jMjMxMmJjN2QzM2MiLCJzdWIiOiJTckFoNHZmQVU4UDJxS0RCNE1oVDZITVZmYXJuc3AwRXBUYlF1R2g1X2VVIiwidGlkIjoiZTc4ZmFiM2UtMjQ4ZS00NDJjLWIyZTEtMTVmMzA5ZTlkMjc2IiwidW5pcXVlX25hbWUiOiJhbGJlcnRAYmV0YXNvZnR3YXJlc2wub25taWNyb3NvZnQuY29tIiwidXBuIjoiYWxiZXJ0QGJldGFzb2Z0d2FyZXNsLm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6ImFfLTJoaUw4NWthVWVPX3VlT1pWQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdfQ.PSm_iZVNGCSpgXFD0JvLSftoLHbu6uzuvWVFVRFSQJowPrr8q5aoOQlaolM7TEX8QlHilM-c70g4gt6TDqqEn6YGjhzt2pIEfPKQFc6iUues-FO7KoAjnY0x-CMl_59lO_OWQ882mi6o2-49lI19Yr6nEwU5ts8R4ElEatekAgT6o87twH1uBgiVAfRK4Wj2JjL9HXlgV8GUj0mHX25AO5C1iifogvkDe5kkDex-4LdyefRLr5_2m46cNUw3yycs_BsWZU4FmdzW0w9wFtA9DVhXtKiR6SQci2ti_JO6_3T49jpze7hwa2GUgKKP101dyISbaI3VuRsfC9bAOWGZbQ',
      'Content-Type': 'text/plain', 
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET'
    });
    console.log('headers license', headers);
    return await this.http.get(this.licensingApiUrl, { 
      headers: headers
    }).toPromise();
    */
    
    return {
      "Tier": "silver",
      "Expiration": "2021-07-29T00:00:00",
      "SharePointUri": 'https://betasoftwaresl.sharepoint.com/sites/JDNPPApp/_api/web/'
    }
  }
}
