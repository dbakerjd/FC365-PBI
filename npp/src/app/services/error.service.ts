import { Injectable } from '@angular/core';

@Injectable({
  providedIn: 'root'
})
export class ErrorService {

  constructor() { }

  handleError(e: Error) {
    console.log(e);
  }
}
