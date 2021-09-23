import { Injectable } from '@angular/core';
import { ToastrService } from 'ngx-toastr';

@Injectable({
  providedIn: 'root'
})
export class ErrorService {

  constructor(public toastr: ToastrService) { }

  handleError(e: Error) {
    this.toastr.error(e.message);
  }
}
