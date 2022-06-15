import { Injectable } from '@angular/core';
import { ToastrService } from 'ngx-toastr';
import { Subject } from 'rxjs';

@Injectable({
  providedIn: 'root'
})
export class ErrorService {
  public subject = new Subject<string>();
  constructor(public toastr: ToastrService) { }

  handleError(e: any) {
    let errorMessage = e.message;
    if(e.status) {
      if (e.status == 403) this.subject.next('unauthorized');
      else if (e.status === 423) errorMessage = "The file is locked by another user right now. Try again later";
    }
    this.toastr.error(errorMessage);
  }
}
