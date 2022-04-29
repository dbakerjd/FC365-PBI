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
    this.toastr.error(e.message);
    if(e.status && e.status == 403) {
      this.subject.next('unauthorized');
    }
  }
}
