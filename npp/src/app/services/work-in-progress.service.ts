import { Injectable } from '@angular/core';
import { Subject } from 'rxjs';

export interface NPPJob {
  id: string;
  name: string;
  startTime: Date;
  error?: Error;
  status?: string;
}

@Injectable({
  providedIn: 'root'
})
export class WorkInProgressService {

  jobs: NPPJob[] = [];
  public working = new Subject<NPPJob[]>();
  public idle = new Subject<boolean>();

  constructor() { 
  }

  startJob(name: string) {
    let id = Date.now() + ' ' + Math.floor(Math.random() * 100);
    let job = {
      id,
      name,
      startTime: new Date()
    }

    this.jobs.push(job);
    this.notify();
    return job;
  }

  finishJob(id: string) {
    this.jobs = this.jobs.filter(el => el.id != id);
    this.notify();
  }

  getWrokingSubject() {
    return this.working;
  }

  getIdleSubject() {
    return this.idle;
  }

  notify() {
    if(this.jobs.length) {
      this.working.next(this.jobs);

      window.onbeforeunload = function (e: any) {
        e = e || window.event;
    
        // For IE and Firefox prior to version 4
        if (e) {
            e.returnValue = 'Some content is still being created/updated, please wait a moment to avoid inconsistent data.';
        }
    
        // For Safari
        return 'Some content is still being created/updated, please wait a moment to avoid inconsistent data.';
      };
    } else {
      
      this.idle.next(true);
      window.onbeforeunload = null;

    }
  }


}
