import { Injectable } from '@angular/core';
import { MatDialog } from '@angular/material/dialog';
import { Subject } from 'rxjs';
import { BlockDialogComponent } from '../modals/block-dialog/block-dialog.component';

export interface NPPJob {
  id: string;
  name: string;
  startTime: Date;
  dialogInstance: any;
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

  constructor(public matDialog: MatDialog) { 
  }

  startJob(name: string, dialogMessage: string | null = null) {
    let id = Date.now() + ' ' + Math.floor(Math.random() * 100);
    const dialogInstance = this.matDialog.open(BlockDialogComponent, {
      height: '300px',
      width: '405px',
      data: {
        message: dialogMessage ? dialogMessage : null
      }
    });

    let job = {
      id,
      name,
      dialogInstance,
      startTime: new Date()
    }

    this.jobs.push(job);
    this.notify();
    return job;
  }

  async finishJob(id: string) {
    const job = this.jobs.find(el => el.id === id);
    if (job) {
      await job.dialogInstance.close();
      this.jobs = this.jobs.filter(el => el.id != id);
    }
    this.notify();
  }

  getWorkingSubject() {
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
