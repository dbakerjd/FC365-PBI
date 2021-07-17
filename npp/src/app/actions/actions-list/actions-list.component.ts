import { Component, OnInit } from '@angular/core';
import { ActivatedRoute } from '@angular/router';
import { DatepickerOptions } from 'ng2-datepicker';
import { Action, Gate, Opportunity, SharepointService } from 'src/app/services/sharepoint.service';

@Component({
  selector: 'app-actions-list',
  templateUrl: './actions-list.component.html',
  styleUrls: ['./actions-list.component.scss']
})
export class ActionsListComponent implements OnInit {
  gates: Gate[] = [];
  opportunityId = 0;
  opportunity: Opportunity | undefined = undefined;
  currentGate: Gate | undefined = undefined;
  currentActions: Action[] | undefined = undefined;
  currentGateProgress: number = 0;
  dateOptions: DatepickerOptions = {
    format: 'M/d/Y'
  };
  constructor(private sharepoint: SharepointService, private route: ActivatedRoute) { }

  ngOnInit(): void {
    this.route.params.subscribe(async (params) => {
      if(params.id && params.id != this.opportunityId) {
        this.opportunityId = params.id;
        this.opportunity = await this.sharepoint.getOpportunity(params.id);
        this.gates = await this.sharepoint.getGates(params.id);
        this.gates.forEach(async (el, index) => {
          
          el.actions = await this.sharepoint.getActions(el.id);
          
          //set current gate
          if(index < (this.gates.length - 1)) {
            let uncompleted = el.actions.filter(a => !a.completed);
            if(!this.currentGate && uncompleted && (uncompleted.length > 0)) {
              this.setGate(el.id);
            } 
          } else {
            if(!this.currentGate) {
              this.setGate(el.id);
            }
          }

        });
      }
    });
  }

  setGate(gateId: number) {
    let gate = this.gates.find(el => el.id == gateId);
    if(gate) {
      this.currentGate = gate;
      this.currentActions = gate.actions;
      if(this.currentActions.length) {
        let completed = this.currentActions.filter(el => el.completed);
        this.currentGateProgress = Math.round((completed.length / this.currentActions.length) * 10000) / 100;
      } else {
        this.currentGateProgress = 0;
      }
     
    }
  }
}
