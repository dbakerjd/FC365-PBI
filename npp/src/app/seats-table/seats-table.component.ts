import { Component, OnInit } from '@angular/core';
import { AppControlService } from '@services/app/app-control.service';
import { AppDataService } from '@services/app/app-data.service';
import { Opportunity } from '@shared/models/entity';
import { User } from '@shared/models/user';
import { environment } from 'src/environments/environment';

@Component({
  selector: 'app-seats-table',
  templateUrl: './seats-table.component.html',
  styleUrls: ['./seats-table.component.scss']
})
export class SeatsTableComponent implements OnInit {

  constructor(
    private readonly appData: AppDataService,
    private readonly appControl: AppControlService
  ) { }

  generalSeatsCount: {
    TotalSeats: number,
    AssignedSeats: number,
    AvailableSeats: number
  } | null = null;
  generatingSeatsTable = true;
  usersList: User[] = [];
  usersOpportunitiesListItem: { type: string | null, userId: number | null, list: Opportunity[] } = {
    type: null,
    userId: null,
    list: []
  };

  entityNameSingular = 'Opportunity';
  entityNamePlural = 'Opportunities';

  async ngOnInit(): Promise<void> {
    if (environment.isInlineApp) {
      this.entityNameSingular = 'Brand';
      this.entityNamePlural = 'Brands';
    }
    if (this.appControl.isReady) {
      this.loadSeatsInfo();
    } else {
      this.appControl.readySubscriptions.subscribe(val => {
        this.loadSeatsInfo();
      });
    }
  }

  private async loadSeatsInfo() {
    this.generatingSeatsTable = true;
    this.usersList = await this.appData.getUsers();
    this.usersList = this.usersList.filter(el => el.Email);

    for (let index = 0; index < this.usersList.length; index++) {
      const user: any = this.usersList[index];
      const result = await this.appData.getSeats(user.Email);
      if (index == 0 && result) {
        this.generalSeatsCount = {
          AssignedSeats: result?.AssignedSeats,
          TotalSeats: result?.TotalSeats,
          AvailableSeats: result?.AvailableSeats
        }
      }
      user['seats'] = result?.UserGroupsCount;
      const groups = await this.appData.getUserGroups(user.Id);
      const OUgroups = groups.filter(g => g.Title.startsWith('OU-'));
      const OOgroups = groups.filter(g => g.Title.startsWith('OO-'));
      user['opportunities'] = OUgroups.length;
      user['owner'] = OOgroups.length;
    }

    this.generatingSeatsTable = false;
  }

  async listOpportunities(userId: number, group: 'OU' | 'OO') {
    if (this.usersOpportunitiesListItem.type == group && this.usersOpportunitiesListItem.userId == userId) {
      this.usersOpportunitiesListItem.type = null;
      this.usersOpportunitiesListItem.userId = null;
      this.usersOpportunitiesListItem.list = [];
      return;
    }
    const groups = await this.appData.getUserGroups(userId);
    const OUgroups = groups.filter(g => g.Title.startsWith(group + '-'));
    const allOpportunities = await this.appData.getAllOpportunities(false, false);
    const oppsList = OUgroups.map(e => {
      const splittedName = e.Title.split('-');
      return splittedName[1];
    });
    const oppsListRelated = allOpportunities.filter(opp => oppsList.includes(opp.ID.toString()));
    this.usersOpportunitiesListItem.type = group;
    this.usersOpportunitiesListItem.userId = userId;
    this.usersOpportunitiesListItem.list = oppsListRelated;
  }

}
