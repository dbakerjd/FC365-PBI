import { Injectable } from '@angular/core';
import { Opportunity, SharepointService, User } from './sharepoint.service';

@Injectable({
  providedIn: 'root',
})
export class NotificationsService {
  currentUser: User | undefined;

  constructor(private sharepoint: SharepointService) {}

  async opportunityOwnerNotification(opportunity: Opportunity) {
    this.currentUser = await this.getCurrentUser();
    if (this.currentUser.Id !== opportunity.OpportunityOwnerId) {
      await this.sharepoint.createNotification(
        opportunity.OpportunityOwnerId,
        `${this.currentUser.Title} has made you the owner of the opportunity '${opportunity.Title}'`
      );
    }
  }

  async newOpportunityAccessNotification(
    userIds: number[],
    opportunity: Opportunity
  ) {
    this.currentUser = await this.getCurrentUser();
    for (const user of userIds) {
      if (this.currentUser.Id == user) continue;
      await this.sharepoint.createNotification(
        user,
        `${this.currentUser.Title} has given you access to a new opportunity: ${opportunity.Title}`
      );
    }
  }

  async stageAccessNotification(userIds: number[], stageTitle: string, opportunityTitle: string | undefined) {
    this.currentUser = await this.getCurrentUser();
    let notificationMessage = `${this.currentUser.Title} has given you access to '${stageTitle}'`;
    if (opportunityTitle) notificationMessage += `of '${opportunityTitle}' opportunity`;
    for (const user of userIds) {
      if (user == this.currentUser.Id) continue;
      await this.sharepoint.createNotification(user, notificationMessage);
    }
  }

  async modelFolderAccessNotification(userIds: number[], opportunityId: number) {
    this.currentUser = await this.getCurrentUser();
    let notificationMessage = `${this.currentUser.Title} has given you access to Forecast Models`;
    const opportunity = await this.sharepoint.getOpportunity(opportunityId);
    if (opportunity.Title) notificationMessage += `at '${opportunity.Title}' opportunity`;
    for (const user of userIds) {
      if (user == this.currentUser.Id) continue;
      await this.sharepoint.createNotification(user, notificationMessage);
    }
  }

  async folderAccessNotification(userIds: number[], opportunityId: number, departmentId: number) {
    this.currentUser = await this.getCurrentUser();
    let notificationMessage = `${this.currentUser.Title} has given you access to ${departmentId}`;
    const opportunity = await this.sharepoint.getOpportunity(opportunityId);
    if (opportunity.Title) notificationMessage += `at '${opportunity.Title}' opportunity`;
    for (const user of userIds) {
      if (user == this.currentUser.Id) continue;
      await this.sharepoint.createNotification(user, notificationMessage);
    }
  }

  private async getCurrentUser(): Promise<User> {
    if (!this.currentUser) {
      this.currentUser = await this.sharepoint.getCurrentUserInfo();
    }
    return this.currentUser;
  }
}
