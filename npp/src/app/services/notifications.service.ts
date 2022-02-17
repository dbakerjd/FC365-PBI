import { Injectable } from '@angular/core';
import { NPPNotification, Opportunity, SharepointService, User } from './sharepoint.service';

@Injectable({
  providedIn: 'root',
})
export class NotificationsService {
  currentUser: User | undefined;

  constructor(private sharepoint: SharepointService) {}


  async getNotifications(): Promise<NPPNotification[]> {
    const currentUser = await this.getCurrentUser();
    const limit = 15;
    const fromDate = new Date();
    fromDate.setMonth(fromDate.getMonth() - 1);
    return await this.sharepoint.getUserNotifications(currentUser.Id, fromDate, limit);
  }

  async getUnreadNotifications(): Promise<number> {
    const currentUser = await this.getCurrentUser();
    return await this.sharepoint.notificationsCount(currentUser.Id, 'ReadAt eq null');
  }

  async updateUnreadNotifications() {
    const currentUser = await this.getCurrentUser();
    const unreadNotifications = await this.sharepoint.getUserNotifications(currentUser.Id, false);
    for (const not of unreadNotifications) {
      await this.sharepoint.updateNotification(not.Id, { ReadAt: new Date() })
    }
  }

  async opportunityOwnerNotification(opportunity: Opportunity) {
    const currentUser = await this.getCurrentUser();
    if (currentUser.Id !== opportunity.EntityOwnerId) {
      await this.sharepoint.createNotification(
        opportunity.EntityOwnerId,
        `${currentUser.Title} has made you the owner of the opportunity '${opportunity.Title}'`
      );
    }
  }

  async newOpportunityAccessNotification(
    userIds: number[],
    opportunity: Opportunity
  ) {
    const currentUser = await this.getCurrentUser();
    for (const user of userIds) {
      if (currentUser.Id == user) continue;
      await this.sharepoint.createNotification(
        user,
        `${currentUser.Title} has given you access to a new opportunity: ${opportunity.Title}`
      );
    }
  }

  async stageAccessNotification(userIds: number[], stageTitle: string, opportunityTitle: string | undefined) {
    const currentUser = await this.getCurrentUser();
    let notificationMessage = `${currentUser.Title} has given you access to '${stageTitle}'`;
    if (opportunityTitle) notificationMessage += ` of '${opportunityTitle}' opportunity`;
    for (const user of userIds) {
      if (user == currentUser.Id) continue;
      await this.sharepoint.createNotification(user, notificationMessage);
    }
  }

  async modelFolderAccessNotification(userIds: number[], opportunityId: number) {
    const currentUser = await this.getCurrentUser();
    let notificationMessage = `${currentUser.Title} has given you access to Forecast Models`;
    const opportunity = await this.sharepoint.getOpportunity(opportunityId);
    if (opportunity.Title) notificationMessage += ` at '${opportunity.Title}' opportunity`;
    for (const user of userIds) {
      if (user == currentUser.Id) continue;
      await this.sharepoint.createNotification(user, notificationMessage);
    }
  }

  async folderAccessNotification(userIds: number[], opportunityId: number, departmentId: number) {
    const currentUser = await this.getCurrentUser();
    const folder = await this.sharepoint.getNPPFolderByDepartment(departmentId);
    if (!folder) return;
    let notificationMessage = `${currentUser.Title} has given you access to ${folder.Title}`;
    const opportunity = await this.sharepoint.getOpportunity(opportunityId);
    if (opportunity.Title) notificationMessage += ` at '${opportunity.Title}' opportunity`;
    for (const user of userIds) {
      if (user == currentUser.Id) continue;
      await this.sharepoint.createNotification(user, notificationMessage);
    }
  }

  async modelSubmittedNotification(fileName: string, opportunityId: number, usersGroups: string[]) {
    const currentUser = await this.getCurrentUser();
    await this.generateModelNotification(
      `${currentUser.Title} has submitted for approval the model '${fileName}'`, 
      usersGroups,
      opportunityId
    );
  }

  async modelApprovedNotification(fileName: string, opportunityId: number, usersGroups: string[]) {
    const currentUser = await this.getCurrentUser();
    await this.generateModelNotification(
      `${currentUser.Title} has approved the model '${fileName}'`, 
      usersGroups,
      opportunityId
    );
  }

  async modelRejectedNotification(fileName: string, opportunityId: number, usersGroups: string[]) {
    const currentUser = await this.getCurrentUser();
    await this.generateModelNotification(
      `${currentUser.Title} has rejected the model '${fileName}'`, 
      usersGroups,
      opportunityId
    );
  }

  async modelNewScenarioNotification(fileName: string, opportunityId: number, usersGroups: string[]) {
    const currentUser = await this.getCurrentUser();
    await this.generateModelNotification(
      `${currentUser.Title} has created a new scenario from '${fileName}'`, 
      usersGroups,
      opportunityId
    );
  }

  private async generateModelNotification(notificationMessage: string, usersGroups: string[], opportunityId: number | null = null) {
    // get unique users involved
    const currentUser = await this.getCurrentUser();
    let users: User[] = [];
    for (const group of usersGroups) {
      users = users.concat(await this.sharepoint.getGroupMembers(group));
    }
    const uniqueUsers = [...new Map(users.map(u => [u.Id, u])).values()].filter((u: User) => u.Id != currentUser.Id);

    if (users.length < 1) return;

    if (opportunityId) {
      const opportunity = await this.sharepoint.getOpportunity(opportunityId);
      if (opportunity.Title) notificationMessage += ` at '${opportunity.Title}' opportunity`;
    }
    
    // create notifications to involved users
    for (const u of uniqueUsers) {
      await this.sharepoint.createNotification(u.Id, notificationMessage);
    }
  }

  private async getCurrentUser(): Promise<User> {
    if (!this.currentUser) {
      this.currentUser = await this.sharepoint.getCurrentUserInfo();
    }
    return this.currentUser;
  }
}
