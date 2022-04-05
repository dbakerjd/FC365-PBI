import { Injectable } from '@angular/core';
import { Action, Country, EntityGeography, Opportunity, Stage } from '@shared/models/entity';
import { GroupPermission, User } from '@shared/models/user';
import { AppDataService } from './app-data.service';
import * as SPFolders from '@shared/sharepoint/folders';
import * as SPLists from '@shared/sharepoint/list-names';
import { SystemFolder } from '@shared/models/file-system';
import { ToastrService } from 'ngx-toastr';

interface SPGroup {
  Id: number;
  Title: string;
  Description: string;
  LoginName: string;
  OnlyAllowMembersViewMembership: boolean;
}

interface SPGroupListItem {
  type: string;
  data: SPGroup;
}

@Injectable({
  providedIn: 'root'
})
export class PermissionsService {

  constructor(
    // private readonly sharepoint: SharepointService,
    private readonly appData: AppDataService,
    private readonly toastr: ToastrService
  ) { }

  /* set permissions related to working groups a list or item */
  async setPermissions(permissions: GroupPermission[], workingGroups: SPGroupListItem[], itemOrFolder: number | string | null = null) {
    let folders = [SPFolders.FILES_FOLDER, SPFolders.FOLDER_APPROVED, SPFolders.FOLDER_ARCHIVED, SPFolders.FOLDER_WIP];
    for (const gp of permissions) {
      const group = workingGroups.find(gr => gr.type === gp.Title); // get created group involved on the permission
      if (group) {
        if ((folders.indexOf(gp.ListName) != -1) && typeof itemOrFolder == 'string') {
          await this.appData.assignPermissionToFolder(itemOrFolder, group.data.Id, gp.Permission);
        } else {
          if (gp.ListFilter === 'List')
            await this.appData.assignPermissionToList(gp.ListName, group.data.Id, gp.Permission);
          else if (typeof itemOrFolder == 'number')
            await this.appData.assignPermissionToList(gp.ListName, group.data.Id, gp.Permission, itemOrFolder);
        }
      }
    }
  }

  async initializeOpportunity(opportunity: Opportunity, stage: Stage | null): Promise<boolean> {
    const groups = await this.createOpportunityGroups(opportunity.EntityOwnerId, opportunity.ID);
    if (groups.length < 1) return false;

    let permissions;
    // add groups to lists
    permissions = (await this.appData.getGroupPermissions()).filter(el => el.ListFilter === 'List');
    await this.setPermissions(permissions, groups);

    // add groups to the Opportunity
    permissions = await this.appData.getGroupPermissions(SPLists.ENTITIES_LIST_NAME);
    await this.setPermissions(permissions, groups, opportunity.ID);

    // add groups to the Opp geographies
    permissions = await this.appData.getGroupPermissions(SPLists.GEOGRAPHIES_LIST_NAME);
    const oppGeographies = await this.appData.getEntityGeographies(opportunity.ID);
    for (const oppGeo of oppGeographies) {
      await this.setPermissions(permissions, groups, oppGeo.Id);
    }

    if (stage) {
      await this.initializeStage(opportunity, stage, oppGeographies);
    } else {
      await this.initializeInternalEntityFolders(opportunity, oppGeographies);
    }
    
    return true;
  }

  async initializeStage(opportunity: Opportunity, stage: Stage, geographies: EntityGeography[]): Promise<boolean> {
    const OUGroup = await this.appData.createGroup('OU-' + opportunity.ID);
    const OOGroup = await this.appData.createGroup('OO-' + opportunity.ID);
    const SUGroup = await this.appData.createGroup(`SU-${opportunity.ID}-${stage.StageNameId}`);

    if (!OUGroup || !OOGroup || !SUGroup) return false; // something happened with groups

    const owner = await this.appData.getUserInfo(opportunity.EntityOwnerId);
    if (!owner.LoginName) return false;

    if (!await this.appData.addUserToGroupAndSeat(owner, OUGroup.Id, true)) {
      return false;
    }
    await this.appData.addUserToGroup(owner, OOGroup.Id);
    
    let groups: SPGroupListItem[] = [];
    groups.push({ type: 'OU', data: OUGroup });
    groups.push({ type: 'OO', data: OOGroup });
    groups.push({ type: 'SU', data: SUGroup });

    // add groups to the Stage
    let permissions = await this.appData.getGroupPermissions(SPLists.ENTITY_STAGES_LIST_NAME);
    await this.setPermissions(permissions, groups, stage.ID);

    // add stage users to group OU and SU
    let addedSU = [];
    for (const userId of stage.StageUsersId) {
      const user = await this.appData.getUserInfo(userId);
      if (!await this.appData.addUserToGroupAndSeat(user, OUGroup.Id, true)) continue;
      await this.appData.addUserToGroup(user, SUGroup.Id);
      addedSU.push(user.Id);
    }
    if (addedSU.length < 1) {
      // add owner as stage user to don't leave the field blank
      // owner has seat assigned in this point
      await this.appData.addUserToGroup(owner, SUGroup.Id);
      await this.appData.updateStage(stage.ID, { StageUsersId: [owner.Id]});
    } else if (addedSU.length < stage.StageUsersId.length) {
      // update with only the stage users with seat
      await this.appData.updateStage(stage.ID, { StageUsersId: addedSU});
    }

    // Actions
    const stageActions = await this.createStageActions(opportunity, stage);

    // add groups into Actions
    permissions = await this.appData.getGroupPermissions(SPLists.ENTITY_ACTIONS_LIST_NAME);
    for (const action of stageActions) {
      await this.setPermissions(permissions, groups, action.Id);
    }

    // Folders
    const folders = await this.createStageFolders(opportunity, stage, geographies, groups);

    // add groups to folders
    permissions = await this.appData.getGroupPermissions(SPFolders.FILES_FOLDER);
    await this.createFolderGroups(opportunity.ID, permissions, folders, groups);
    return true;
  }

  async initializeInternalEntityFolders(opportunity: Opportunity, geographies: EntityGeography[]) {
    const OUGroup = await this.appData.createGroup('OU-' + opportunity.ID);
    const OOGroup = await this.appData.createGroup('OO-' + opportunity.ID);

    if (!OUGroup || !OOGroup) return false; // something happened with groups

    const owner = await this.appData.getUserInfo(opportunity.EntityOwnerId);
    if (!owner.LoginName) return false;

    if (!await this.appData.addUserToGroupAndSeat(owner, OUGroup.Id, true)) {
      return false;
    }
    await this.appData.addUserToGroup(owner, OOGroup.Id);
    
    let groups: SPGroupListItem[] = [];
    groups.push({ type: 'OU', data: OUGroup });
    groups.push({ type: 'OO', data: OOGroup });

    // Folders
    const folders = await this.createInternalFolders(opportunity, groups, geographies);

    // add groups to folders
    const RefDocsPermissions = await this.appData.getGroupPermissions(SPFolders.FILES_FOLDER);
    await this.createFolderGroups(opportunity.ID, RefDocsPermissions, folders.rw.filter(el => el.DepartmentID), groups);
    const WIPpermissions = await this.appData.getGroupPermissions(SPFolders.FOLDER_WIP);
    await this.createFolderGroups(opportunity.ID, WIPpermissions, folders.rw.filter(el => el.GeographyID), groups);
    const approvedPermissions = await this.appData.getGroupPermissions(SPFolders.FOLDER_APPROVED);
    await this.createFolderGroups(opportunity.ID, approvedPermissions, folders.ro.filter(el => el.ServerRelativeUrl.includes(SPFolders.FOLDER_APPROVED)), groups);
    const archivedPermissions = await this.appData.getGroupPermissions(SPFolders.FOLDER_ARCHIVED);
    await this.createFolderGroups(opportunity.ID, archivedPermissions, folders.ro.filter(el => el.ServerRelativeUrl.includes(SPFolders.FOLDER_ARCHIVED)), groups);
      
    return true;
  }

  /** Creates the DU folder groups and sets permissions for a list of folders 
   * 
   * @param oppId The opportunity ID containing the folders
   * @param permissions List of group permissions to set
   * @param folders List of folders to create the groups
   * @param baseGroups Base of groups to include in the permissions
  */
  private async createFolderGroups(oppId: number, permissions: GroupPermission[], folders: SystemFolder[], baseGroups: SPGroupListItem[]) {
    for (const f of folders) {
      let folderGroups = [...baseGroups]; // copy default groups
      if (f.DepartmentID) {
        let DUGroup = await this.appData.createGroup(`DU-${oppId}-${f.DepartmentID}`, 'Department ID ' + f.DepartmentID);
        if (DUGroup) folderGroups.push({ type: 'DU', data: DUGroup });
      } else if (f.GeographyID) {
        let DUGroup = await this.appData.createGroup(`DU-${oppId}-0-${f.GeographyID}`, 'Geography ID ' + f.GeographyID);
        if (DUGroup) folderGroups.push({ type: 'DU', data: DUGroup });
      }
      await this.setPermissions(permissions, folderGroups, f.ServerRelativeUrl);
    }
  }

  /** Sets the access for the entity departments groups updating their members */
  async updateDepartmentUsers(
    oppId: number,
    stageId: number,
    departmentId: number,
    folderDepartmentId: number,
    geoId: number | null,
    currentUsersList: number[],
    newUsersList: number[]
  ): Promise<boolean> {
    // groups needed
    const OUGroup = await this.appData.getGroup('OU-' + oppId);
    const OOGroup = await this.appData.getGroup('OO-' + oppId);
    let SUGroup = null;
    if(stageId) SUGroup = await this.appData.getGroup('SU-' + oppId + '-' + stageId);
    let groupName = `DU-${oppId}-${departmentId}`;
    let geoCountriesList: Country[] = [];
    if (geoId) {
      groupName += `-${geoId}`;
      geoCountriesList = await this.getCountriesOfEntityGeography(geoId);
    }
    const DUGroup = await this.appData.getGroup(groupName);

    if (!OUGroup || !OOGroup || (!SUGroup && stageId) || !DUGroup) throw new Error("Permission groups missing.");

    const removedUsers = currentUsersList.filter(item => newUsersList.indexOf(item) < 0);
    const addedUsers = newUsersList.filter(item => currentUsersList.indexOf(item) < 0);

    let success = true;
    for (const userId of removedUsers) {
      success = success && await this.appData.removeUserFromGroup(DUGroup.Id, userId);
      if (success && geoId) { // it's model folder
        this.appData.removePowerBI_RLS(oppId, geoCountriesList, userId);
      }
      success = success && await this.removeUserFromAllGroups(oppId, userId, ['OU']); // remove (if needed) of OU group
    }

    if (!success) return success;

    for (const userId of addedUsers) {
      const user = await this.appData.getUserInfo(userId);
      if (!await this.appData.addUserToGroupAndSeat(user, OUGroup.Id, true)) {
        continue;
      }
      success = success && await this.appData.addUserToGroup(user, DUGroup.Id);
      if (success && geoId) { // it's model folder
        this.appData.addPowerBI_RLS(user, oppId, geoCountriesList);
      }
      if (!success) return success;
    }
    return success;
  }

  async changeEntityOwnerPermissions(oppId: number, currentOwnerId: number, newOwnerId: number): Promise<boolean> {
  
    const newOwner = await this.appData.getUserInfo(newOwnerId);
    const OOGroup = await this.appData.getGroup('OO-' + oppId); // Opportunity Owner (OO)
    const OUGroup = await this.appData.getGroup('OU-' + oppId); // Opportunity Users (OO)
    if (!newOwner.LoginName || !OOGroup || !OUGroup) return false;

    let success = await this.removeUserFromAllGroups(oppId, currentOwnerId, ['OO', 'OU']);

    if (success = await this.appData.addUserToGroupAndSeat(newOwner, OUGroup.Id, true) && success) {
      success = await this.appData.addUserToGroup(newOwner, OOGroup.Id) && success;
    }

    return success;
  }

  async changeStageUsersPermissions(oppId: number, masterStageId: number, currentUsers: number[], newUsers: number[]): Promise<boolean> {
    const removedUsers = currentUsers.filter(item => newUsers.indexOf(item) < 0);
    const addedUsers = newUsers.filter(item => currentUsers.indexOf(item) < 0);

    let success = true;
    for (const userId of removedUsers) {
      success = success && await this.removeUserFromAllGroups(oppId, userId, ['SU'], masterStageId.toString());
      success = success && await this.removeUserFromAllGroups(oppId, userId, ['OU']); // remove (if needed) of OU group
    }

    if (!success) return false;

    if (addedUsers.length > 0) {
      const OUGroup = await this.appData.getGroup('OU-' + oppId);
      const SUGroup = await this.appData.getGroup(`SU-${oppId}-${masterStageId}`);
      if (!OUGroup || !SUGroup) return false;

      for (const userId of addedUsers) {
        const user = await this.appData.getUserInfo(userId);
        if (!(success = await this.appData.addUserToGroupAndSeat(user, OUGroup.Id, true) && success)) {
          continue;
        }
        success = success && await this.appData.addUserToGroup(user, SUGroup.Id);
        if (!success) return false;
      }
    }
    return success;
  }

  /** Updates the Entity Geographies with the new sent geographies.
   *  Creates new geographies and soft delete the old ones including their related groups
   */
  async updateEntityGeographies(entity: Opportunity, newGeographies: string[]) {
    const owner = await this.appData.getUserInfo(entity.EntityOwnerId);
    if (!owner.LoginName) throw new Error("Could not determine entity's owner");

    let allGeo: EntityGeography[] = await this.appData.getEntityGeographies(entity.ID, true);

    let neoGeo = newGeographies.filter(el => {
      let arrId = el.split("-");
      let kindOfGeo = arrId[0];
      let id = arrId[1];
      let geo = allGeo.find(el => {
        if (kindOfGeo == 'G') {
          return el.GeographyId == parseInt(id);
        } else {
          return el.CountryId == parseInt(id);
        }
      });

      return !geo;
    });

    let neoCountry = neoGeo.filter(el => {
      let arrId = el.split("-");
      let kindOfGeo = arrId[0];
      return kindOfGeo == 'C';
    }).map(el => {
      let arrId = el.split("-");
      return parseInt(arrId[1]);
    });

    let neoGeography = neoGeo.filter(el => {
      let arrId = el.split("-");
      let kindOfGeo = arrId[0];
      return kindOfGeo == 'G';
    }).map(el => {
      let arrId = el.split("-");
      return parseInt(arrId[1]);
    })

    let restoreGeo: EntityGeography[] = [];
    newGeographies.forEach(el => {
      let arrId = el.split("-");
      let kindOfGeo = arrId[0];
      let id = arrId[1];
      let geo = allGeo.find(el => {
        if (kindOfGeo == 'G') {
          return el.GeographyId == parseInt(id);
        } else {
          return el.CountryId == parseInt(id);
        }
      });

      if (geo && geo.Removed) {
        restoreGeo.push(geo);
      }
    });

    let removeGeo = allGeo.filter(el => {
      let isCountry = !!el.CountryId;
      let geo = newGeographies.find(g => {
        if (isCountry) {
          return g == 'C-' + el.CountryId;
        } else {
          return g == 'G-' + el.GeographyId;
        }
      });

      return !geo && !el.Removed;
    });

    if (removeGeo.length > 0) await this.deleteGeographies(entity, removeGeo);
    if (restoreGeo.length > 0) await this.restoreGeographies(entity, restoreGeo);

    let newGeos: EntityGeography[] = [];
    if (neoGeography.length > 0 || neoCountry.length > 0) {
      newGeos = await this.createGeographies(entity.ID, neoGeography, neoCountry);
    }
    if (newGeos.length < 1) return; // finish

    let OOGroup = await this.appData.getGroup(`OO-${entity.ID}`);
    let OUGroup = await this.appData.getGroup(`OU-${entity.ID}`);
    if (!OOGroup || !OUGroup) throw new Error("Error obtaining user groups.");

    let groups: SPGroupListItem[] = [];
    groups.push({ type: 'OO', data: OOGroup });
    groups.push({ type: 'OU', data: OUGroup });

    let permissions = await this.appData.getGroupPermissions(SPLists.GEOGRAPHIES_LIST_NAME);
    let stages = await this.appData.getEntityStages(entity.ID);
    if (stages && stages.length) {
      for (const oppGeo of newGeos) {
        await this.setPermissions(permissions, groups, oppGeo.Id); // assign permissions to new entity geo items
        for (let index = 0; index < stages.length; index++) {
          let stage = stages[index];
          let stageFolders = await this.appData.getStageFolders(stage.StageNameId, entity.ID, entity.BusinessUnitId);
          let mf = stageFolders.find(el => el.Title == SPFolders.FORECAST_MODELS_FOLDER_NAME);

          if (!mf) throw new Error("Could not find Models folder");

          let folder = await this.appData.createFolder(`/${entity.BusinessUnitId}/${entity.ID}/${stage.StageNameId}/${mf.DepartmentID}/${oppGeo.Id}`);
          if (folder) {
            // department group and Stage Users Group
            const DUGroupName = `DU-${entity.ID}-${mf.DepartmentID}-${oppGeo.Id}`;
            let DUGroup = await this.appData.createGroup(DUGroupName, 'Department ID ' + mf.DepartmentID + ' / Geography ID ' + oppGeo.Id);
            let SUGroup = await this.appData.getGroup(`SU-${entity.ID}-${stage.StageNameId}`);
            if (DUGroup && SUGroup) {
              const permissions = await this.appData.getGroupPermissions(SPFolders.FILES_FOLDER);
              let folderGroups: SPGroupListItem[] = [...groups, { type: 'DU', data: DUGroup }, { type: 'SU', data: SUGroup }];
              await this.setPermissions(permissions, folderGroups, folder.ServerRelativeUrl);
            } else {
              if (!DUGroup) throw new Error("Error creating geography group permissions.");
              else throw new Error("Error getting SU group.");
            }
            await this.appData.createFolder(`/${entity.BusinessUnitId}/${entity.ID}/${stage.StageNameId}/${mf.DepartmentID}/${oppGeo.Id}/0`);
          }
        }
      }
    } else {
      const folders = await this.createInternalFolders(entity, groups, newGeos);

      for (const oppGeo of newGeos) {
        await this.setPermissions(permissions, groups, oppGeo.Id); // assign permissions to new entity geo items
      }
      // add groups to folders
      // (department folders non needed)
      // const departmentPermissions = await this.getGroupPermissions( SPLists.FILES_FOLDER);
      // await this.createFolderGroups(entity.ID, departmentPermissions, folders.rw.filter(el => el.DepartmentID), groups);
      const WIPPermissions = await this.appData.getGroupPermissions(SPFolders.FOLDER_WIP);
      await this.createFolderGroups(entity.ID, WIPPermissions, folders.rw.filter(el => el.GeographyID), groups);
      const approvedPermissions = await this.appData.getGroupPermissions(SPFolders.FOLDER_APPROVED);
      await this.createFolderGroups(entity.ID, approvedPermissions, folders.ro.filter(el => el.ServerRelativeUrl.includes(SPFolders.FOLDER_APPROVED)), groups);
      const archivedPermissions = await this.appData.getGroupPermissions(SPFolders.FOLDER_ARCHIVED);
      await this.createFolderGroups(entity.ID, archivedPermissions, folders.ro.filter(el => el.ServerRelativeUrl.includes(SPFolders.FOLDER_ARCHIVED)), groups);
    }
  }

  async createGeographies(oppId: number, geographies: number[], countries: number[]): Promise<EntityGeography[]> {
    const geographiesList = await this.appData.getGeographiesList();
    const countriesList = await this.appData.getCountriesList();
    let res: EntityGeography[] = [];
    for (const g of geographies) {
      const geoTitle = geographiesList.find(el => el.value == g)?.label;
      if (geoTitle) {
        const newGeo = await this.appData.createEntityGeography({
          Title: geoTitle,
          EntityNameId: oppId,
          GeographyId: g,
          EntityGeographyType: 'Geography'
        });
        res.push(newGeo);
      }
    }
    for (const c of countries) {
      const geoTitle = countriesList.find(el => el.value == c)?.label;
      if (geoTitle) {
        let newGeo: EntityGeography = await this.appData.createEntityGeography({
          Title: geoTitle,
          EntityNameId: oppId,
          CountryId: c,
          EntityGeographyType: 'Country'
        });
        res.push(newGeo);
      }
    }
    return res;
  }

  /** Soft delete entity geographies. Delete DU geography groups related */
  private async deleteGeographies(entity: Opportunity, removeGeos: EntityGeography[]) {
    //removes groups
    let stages = await this.appData.getEntityStages(entity.ID);
    if (stages && stages.length) {
      // external
      for (const geo of removeGeos) {
        for (const stage of stages) {
          let stageFolders = await this.appData.getStageFolders(stage.StageNameId, entity.ID, entity.BusinessUnitId);
          let modelFolders = stageFolders.filter(el => el.containsModels);
          if (modelFolders.length < 1) continue;

          for (const mf of modelFolders) {
            const DUGroupId = await this.appData.getGroupId(`DU-${entity.ID}-${mf.DepartmentID}-${geo.Id}`);
            if (DUGroupId) await this.appData.deleteGroup(DUGroupId);
          }
        }
      }
    } else {
      // internal
      for (const geo of removeGeos) {
        const DUGroupId = await this.appData.getGroupId(`DU-${entity.ID}-0-${geo.Id}`);
        if (DUGroupId) await this.appData.deleteGroup(DUGroupId);
      }
    }

    // soft delete entity geographies
    for (let i = 0; i < removeGeos.length; i++) {
      await this.appData.updateEntityGeography(removeGeos[i].ID, { Removed: "true" });

      // Power BI RLS access 
      const geoCountriesList = await this.getCountriesOfEntityGeography(removeGeos[i].ID);
      await this.appData.removePowerBI_RLS(entity.ID, geoCountriesList);
    }
  }

  /** Restore previously soft deleted entity geographies and create DU groups related */
  private async restoreGeographies(entity: Opportunity, restoreGeos: EntityGeography[]) {
    let OOGroup = await this.appData.getGroup(`OO-${entity.ID}`);
    let OUGroup = await this.appData.getGroup(`OU-${entity.ID}`);
    if (!OOGroup || !OUGroup) throw new Error("Error obtaining user groups.");

    let groups: SPGroupListItem[] = [];
    groups.push({ type: 'OO', data: OOGroup });
    groups.push({ type: 'OU', data: OUGroup });

    let stages = await this.appData.getEntityStages(entity.ID);
    if (stages && stages.length) {
      // external
      for (const geo of restoreGeos) {
        for (const stage of stages) {
          let stageFolders = await this.appData.getStageFolders(stage.StageNameId, entity.ID, entity.BusinessUnitId);
          let modelFolders = stageFolders.filter(el => el.containsModels);
          if (modelFolders.length < 1) continue;

          // not needed because SU group is never removed
          // let SUGroup = await this.createGroup(`SU-${entity.ID}-${stage.StageNameId}`);
          // if (!SUGroup) throw new Error('Error obtaining user group (SU)');

          const permissions = await this.appData.getGroupPermissions(SPFolders.FILES_FOLDER);
          for (const mf of modelFolders) {
            const folder = await this.appData.getFolder(SPFolders.FILES_FOLDER + `/${entity.BusinessUnitId}/${entity.ID}/${stage.StageNameId}/${mf.DepartmentID}/${geo.Id}`);
            const DUGroupName = `DU-${entity.ID}-${mf.DepartmentID}-${geo.Id}`;
            let DUGroup = await this.appData.createGroup(DUGroupName, 'Department ID ' + mf.DepartmentID + ' / Geography ID ' + geo.Id);
            if (folder && DUGroup) {
              groups.push({ type: 'DU', data: DUGroup });
              await this.createFolderGroups(entity.ID, permissions, [folder], groups);
            }
          }
        }
      }
    } else {
      // internal
      const folders = await this.createInternalFolders(entity, groups, restoreGeos);

      const WIPPermissions = await this.appData.getGroupPermissions(SPFolders.FOLDER_WIP);
      await this.createFolderGroups(entity.ID, WIPPermissions, folders.rw.filter(el => el.GeographyID), groups);
      const approvedPermissions = await this.appData.getGroupPermissions(SPFolders.FOLDER_APPROVED);
      await this.createFolderGroups(entity.ID, approvedPermissions, folders.ro.filter(el => el.ServerRelativeUrl.includes(SPFolders.FOLDER_APPROVED)), groups);
      const archivedPermissions = await this.appData.getGroupPermissions(SPFolders.FOLDER_ARCHIVED);
      await this.createFolderGroups(entity.ID, archivedPermissions, folders.ro.filter(el => el.ServerRelativeUrl.includes(SPFolders.FOLDER_ARCHIVED)), groups);   
    }

    // restore entity geographies
    for (let i = 0; i < restoreGeos.length; i++) {
      await this.appData.updateEntityGeography(restoreGeos[i].ID, { Removed: "false" });
    }
  }

  private async createInternalFolders(entity: Opportunity, groups: SPGroupListItem[], geographies?: EntityGeography[]): Promise<{rw: SystemFolder[], ro: SystemFolder[]}> {
    let ReadWriteNames = [SPFolders.FOLDER_WIP, SPFolders.FOLDER_DOCUMENTS];
    let ReadOnlyNames = [SPFolders.FOLDER_APPROVED, SPFolders.FOLDER_ARCHIVED];

    const OUGroup = groups.find(el => el.type == "OU");
    if (!OUGroup) throw new Error("Error creating group permissions for internal folders.");
    
    if(!geographies) {
      geographies = await this.appData.getEntityGeographies(entity.ID);
    }

    let rwFolders: SystemFolder[] = [];
    for (const mf of ReadWriteNames) {
      const mfFolder = await this.appData.createFolder(`${mf}`, true);
      if(mfFolder) {
        const BUFolder = await this.appData.createFolder(`${mf}/${entity.BusinessUnitId}`, true);
        if(BUFolder) {
          const folder = await this.appData.createFolder(`${mf}/${entity.BusinessUnitId}/${entity.ID}`, true);
          if (folder) {
            await this.appData.assignReadPermissionToFolder(folder.ServerRelativeUrl, OUGroup.data.Id);
            const emptyStageFolder = await this.appData.createFolder(`${mf}/${entity.BusinessUnitId}/${entity.ID}/0`, true);
            if(emptyStageFolder) {
              await this.appData.assignReadPermissionToFolder(emptyStageFolder.ServerRelativeUrl, OUGroup.data.Id);
              if(mf != SPFolders.FOLDER_DOCUMENTS) {
                const forecastFolder = await this.appData.createFolder(`${mf}/${entity.BusinessUnitId}/${entity.ID}/0/0`, true);
                if(forecastFolder) {
                  rwFolders = rwFolders.concat(await this.createEntityGeographyFolders(entity, geographies, mf));
                }
              } else {
                rwFolders = rwFolders.concat(await this.createDepartmentFolders(entity, mf));
              } 
            }
          }
        }
      }
    }

    let roFolders: SystemFolder[] = [];
    for (const mf of ReadOnlyNames) {
      const mfFolder = await this.appData.createFolder(`${mf}`, true);
      if(mfFolder) {
        const BUFolder = await this.appData.createFolder(`${mf}/${entity.BusinessUnitId}`, true);
        if(BUFolder) {
          const folder = await this.appData.createFolder(`${mf}/${entity.BusinessUnitId}/${entity.ID}`, true);
          if (folder) {
            await this.appData.assignReadPermissionToFolder(folder.ServerRelativeUrl, OUGroup.data.Id);
            const emptyStageFolder = await this.appData.createFolder(`${mf}/${entity.BusinessUnitId}/${entity.ID}/0`, true);
            if(emptyStageFolder) {
              await this.appData.assignReadPermissionToFolder(emptyStageFolder.ServerRelativeUrl, OUGroup.data.Id);
              const forecastFolder = await this.appData.createFolder(`${mf}/${entity.BusinessUnitId}/${entity.ID}/0/0`, true);
              if(forecastFolder) {  
                roFolders = roFolders.concat(await this.createEntityGeographyFolders(entity, geographies, mf));
              }
            }
          }
        }
      }
    }
    return {
      rw: rwFolders,
      ro: roFolders
    };
  }

  private async createEntityGeographyFolders(entity: Opportunity, geographies: EntityGeography[], mf: string, departmentId: number = 0, cycleId: number = 0): Promise<SystemFolder[]> {
    let folders: SystemFolder[] = [];
    let basePath = `${mf}/${entity.BusinessUnitId}/${entity.ID}/0/${departmentId}`;
    for (const geo of geographies) {
      let geoFolder = await this.appData.createFolder(`${basePath}/${geo.ID}`, true);
      if (geoFolder) {
        geoFolder.GeographyID = geo.ID;
        geoFolder.DepartmentID = departmentId;
        folders.push(geoFolder);
        await this.appData.createFolder(`${basePath}/${geo.ID}/${cycleId}`, true);
      }
    }
    
    return folders;
  }

  private async createDepartmentFolders(entity: Opportunity, mf: string): Promise<SystemFolder[]> {
    let folders: SystemFolder[] = [];
    let basePath = `${mf}/${entity.BusinessUnitId}/${entity.ID}/0`;
    let departmentFolders = await this.appData.getInternalDepartments();
    for(const dept of departmentFolders) {
      let folder = await this.appData.createFolder(`${basePath}/${dept.DepartmentID}`, true);
      if (folder) {
        folder.DepartmentID = dept.DepartmentID;
        folders.push(folder);
        folder = await this.appData.createFolder(`${basePath}/${dept.DepartmentID}/0`, true);
        if(folder) {
          folder = await this.appData.createFolder(`${basePath}/${dept.DepartmentID}/0/0`, true);
        }
      }
    }
    return folders;
  }

  private async createOpportunityGroups(ownerId: number, oppId: number): Promise<SPGroupListItem[]> {
    let group;
    let groups: SPGroupListItem[] = [];
    const owner = await this.appData.getUserInfo(ownerId);
    if (!owner.LoginName) return [];

    // Opportunity Users (OU)
    group = await this.appData.createGroup(`OU-${oppId}`);
    if (group) {
      groups.push({ type: 'OU', data: group });
      if (!await this.appData.addUserToGroupAndSeat(owner, group.Id, true)) {
        return [];
      }
    }

    // Opportunity Owner (OO)
    group = await this.appData.createGroup(`OO-${oppId}`);
    if (group) {
      groups.push({ type: 'OO', data: group });
      await this.appData.addUserToGroupAndSeat(owner, group.Id);
    }

    return groups;
  }

  private async createStageFolders(opportunity: Opportunity, stage: Stage, geographies: EntityGeography[], groups: SPGroupListItem[]): Promise<SystemFolder[]> {

    const OUGroup = groups.find(el => el.type == "OU");
    if (!OUGroup) throw new Error("Error creating group permissions.");

    const masterFolders = await this.appData.getStageFolders(stage.StageNameId);
    const buFolder = await this.appData.createFolder(`/${opportunity.BusinessUnitId}`);
    const oppFolder = await this.appData.createFolder(`/${opportunity.BusinessUnitId}/${stage.EntityNameId}`);
    const stageFolder = await this.appData.createFolder(`/${opportunity.BusinessUnitId}/${stage.EntityNameId}/${stage.StageNameId}`);
    if (!oppFolder || !stageFolder) throw new Error("Error creating opportunity folders.");

    // assign OU to parent folders
    await this.appData.assignPermissionToFolder(oppFolder.ServerRelativeUrl, OUGroup.data.Id, 'ListRead');
    await this.appData.assignReadPermissionToFolder(stageFolder.ServerRelativeUrl, OUGroup.data.Id);

    let folders: SystemFolder[] = [];

    for (const mf of masterFolders) {
      let folder = await this.appData.createFolder(`/${opportunity.BusinessUnitId}/${stage.EntityNameId}/${stage.StageNameId}/${mf.DepartmentID}`);
      if (folder) {
        if (mf.DepartmentID) {
          folder.DepartmentID = mf.DepartmentID;
          folders.push(folder);
          folder = await this.appData.createFolder(`/${opportunity.BusinessUnitId}/${stage.EntityNameId}/${stage.StageNameId}/${mf.DepartmentID}/0`);
          if (folder) {
            folder = await this.appData.createFolder(`/${opportunity.BusinessUnitId}/${stage.EntityNameId}/${stage.StageNameId}/${mf.DepartmentID}/0/0`);
          }
        } else {
          for (let geo of geographies) {
            let folder = await this.appData.createFolder(`/${opportunity.BusinessUnitId}/${stage.EntityNameId}/${stage.StageNameId}/${mf.DepartmentID}/${geo.Id}`);
            if (folder) {
              folder.DepartmentID = 0;
              folder.GeographyID = geo.ID;
              folders.push(folder);
              folder = await this.appData.createFolder(`/${opportunity.BusinessUnitId}/${stage.EntityNameId}/${stage.StageNameId}/${mf.DepartmentID}/${geo.Id}/0`);
            }
          }
        }
      }
    }

    return folders;
  }

  private async createStageActions(opportunity: Opportunity, stage: Stage): Promise<Action[]> {
    const masterActions = await this.appData.getMasterActions(stage.StageNameId, opportunity.OpportunityTypeId);

    let actions: Action[] = [];
    for (const ma of masterActions) {
      const a = await this.appData.createStageActionFromMaster(ma, opportunity.ID);
      if (a.Id) actions.push(a);
    }
    return actions;
  }

  private async removeUserFromAllGroups(oppId: number, userId: number, groups: string[], sufix: string = ''): Promise<boolean> {
    const userGroups = await this.appData.getUserGroups(userId);
    const involvedGroups = userGroups.filter(userGroup => {
      for (const groupType of groups) {
        if (userGroup.Title.startsWith(groupType + '-' + oppId + (sufix ? '-' + sufix : ''))) return true;
      }
      return false;
    });
    let success = true;
    for (const ig of involvedGroups) {
      if (!ig.Title.startsWith('OU')) success = await this.appData.removeUserFromGroup(ig.Title, userId) && success;
    }

    if (!success) return false;

    // has to be removed of OU -> extra check if the user is not in any opportunity group
    if (involvedGroups.some(ig => ig.Title.startsWith('OU'))) {
      const updatedGroups = await this.appData.getUserGroups(userId);
      if (updatedGroups.filter(userGroup => userGroup.Title.split('-')[1] === oppId.toString()).length === 1) {
        // not involved in any group of the opportunity
        success = await this.appData.removeUserFromGroup('OU-' + oppId, userId, true);
      }
    }
    return success;
  }

  /** Returns the entire list of countries related to Entity Geography */
  private async getCountriesOfEntityGeography(geoId: number): Promise<Country[]> {
    const countryExpandOptions = '$select=*,Country/ID,Country/Title&$expand=Country';
    const entityGeography = await this.appData.getEntityGeography(geoId);
    if (entityGeography.CountryId && entityGeography.Country) {
      return [entityGeography.Country];
    }
    else if (entityGeography.GeographyId) {
      const masterGeography = await this.appData.getMasterGeography(entityGeography.GeographyId);
      return masterGeography.Country;
    }
    return [];
  }
}
