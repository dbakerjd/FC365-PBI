import { Injectable } from '@angular/core';
import { EntityGeography, Opportunity } from '@shared/models/entity';
import { NPPFile } from '@shared/models/file-system';
import { FILES_FOLDER, FOLDER_APPROVED, FOLDER_ARCHIVED, FOLDER_DOCUMENTS, FOLDER_POWER_BI_APPROVED, FOLDER_POWER_BI_ARCHIVED, FOLDER_POWER_BI_DOCUMENTS, FOLDER_POWER_BI_WIP, FOLDER_WIP } from '@shared/sharepoint/folders';
import { AppDataService } from './app-data.service';

@Injectable({
  providedIn: 'root'
})
export class FilesService {

  constructor(private readonly appData: AppDataService) { }

  async uploadFileToFolder(fileData: string, folder: string, fileName: string, metadata?: any): Promise<any> {
    if (metadata) {
      let scenarios = metadata.ModelScenarioId;
      if (scenarios) {
        let file = await this.getFileByScenarios(folder, scenarios);
        if (file) this.appData.deleteFile(file?.ServerRelativeUrl);
      }
    }

    let uploaded: any = await this.appData.uploadFile(fileData, folder, fileName);

    if (metadata && uploaded.ListItemAllFields?.ID/* && uploaded.ServerRelativeUrl*/) {

      // GetFileByServerRelativeUrl('/Folder Name/{file_name}')/CheckOut()
      // GetFileByServerRelativeUrl('/Folder Name/{file_name}')/CheckIn(comment='Comment',checkintype=0)
      let arrFolder = folder.split("/");
      let rootFolder = arrFolder[0];
      if (!metadata.Comments) {
        metadata.Comments = " ";
      }
      await this.appData.updateFilePropertiesById(uploaded.ListItemAllFields.ID, rootFolder, metadata);
    }
    return uploaded;
  }

  /** Removes the file and related CSVs, if needed */
  async deleteFile(fileUri: string, checkCSV: boolean = true): Promise<boolean> {
    //First check if it has related CSV files to remove
    if (checkCSV) {
      await this.deleteRelatedCSV(fileUri);
    }
    //then remove
    return await this.appData.deleteFile(fileUri);
  }

  /** Adds a comments related to file */
  async addFileComment(file: NPPFile, str: string) {
    let comments = file.ListItemAllFields?.Comments?.replace(/""/g, '"');
    let parsedComments = [];
    let commentsStr = "";
    if(comments) {
      try {
        parsedComments = JSON.parse(comments);
      } catch(e) {

      }
      let currentUser = await this.appData.getCurrentUserInfo();
      let newComment = {
        text: str,
        email: currentUser.Email,
        name: currentUser.Title?.indexOf("@") == -1 ? currentUser.Title : currentUser.Email,
        userId: currentUser.Id,
        createdAt: new Date().toISOString()
      }
      parsedComments.push(newComment);
      commentsStr = JSON.stringify(parsedComments)
      if(file.ListItemAllFields) file.ListItemAllFields.Comments = commentsStr;
    }
    return commentsStr;   
  }

  /** Search for a model with scenarios assigned */
  async getFileByScenarios(path: string, scenarios: number[]) {
    let files = await this.appData.getFolderFiles(path, false);
    for (let i = 0; i < files.length; i++) {
      let model = files[i];
      let sameScenario = this.haveSameScenarios(model, scenarios);
      if (sameScenario) {
        return model;
      }
    }
    return null;
  }

  /** Set the selected status to all models of an entity folder */
  async setAllEntityModelsStatusInFolder(entity: Opportunity, folder: string, status: string) {

    const geographies = await this.appData.getEntityGeographies(entity.ID); // 1 = stage id would be dynamic in the future

    let arrFolder = folder.split("/");
    let rootFolder = arrFolder[0];

    for (let i = 0; i < geographies.length; i++) {
      let geo = geographies[i];
      let files = await this.appData.getFolderFiles(folder + "/" + geo.ID + "/0", true);
      for (let j = 0; files && j < files.length; j++) {
        let model = files[j];
        await this.setFileApprovalStatus(rootFolder, model, entity, status);
      }
    }

  }

  async setFileApprovalStatus(rootFolder: string, file: NPPFile, entity: Opportunity | null, status: string, comments: string | null = null) {
    if (file.ListItemAllFields) {
      const statusId = await this.appData.getMasterApprovalStatusId(status);
      if (!statusId) {
        throw new Error("File status not found");
      };
      /*TODO use something like this to ensure unique name
      while (await this.sharepoint.existsFile(fileName, destinationFolder) && ++attemps < 11) {
        fileName = baseFileName + '-copy-' + attemps + '.' + extension;
      }*/
      let data = { ApprovalStatusId: statusId };
      if (comments) Object.assign(data, { Comments: comments });

      await this.appData.updateFilePropertiesById(file.ListItemAllFields.ID, rootFolder, data);
      let res;
      if (status === "Approved" && entity && file.ServerRelativeUrl.indexOf(FILES_FOLDER) == -1) {
        let arrFolder = file.ServerRelativeUrl.split("/");
        await this.removeNPPOldApprovedModel(entity, file);
        res = await this.appData.copyFile(file.ServerRelativeUrl, '/' + arrFolder[1] + '/' + arrFolder[2] + '/' + FOLDER_APPROVED + '/' + entity.BusinessUnitId + '/' + entity.ID + '/0/0/' + arrFolder[arrFolder.length - 3] + '/0/', file.Name);

        if (res) {
          await this.appData.updateFilePropertiesByPath(res, { OriginalModelId: file.ListItemAllFields.ID })
          await this.copyCSV(file, res);
        }
        return res;
      };

      return true;
    } else {
      throw new Error("Missing file metadata.");
    }
  }

  /** TOCHECK similud amb setentityapprovalstatus */
  /** TODEL */
  // async setBrandApprovalStatus(rootFolder: string, file: NPPFile, brand: Opportunity | null, status: string, comments: string | null = null) {
  //   if(file.ListItemAllFields) {
  //     const statusId = await this.getMasterApprovalStatusId(status);
  //     if (!statusId) return false;
  //     /*TODO use something like this to ensure unique name
  //     while (await this.sharepoint.existsFile(fileName, destinationFolder) && ++attemps < 11) {
  //       fileName = baseFileName + '-copy-' + attemps + '.' + extension;
  //     }*/
  //     let data = { ApprovalStatusId: statusId };
  //     if (comments) Object.assign(data, { Comments: comments });
  
  //     await this.sharepoint.updateItem(file.ListItemAllFields.ID, `lists/getbytitle('${rootFolder}')`, data);
  //     let res;
  //     if(status === "Approved" && brand) {
  //       let arrFolder = file.ServerRelativeUrl.split("/");
  //       await this.removeOldAcceptedModel(brand, file);
  //       res = await this.appData.copyFile(file.ServerRelativeUrl, '/'+arrFolder[1]+'/'+arrFolder[2]+'/'+FOLDER_APPROVED+'/'+brand.BusinessUnitId+'/'+brand.ID+'/0/0/'+arrFolder[arrFolder.length - 3]+'/0/', file.Name);
  //       return res;
  //     };
      
  //     return true;
  //   } else {
  //     throw new Error("Missing file metadata.");
  //   }
  // }

  async addScenarioSufixToFilename(originFilename: string, scenarioId: number): Promise<string | false> {
    const scenarios = await this.appData.getScenariosList();
    const extension = originFilename.split('.').pop();
    if (!extension) return false;

    const baseFileName = originFilename.substring(0, originFilename.length - (extension.length + 1));
    return baseFileName
      + '-' + scenarios.find(el => el.value === scenarioId)?.label.replace(/ /g, '').toLocaleLowerCase()
      + '.' + extension;
  }

  async getBrandFolderFilesCount(brand: Opportunity, folder: string) {
    let currentFolder = folder+'/'+brand.BusinessUnitId+'/'+brand.ID+'/0/0';
    const geoFolders = await this.appData.getSubfolders(currentFolder);
    let currentFiles = [];
    for (const geofolder of geoFolders) {
      let folder = currentFolder + '/' + geofolder.Name+'/0';
      currentFiles.push(...await this.appData.getFolderFiles(folder, true));
    }
    return currentFiles.length;
  }

  async moveAllFolderFiles(origin: string, dest: string, moveCSVs: boolean = true) {
    let files = await this.appData.getFolderFiles(origin);
    for(let i=0;i<files.length; i++){
      let model = files[i];
      let path = await this.appData.moveFile(model.ServerRelativeUrl, dest);
      if(moveCSVs) {
        await this.moveCSV(model, path);
      }
    }
  }

  async deleteRelatedCSV(url: string) {
    let metadata = await this.appData.getFileProperties(url);
    let csvFiles = await this.getModelCSVFiles({ ServerRelativeUrl: url, ListItemAllFields: metadata } as NPPFile);
    for(let i = 0; i < csvFiles.length; i++) {
      this.deleteFile(csvFiles[i].ServerRelativeUrl, false);
    } 
  }

  /** Copy files of one external opportunity to an internal one */
  async copyFilesExternalToInternal(extOppId: number, intOppId: number) {
    const externalEntity = await this.appData.getEntity(extOppId);
    const internalEntity = await this.appData.getEntity(intOppId);

    // copy models
    // [TODO] search for last stage number (now 3, but could change?)
    const externalModelsFolder =  FILES_FOLDER + `/${externalEntity.BusinessUnitId}/${externalEntity.ID}/3/0`;
    const internalModelsFolder = FOLDER_WIP + `/${internalEntity.BusinessUnitId}/${internalEntity.ID}/0/0`;
    const externalGeographies = await this.appData.getEntityGeographies(externalEntity.ID);
    const internalGeographies = await this.appData.getEntityGeographies(internalEntity.ID);
    for (const extGeo of externalGeographies) {
      const intGeo = internalGeographies.find((g: EntityGeography) => {
        if (g.EntityGeographyType == 'Geography') return extGeo.GeographyId === g.GeographyId;
        else if (g.EntityGeographyType == 'Country') return extGeo.CountryId === g.CountryId;
        else return false;
      });

      if (intGeo) {
        await this.copyAllFolderFiles(`${externalModelsFolder}/${extGeo.Id}/0/`, `${internalModelsFolder}/${intGeo.Id}/0/`);
      }
    }
  }

  private async copyAllFolderFiles(origin: string, dest: string, copyCSVs: boolean = true) {
    let files = await this.appData.getFolderFiles(origin);
    for(let i=0;i<files.length; i++){
      let model = files[i];
      let path = await this.appData.copyFile(model.ServerRelativeUrl, dest, model.Name);
      if(copyCSVs) {
        let arrUrl = model.ServerRelativeUrl.split("/"); // server relative url base for path
        await this.copyCSV(model, "/"+arrUrl[1]+"/"+arrUrl[2]+"/"+path);
      }
    }
  }

  private async removeOldAcceptedModel(brand: Opportunity, file: NPPFile) {
    if(file.ListItemAllFields && file.ListItemAllFields.ModelScenarioId) {
      let arrFolder = file.ServerRelativeUrl.split("/");
      let path = '/'+arrFolder[1]+'/'+arrFolder[2]+'/'+FOLDER_APPROVED+'/'+brand.BusinessUnitId+'/'+brand.ID+'/0/0/'+arrFolder[arrFolder.length - 3]+'/0/';
      let scenarios = file.ListItemAllFields.ModelScenarioId;

      let model = await this.getFileByScenarios(path, scenarios);
      if(model) {
        await this.deleteFile(model.ServerRelativeUrl);
      }
    }
  }

  private async removeNPPOldApprovedModel(entity: Opportunity, file: NPPFile) {
    if (file.ListItemAllFields && file.ListItemAllFields.ModelScenarioId) {
      const arrFolder = file.ServerRelativeUrl.split("/");
      const path = '/' + arrFolder[1] + '/' + arrFolder[2] + '/' + FOLDER_APPROVED + '/' + entity.BusinessUnitId + '/' + entity.ID + '/0/0/' + arrFolder[arrFolder.length - 3] + '/0/';
      const scenarios = file.ListItemAllFields.ModelScenarioId;

      const model = await this.getFileByScenarios(path, scenarios);
      if (model) {
        await this.deleteFile(model.ServerRelativeUrl);
      }
    }
  }

  /** Check if the model have exactly the same scenarios */
  private haveSameScenarios(model: NPPFile, scenarios: number[]): boolean {
    if (model.ListItemAllFields && model.ListItemAllFields.ModelScenarioId) {

      let sameScenario = model.ListItemAllFields.ModelScenarioId.length === scenarios.length;

      for (let j = 0; sameScenario && j < model.ListItemAllFields.ModelScenarioId.length; j++) {
        let scenarioId = model.ListItemAllFields.ModelScenarioId[j];
        sameScenario = sameScenario && (scenarios.indexOf(scenarioId) != -1);
      }

      return sameScenario;

    } else return false;
  }

  private async copyCSV(file: NPPFile, path: string) {
    if (file.ListItemAllFields) {
      let arrFolder = file.ServerRelativeUrl.split("/");
      let destLibrary = this.getPowerBICSVRootPathFromModelPath(path);
  
      let csvFiles = await this.getModelCSVFiles(file);
      let destModel = await this.appData.getFileProperties(path);
  
      for(let i = 0; i < csvFiles.length; i++) {
        let tmpFile = csvFiles[i];
        let newFileName = tmpFile.Name.replace('_'+file.ListItemAllFields.ID+'.', '_'+destModel.ID+'.');
        let newPath = '/'+arrFolder[1]+'/'+arrFolder[2]+'/'+destLibrary+'/';
        await this.appData.copyFile(tmpFile.ServerRelativeUrl, newPath, newFileName);
        await this.appData.updateFilePropertiesByPath(newPath+newFileName, {ForecastId: destModel.ID});
      } 
    }
  }

  private async moveCSV(file: NPPFile, path: string) {
    if (file.ListItemAllFields) {
      let arrFolder = file.ServerRelativeUrl.split("/");
      let destLibrary = this.getPowerBICSVRootPathFromModelPath(path);
  
      let csvFiles = await this.getModelCSVFiles(file);
      let destModel = await this.appData.getFileProperties(path);
  
      for(let i = 0; i < csvFiles.length; i++) {
        let tmpFile = csvFiles[i];
        let newFileName = tmpFile.Name.replace('_'+file.ListItemAllFields.ID+'.', '_'+destModel.ID+'.');
        let newPath = destLibrary+'';
        await this.appData.moveFile(tmpFile.ServerRelativeUrl, newPath, newFileName);
        await this.appData.updateFilePropertiesByPath("/"+arrFolder[1]+"/"+arrFolder[2]+"/"+newPath+"/"+newFileName, {ForecastId: destModel.ID});
      } 
    }
  }

  private async getModelCSVFiles(file: NPPFile) {
    let powerBiLibrary = this.getPowerBICSVRootPathFromModelPath(file.ServerRelativeUrl);
    let files: NPPFile[] = []

    if (powerBiLibrary && file.ListItemAllFields) {
      const result = await this.appData.getFileByForecast(powerBiLibrary, file.ListItemAllFields.ID);
      if (result.value) {
        files = result.value;
      }   
    }

    return files;
  }

  /** Clone a forecast model to a new file with new scenarios */
  async cloneForecastModel(originFile: NPPFile, newFilename: string, newScenarios: number[], authorId: number, comments = ''): Promise<boolean> {

    const destinationFolder = originFile.ServerRelativeUrl.replace('/' + originFile.Name, '/');

    let success = await this.appData.cloneFile(originFile.ServerRelativeUrl, destinationFolder, newFilename);
    if (!success) return false;

    let newFileInfo = await this.appData.getFileByName(destinationFolder, newFilename);

    if (newFileInfo.value[0].ListItemAllFields && originFile.ListItemAllFields) {
      const newData:any = {
        ModelScenarioId: newScenarios,
        Comments: comments ? comments : null,
        ApprovalStatusId: await this.appData.getMasterApprovalStatusId("In Progress")
      }
      
      let arrFolder = destinationFolder.split("/");
      let rootFolder = arrFolder[3];
      
      success = await this.appData.updateFilePropertiesById(newFileInfo.value[0].ListItemAllFields.ID, rootFolder, newData);
      // TOCHECK pass updateReadOnlyFiled inside updateFilePropertiesById
      if(success && authorId) {
        const user = await this.appData.getUserInfo(authorId);
        // TOCHECK this call is commented temporally
        // if (user.LoginName)
          // await this.sharepoint.updateReadOnlyField(rootFolder, newFileInfo.value[0].ListItemAllFields.ID, 'Editor', user.LoginName);
      }
    }

    return success;
  }

  /** TODEL */
  // async cloneForecastModel(originFile: NPPFile, newFilename: string, newScenarios: number[], comments = ''): Promise<boolean> {

  //   const destinationFolder = originFile.ServerRelativeUrl.replace('/' + originFile.Name, '/');

  //   let success = await this.appData.cloneFile(originFile.ServerRelativeUrl, destinationFolder, newFilename);
  //   if (!success) return false;

  //   let newFileInfo = await this.appData.getFileByName(destinationFolder, newFilename);

  //   if (newFileInfo.value[0].ListItemAllFields && originFile.ListItemAllFields) {
  //     const newData = {
  //       ModelScenarioId: newScenarios,
  //       Comments: comments ? comments : null,
  //       ApprovalStatusId: await this.appData.getMasterApprovalStatusId("In Progress")
  //     }
  //     success = await this.appData.updateFilePropertiesById(newFileInfo.value[0].ListItemAllFields.ID, FILES_FOLDER, newData);
  //   }

  //   return success;
  // }

  /** Get the equivalent folder for Power BI files */
  private getPowerBICSVRootPathFromModelPath(path: string): string | undefined {
    let mappings: any = {}
    mappings[FOLDER_DOCUMENTS] =  FOLDER_POWER_BI_DOCUMENTS,
    mappings[FOLDER_WIP] =  FOLDER_POWER_BI_WIP,
    mappings[FOLDER_APPROVED] =  FOLDER_POWER_BI_APPROVED,
    mappings[FOLDER_ARCHIVED] =  FOLDER_POWER_BI_ARCHIVED
    
    for (const [key, value] of Object.entries(mappings)) {
      if(path.indexOf(key) !== -1) {
        return value as string;
      }
    }
    return undefined;
  }

}
