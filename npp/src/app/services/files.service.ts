import { Injectable } from '@angular/core';
import { EntityGeography, Opportunity } from '@shared/models/entity';
import { NPPFile } from '@shared/models/file-system';
import { FILES_FOLDER, FOLDER_APPROVED, FOLDER_ARCHIVED, FOLDER_DOCUMENTS, FOLDER_POWER_BI_APPROVED, FOLDER_POWER_BI_ARCHIVED, FOLDER_POWER_BI_DOCUMENTS, FOLDER_POWER_BI_WIP, FOLDER_WIP } from '@shared/sharepoint/folders';
import { AppDataService } from './app/app-data.service';

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

  /** Get the upload folder according to the selected options */
  async constructUploadFolder(entity: Opportunity, stageId: number | undefined, categoryId: number | undefined, geographyId: number | undefined) {

    let fileFolder = '/' + entity.BusinessUnitId + '/' + entity.ID + '/';
    let rootFolder = FILES_FOLDER;

    if (stageId) {
      fileFolder += stageId + '/';
    } else {
      fileFolder += '0/';
      rootFolder = FOLDER_WIP;
    }

    if (categoryId !== 0) { // document folder selected
      rootFolder = FOLDER_DOCUMENTS;
      fileFolder += categoryId + '/';
    } else {
      fileFolder += '0/';
    }

    if (geographyId) {
      fileFolder += geographyId + '/';  
    } else {
      fileFolder += '0/';
    }
    return rootFolder + fileFolder + '0'; // always 0 for forecast cycle in upload
  }

  /** Prepare regular file information data to upload */
  async prepareUploadFileData(data: any) {
    return {
      StageNameId: data.StageNameId,
      EntityNameId: data.EntityNameId,
      Comments: data.description
    };
  }

  /** Prepare a model information data to upload */
  async prepareUploadModelData(data: any) {
    const user = await this.appData.getCurrentUserInfo();
    const userName = user.Title && user.Title.indexOf("@") == -1 ? user.Title : user.Email;

    return {
      EntityNameId: data.EntityNameId,
      StageNameId: data.StageNameId,
      EntityGeographyId: data.geography,
      ModelScenarioId: data.scenario,
      Comments: data.description ? '[{"text":"'+data.description.replace(/'/g, "{COMMA}")+'","email":"'+user.Email+'","name": "'+ userName +'","userId":'+user.Id+',"createdAt":"'+new Date().toISOString()+'"}]' : '[]',
      ApprovalStatusId: await this.appData.getMasterApprovalStatusId("In Progress"),
      IndicationId: data.IndicationId
    };
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

  async fileExistsInFolder(filename: string, folder: string) {
    return await this.appData.existsFile(filename, folder);
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

  /** Restart the state and the model files information of the folder */
  async restartModelsInFolder(entity: Opportunity, folder: string) {

    const geographies = await this.appData.getEntityGeographies(entity.ID); // 1 = stage id would be dynamic in the future

    let arrFolder = folder.split("/");
    let rootFolder = arrFolder[0];

    for (let i = 0; i < geographies.length; i++) {
      let geo = geographies[i];
      let files = await this.appData.getFolderFiles(folder + "/" + geo.ID + "/0", true);
      for (let j = 0; files && j < files.length; j++) {
        let model = files[j];
        await this.setFileApprovalStatus(rootFolder, model, entity, 'In Progress');
        await this.cleanFileComments(rootFolder, model);
      }
    }

  }

  /** Cleans the comments history of the file */
  async cleanFileComments(rootFolder: string, file: NPPFile) {
    if (!file.ListItemAllFields) return;
    await this.appData.updateFilePropertiesById(file.ListItemAllFields.ID, rootFolder, { Comments: "[]" });
  }

  /** Set the approval status for a file */
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
        await this.removeOldApprovedModel(entity, file);
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

  async addScenarioSufixToFilename(originFilename: string, scenarioId: number): Promise<string | false> {
    const scenarios = await this.appData.getMasterScenarios();
    const extension = originFilename.split('.').pop();
    if (!extension) return false;

    const baseFileName = originFilename.substring(0, originFilename.length - (extension.length + 1));
    return baseFileName
      + '-' + scenarios.find(el => el.ID === scenarioId)?.Title.replace(/ /g, '').toLocaleLowerCase()
      + '.' + extension;
  }

  /** Count the number of files contained by the entity folder */
  async getFolderFilesCount(entityFolder: string) {
    const geoFolders = await this.appData.getSubfolders(entityFolder, true);
    let currentFiles = [];
    for (const geofolder of geoFolders) {
      let gf = entityFolder + '/' + geofolder.Name + '/0';
      currentFiles.push(...await this.appData.getFolderFiles(gf, true));
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

  private async removeOldApprovedModel(entity: Opportunity, file: NPPFile) {
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

    let fileWithSameScenarios = await this.getFileByScenarios(destinationFolder, newScenarios);
    if (fileWithSameScenarios) this.appData.deleteFile(fileWithSameScenarios.ServerRelativeUrl);

    let success = await this.appData.cloneFile(originFile.ServerRelativeUrl, destinationFolder, newFilename);
    if (!success) return false;

    let newFileInfo = await this.appData.getFileByName(destinationFolder, newFilename);

    if (newFileInfo[0].ListItemAllFields && originFile.ListItemAllFields) {
      const newData:any = {
        ModelScenarioId: newScenarios,
        Comments: comments ? comments : null,
        ApprovalStatusId: await this.appData.getMasterApprovalStatusId("In Progress")
      }
      
      let arrFolder = destinationFolder.split("/");
      let rootFolder = arrFolder[3];
      
      success = await this.appData.updateFilePropertiesById(newFileInfo[0].ListItemAllFields.ID, rootFolder, newData);
      if (success) {
        await this.appData.changeFileEditor(authorId, rootFolder, newFileInfo[0].ListItemAllFields.ID);
      }
    }
    return success;
  }

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
