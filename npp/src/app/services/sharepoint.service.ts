import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Observable, of } from 'rxjs';
import { ErrorService } from './error.service';
import { LicensingService } from './licensing.service';
import { map } from 'rxjs/operators';
import { NPPFile, NPPFileMetadata, SystemFolder } from '@shared/models/file-system';
import * as SPFolders from '@shared/sharepoint/folders';

export const ReadPermission = 'ListRead';
export const WritePermission = 'ListRead';

export interface SelectInputList {
  label: string;
  value: any;
  group?: string;
}

interface SharepointResult {
  'odata.metadata': string;
  value: any;
}

interface FilterTerm {
  term: string;
  field?: string;
  matchCase?: boolean;
}

interface SPGroup {
  Id: number;
  Title: string;
  Description: string;
  LoginName: string;
  OnlyAllowMembersViewMembership: boolean;
}


// export interface AppType {
//   ID: number;
//   Title: string;
// }


@Injectable({
  providedIn: 'root'
})
export class SharepointService {
  
  SPRoleDefinitions: {
    name: string;
    id: number;
  }[] = [];
  // public app: AppType | undefined;

  constructor(
    private http: HttpClient, 
    private error: ErrorService, 
    private licensing: LicensingService, 
  ) { }

  

  query(partial: string, conditions: string = '', count: number | 'all' = 'all', filter?: FilterTerm): Observable<any> {
    //TODO implement usage of count

    let filterUri = '';
    if (filter && filter.term) {
      filter.field = filter.field ? filter.field : 'Title';
      filter.matchCase = filter.matchCase ? filter.matchCase : false;

      if (filter.matchCase) {
        filterUri = `$filter=substringof('${filter.term}',${filter.field})`;
      } else {
        let capitalized = filter.term.charAt(0).toUpperCase() + filter.term.slice(1);
        filterUri = `$filter=substringof('${filter.term}',${filter.field}) or substringof('${capitalized}',${filter.field})`;
      }
    }
    let endpoint = this.licensing.getSharepointApiUri() + partial;
    if (conditions || filterUri) endpoint += '?';
    if (conditions) endpoint += conditions;
    if (filterUri) endpoint += conditions ? '&' + filterUri : filterUri;
    try {
      return this.http.get(endpoint);
    } catch (e: any) {
      this.error.handleError(e);
      return of([]);
    }
  }

  async getAllItems(listName: string, conditions: string = ''): Promise<any[]> {
    try {
      let endpoint = this.licensing.getSharepointApiUri() + this.getListUri(listName) + '/items';
      if (conditions) endpoint += '?' + conditions;
      const listResult = await this.http.get(endpoint).toPromise() as SharepointResult;
      if (listResult.value && listResult.value.length > 0) {
        return listResult.value;
      }
      return [];
    } catch (e: any) {
      this.error.handleError(e);
      return [];
    }
  }

  async getOneItem(listName: string, conditions: string = ''): Promise<any> {
    try {
      let endpoint = this.licensing.getSharepointApiUri() + this.getListUri(listName) + '/items';
      if (conditions) endpoint += '?' + conditions;
      let lists = await this.http.get(endpoint).toPromise() as SharepointResult;
      if (lists.value && lists.value.length == 1) {
        return lists.value[0];
      }
      return null;
    } catch (e: any) {
      this.error.handleError(e);
      return null;
    }
  }

  async getOneItemById(id: number, listName: string, conditions: string = ''): Promise<any> {
    try {
      let endpoint = this.licensing.getSharepointApiUri() + this.getListUri(listName) + `/items(${id})`;
      if (conditions) endpoint += '?' + conditions;
      return await this.http.get(endpoint).toPromise();
    } catch (e: any) {
      this.error.handleError(e);
      return null;
    }
    return null;
  }

  async countItems(listName: string, conditions: string = ''): Promise<number> {
    try {
      let endpoint = this.licensing.getSharepointApiUri() + this.getListUri(listName) + '/ItemCount';
      if (conditions) endpoint += '?' + conditions;
      let lists = await this.http.get(endpoint).toPromise() as SharepointResult;
      if (lists.value) {
        return lists.value;
      }
      return 0;
    } catch (e: any) {
      this.error.handleError(e);
      return 0;
    }
  }

  async createItem(listName: string, data: any): Promise<any> {
    try {
      return await this.http.post(
        this.licensing.getSharepointApiUri() + this.getListUri(listName) + "/items",
        data
      ).toPromise();
    } catch (e: any) {
      this.error.handleError(e);
      return null;
    }
  }

  public async updateItem(id: number, listName: string, data: any): Promise<boolean> {
    try {
      await this.http.post(
        this.licensing.getSharepointApiUri() + this.getListUri(listName) + `/items(${id})`,
        data,
        {
          headers: new HttpHeaders({
            'If-Match': '*',
            'X-HTTP-Method': "MERGE"
          })
        }
      ).toPromise();
    } catch (e: any) {
      this.error.handleError(e);
      return false;
    }
    return true;
  }

  public async deleteItem(id: number, listName: string): Promise<boolean> {
    try {
      await this.http.post(
        this.licensing.getSharepointApiUri() + this.getListUri(listName) + `/items(${id})`,
        null,
        {
          headers: new HttpHeaders({
            'If-Match': '*',
            'X-HTTP-Method': "DELETE"
          }),
        }
      ).toPromise();
      return true;
    } catch (e: any) {
      this.error.handleError(e);
      return false;
    }
  }

  private getListUri(listName: string): string {
    return `lists/getbytitle('${listName}')`;
  }

  /** --- FILES --- **/

  /** TOCHECK pillar a les crides la folder directament */
  getBaseFilesFolder(): string {
    return  SPFolders.FILES_FOLDER;
  }

  /** ok */
  async createFolder(folderPath: string): Promise<SystemFolder | null> {
    try {
      return await this.http.post(
        this.licensing.getSharepointApiUri() + "folders",
        {
          ServerRelativeUrl: folderPath
        }
      ).toPromise() as SystemFolder;
    } catch (e: any) {
      console.log("Error creating folder: "+e.message);
      this.error.handleError(e);
      return null;
    }
  }

  /** ok */
  async getFolderByUrl(folderUrl: string): Promise<SystemFolder | null> {
    try {
      let folder = await this.query(
        `GetFolderByServerRelativeUrl('${folderUrl}')`
      ).toPromise();
      return folder ? folder : null;
    } catch (e) {
      return null;
    }
  }

  /** ok */
  async readFile(fileUri: string): Promise<any> {
    try {
      return this.http.get(
        this.licensing.getSharepointApiUri() + `GetFileByServerRelativeUrl('${fileUri}')/$value`,
        { responseType: 'arraybuffer' }
      ).toPromise();
    } catch (e: any) {
      this.error.handleError(e);
      return [];
    }
  }

  /** ok */
  async deleteFile(fileUri: string): Promise<boolean> {
    try {
      await this.http.post(
        this.licensing.getSharepointApiUri() + `GetFileByServerRelativeUrl('${fileUri}')`,
        null,
        {
          headers: new HttpHeaders({
            'If-Match': '*',
            'X-HTTP-Method': "DELETE"
          }),
        }
      ).toPromise();
    } catch (e: any) {
      this.error.handleError(e);
      return false;
    }
    return true;
  }

  /** ok */
  async renameFile(fileUri: string, newName: string): Promise<boolean> {
    try {
      await this.http.post(
        this.licensing.getSharepointApiUri() + `GetFileByServerRelativeUrl('${fileUri}')/ListItemAllFields`,
        {
          Title: newName,
          FileLeafRef: newName
        },
        {
          headers: new HttpHeaders({
            'If-Match': '*',
            'X-HTTP-Method': "MERGE"
          }),
        }
      ).toPromise();
    } catch (e) {
      return false;
    }
    return true;
  }
  
  /** ok */
  async existsFile(filename: string, folder: string): Promise<boolean> {
    try {
      let file = await this.query(
        `GetFolderByServerRelativeUrl('${folder}')/Files`,
        `$expand=ListItemAllFields&$filter=Name eq '${filename}'`,
      ).toPromise();
      return file.value.length > 0;
    } catch (e) {
      return false;
    }
  }

  /** ok */
  async cloneFile(originServerRelativeUrl: string, destinationFolder: string, newFileName: string): Promise<boolean> {
    const originUrl = `getfilebyserverrelativeurl('${originServerRelativeUrl}')/`;
    let destinationUrl = `copyTo('${destinationFolder + newFileName}')`;
    try {
      const r = await this.http.post(
        this.licensing.getSharepointApiUri() + originUrl + destinationUrl,
        null
      ).toPromise();
      return true;
    }
    catch (e) {
      return false;
    }
  }

  /** TODEL ? */
  async readFolderFiles(folder: string, expandProperties = false): Promise<NPPFile[]> {
    let files: NPPFile[] = []
    const result = await this.query(
      `GetFolderByServerRelativeUrl('${this.getBaseFilesFolder()}/${folder}')/Files`,
      '$expand=ListItemAllFields',
    ).toPromise();

    if (result.value) {
      files = result.value;
    }
    if (expandProperties && files.length > 0) {
      for (let i = 0; i < files.length; i++) {
        let fileItems = files[i].ListItemAllFields;
        if (fileItems) {
          fileItems = Object.assign(fileItems, await this.getFileInfo(fileItems.ID));
        }
      }
    }
    return files;
  }

  async getEntityFileFromURL(url: string)  {
    return await this.query(
      `GetFileByServerRelativeUrl('${url}')/listItemAllFields`
    ).toPromise();
  }

  /** TOCHECK creada nova per treure crida de AppData, fa el mateix que getEntityFileFromURL()? */
  async readFileMetadata(url: string): Promise<NPPFileMetadata> {
    return (await this.http.get(
      this.licensing.getSharepointApiUri() + `GetFileByServerRelativeUrl('${url}')/ListItemAllFields`).toPromise()) as NPPFileMetadata;
  }

  /** TODEL */
  async getFileInfo(fileId: number): Promise<NPPFile> {
    return await this.query(
      `lists/getbytitle('${SPFolders.FILES_FOLDER}')` + `/items(${fileId})`,
      '$select=*,Author/Id,Author/FirstName,Author/LastName,StageName/Id,StageName/Title, \
        EntityGeography/Title,EntityGeography/EntityGeographyType,ModelScenario/Title,ApprovalStatus/Title \
        &$expand=StageName,Author,EntityGeography,ModelScenario,ApprovalStatus',
      'all'
    ).toPromise();
  }
  
  /** ok */
  async uploadFileQuery(fileData: string, folder: string, filename: string) {
    try {
      let url = `GetFolderByServerRelativeUrl('${folder}')/Files/add(url='${filename}',overwrite=true)?$expand=ListItemAllFields`;
      return await this.http.post(
        this.licensing.getSharepointApiUri() + url,
        fileData,
        {
          headers: { 'Content-Type': 'blob' }
        }
      ).toPromise();
    } catch (e: any) {
      this.error.handleError(e);
      return {};
    }
  }

  /** --- PERMISSIONS --- **/

  /** Create a Sharepoint group */
  async createGroup(name: string, description: string = ''): Promise<SPGroup | null> {
    try {
      return await this.http.post(
        this.licensing.getSharepointApiUri() + 'sitegroups',
        {
          Title: name,
          Description: description,
          OnlyAllowMembersViewMembership: false
        }
      ).toPromise() as SPGroup;
    } catch (e: any) {
      this.error.handleError(e);
      return null;
    }
  }

  /** Deletes the sharepoint group by Id */
  async deleteGroup(id: number): Promise<boolean> {
    try {
      await this.http.post(
        this.licensing.getSharepointApiUri() + `/sitegroups/removebyid(${id})`,
        null,
        {
          headers: new HttpHeaders({
            'If-Match': '*',
            'X-HTTP-Method': "DELETE"
          })
        }
      ).toPromise();
      return true;
    } catch (e) {
      return false;
    }
  }
  
  /** ok */
  async addRolePermissionToList(list: string, groupId: number, roleName: string, id: number = 0): Promise<boolean> {
    const baseUrl = this.licensing.getSharepointApiUri() + list + (id === 0 ? '' : `/items(${id})`);
    return await this.setRolePermission(baseUrl, groupId, roleName);
  }

  /** ok */
  async addRolePermissionToFolder(folderUrl: string, groupId: number, roleName: string): Promise<boolean> {
    const baseUrl = this.licensing.getSharepointApiUri() + `GetFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields`;
    // permissions to folders without inherit
    let success = await this.setRolePermission(baseUrl, groupId, roleName, false);
    // TOCHECK move remove to appData
    // return success && await this.removeRolePermission(baseUrl, (await this.getCurrentUserInfo()).Id);
    return success;
  }

  /** ok */
  private async setRolePermission(baseUrl: string, groupId: number, roleName: string, inherit = true) {
    // const roleId = 1073741826; // READ
    const roleId = await this.getRoleDefinitionId(roleName);
    try {
      await this.http.post(
        baseUrl + `/breakroleinheritance(copyRoleAssignments=${inherit ? 'true' : 'false'},clearSubscopes=${inherit ? 'true' : 'false'})`,
        null).toPromise();
      await this.http.post(
        baseUrl + `/roleassignments/addroleassignment(principalid=${groupId},roledefid=${roleId})`,
        null).toPromise();
      return true;
    } catch (e: any) {
      this.error.handleError(e);
      return false;
    }
  }

  /** ok */
  private async removeRolePermission(baseUrl: string, groupId: number) {
    try {
      await this.http.post(
        baseUrl + `/roleassignments/removeroleassignment(principalid=${groupId})`,
        null).toPromise();
      return true;
    } catch (e: any) {
      this.error.handleError(e);
      return false;
    }
  }

  /** TOCHECK no ha d'anar aquí, però on? */
  searchByTermInputList(query: string, field: string, term: string, matchCase = false): Observable<SelectInputList[]> {
    return this.query(query, '', 'all', { term, field, matchCase })
      .pipe(
        map((res: any) => {
          return res.value.map(
            (el: any) => { return { value: el.Id, label: el.Title } as SelectInputList }
          );
        })
      );
  }
  
/*
  //return all geographies for now
  async getBrandAccessibleGeographiesList(brand: Brand): Promise<SelectInputList[]> {
    const geographiesList = await this.getBrandGeographies(brand.ID);

    const geoFoldersWithAccess = await this.getSubfolders(`${SPLists.FOLDER_WIP}/${brand.BusinessUnitId}/${brand.ID}/${SPLists.FORECAST_MODELS_FOLDER_NAME}`);
    return geographiesList.filter(mf => geoFoldersWithAccess.some((gf: any) => +gf.Name === mf.Id))
      .map(t => { return { value: t.Id, label: t.Title } });
  }
*/
  
  
/*
  async getBrandGeographies(brandId: number, all?: boolean) {
    let filter = `$filter=BrandId eq ${brandId}`;
    if (!all) {
      filter += ' and Removed ne 1';
    }
    return await this.getAllItems(
       SPLists.GEOGRAPHIES_LIST, filter,
    );
  }*/

  /** ok */
  /** Updates a read only field fieldname of the list's element with the value */
  async updateReadOnlyField(list: string, elementId: number, fieldname: string, value: string) {

    await this.http.post(
      this.licensing.getSharepointApiUri() + `lists/getByTitle('${list}')/items(${elementId})/validateUpdateListItem`,
      JSON.stringify({
        "formValues": [
          {
            "__metadata": { "type": "SP.ListItemFormUpdateValue" },
            "FieldName": fieldname,
            "FieldValue": "[{'Key':'" + value + "'}]"
          }
        ],
        "bNewDocumentUpdate": false
      }),
      {
        headers: new HttpHeaders({
          "Accept": "application/json; odata=verbose",
          "Content-Type": "application/json; odata=verbose"
        })
      }).toPromise();
  }
  
  /** ok */
  async copyFile(originServerRelativeUrl: string, destinationFolder: string, newFileName: string): Promise<any> {
    const originUrl = `getfilebyserverrelativeurl('${originServerRelativeUrl}')/`;
    let path = destinationFolder + newFileName;
    let destinationUrl = `copyTo('${path}')`;
    try {
      const r = await this.http.post(
        this.licensing.getSharepointApiUri() + originUrl + destinationUrl,
        null
      ).toPromise();
      return path;
    }
    catch (e) {
      return false;
    }
  }

  /** ok */
  async moveFile(originServerRelativeUrl: string, destinationFolder: string, newFilename: string = ''): Promise<any> {
    let arrUrl = originServerRelativeUrl.split("/");
    let fileName = arrUrl[arrUrl.length - 1];
    const originUrl = `getfilebyserverrelativeurl('${originServerRelativeUrl}')/`;
    let path = destinationFolder + "/" + (newFilename ? newFilename : fileName);
    let destinationUrl = `moveTo('${path}')`;
    const r = await this.http.post(
      this.licensing.getSharepointApiUri() + originUrl + destinationUrl,
      null
    ).toPromise();

    return "/"+arrUrl[1]+"/"+arrUrl[2]+"/"+path;
  }

  /** ok */
  async updateFileFields(path: string, fields: any) {
    this.http.post(
      this.licensing.getSharepointApiUri() + `GetFileByServerRelativeUrl('${path}')/ListItemAllFields`,
      fields,
      {
        headers: new HttpHeaders({
          'If-Match': '*',
          'X-HTTP-Method': "MERGE"
        }),
      }
    ).toPromise();
  }

  /** ok */
  /** Adds the user to the Sharepoint Group */
  async addUserToSharepointGroup(userLoginName: string, groupId: number) {
    try {
      await this.http.post(
        this.licensing.getSharepointApiUri() + `sitegroups(${groupId})/users`,
        { LoginName: userLoginName }
      ).toPromise();
      return true;
    } catch (e: any) {
      return false;
    }
  }

  /** ok */
  /** Remove the user from the Sharepoint Group (id or name) */
  async removeUserFromSharepointGroup(userId: number, group: number | string): Promise<boolean> {
    let url = '';
    if (typeof group == 'string') {
      url = this.licensing.getSharepointApiUri() + `sitegroups/getbyname('${group}')/users/removebyid(${userId})`;
    } else if (typeof group == 'number') {
      url = this.licensing.getSharepointApiUri() + `sitegroups(${group})/users/removebyid(${userId})`;
    }
    try {
      await this.http.post(
        url,
        null,
        {
          headers: new HttpHeaders({
            'If-Match': '*',
            'X-HTTP-Method': "DELETE"
          })
        }
      ).toPromise();
      return true
    } catch (e: any) {
      return false;
    }
  }

  /** ok */
  async getPathSubfolders(path: string) {
    const result = await this.query(
      `GetFolderByServerRelativeUrl('${path}')/folders`,
      '$expand=ListItemAllFields',
    ).toPromise();
    return result.value ? result.value : [];
  }

  /** ok */
  private async getRoleDefinitionId(name: string): Promise<number | null> {
    const cache = this.SPRoleDefinitions.find(g => g.name === name);
    if (cache) {
      return cache.id;
    } else {
      try {
        const result = await this.query(`roledefinitions/getbyname('${name}')/id`).toPromise();
        this.SPRoleDefinitions.push({ name, id: result.value }); // add for local caching
        return result.value;
      }
      catch (e) {
        return null;
      }
    }
  }
} 
