import { Injectable } from '@angular/core';
import { SelectInputList } from '@shared/models/app-config';
import { Opportunity } from '@shared/models/entity';
import { User } from '@shared/models/user';
import { FILES_FOLDER, FOLDER_WIP } from '@shared/sharepoint/folders';
import { AppDataService } from './app/app-data.service';

@Injectable({
  providedIn: 'root'
})
export class SelectListsService {

  masterTherapiesList: SelectInputList[] = [];

  constructor(
    private readonly appData: AppDataService
  ) { }

  async getOpportunityFilterFields() {
    return [
      { value: 'title', label: 'Opportunity Name' },
      { value: 'projectStart', label: 'Project Start Date' },
      { value: 'projectEnd', label: 'Project End Date' },
      { value: 'opportunityType', label: 'Project Type' },
      { value: 'molecule', label: 'Molecule' },
      { value: 'indication', label: 'Indication' },
    ];
  }

  async getBrandFilterFields() {
    return [
      { value: 'Title', label: 'Brand Name' },
      //{ value: 'FCDueDate', label: 'Forecast Cycle Due Date' },
      { value: 'BusinessUnit.Title', label: 'Business Unit' },
      { value: 'Indication.Title', label: 'Indication Name' },
    ];
  }

  async getOpportunityTypesList(type: string | null = null): Promise<SelectInputList[]> {
    let res = await this.appData.getOpportunityTypes(type);
    return res.map(t => { return { value: t.ID, label: t.Title, extra: t } });
  }

  async getCountriesList(): Promise<SelectInputList[]> {
    const masterCountries = await this.appData.getMasterCountries();
    return masterCountries.map(t => { return { value: t.ID, label: t.Title } });
  }

  async getGeographiesList(): Promise<SelectInputList[]> {
    const masterGeographies = await this.appData.getMasterGeographies();
    return masterGeographies.map(t => { return { value: t.ID, label: t.Title } });
  }

  async getScenariosList(): Promise<SelectInputList[]> {
    const masterScenarios = await this.appData.getMasterScenarios();
    return masterScenarios.map(t => { return { value: t.ID, label: t.Title } });
  }

  async getClinicalTrialPhases(): Promise<SelectInputList[]> {
    const masterCTP = await this.appData.getMasterClinicalTrialPhases();
    return masterCTP.map(t => { return { value: t.ID, label: t.Title } });
  }

  async getBusinessUnitsList(): Promise<SelectInputList[]> {
    const masterBU = await this.appData.getMasterBusinessUnits();
    return masterBU.map(el => { return {value: el.ID, label: el.Title }});
  }

  async getForecastCyclesList(): Promise<SelectInputList[]> {
    const masterFC = await this.appData.getMasterForecastCycles();
    return masterFC.map(el => { return {value: el.ID, label: el.Title }});
  }

  async getUsersList(usersId: number[]): Promise<SelectInputList[]> {
    const users = await this.appData.getUsersByIds(usersId);
    return users.map((u: User) => { return { label: u.Title!, value: u.Email! } });
  }

  async getSiteOwnersList(): Promise<SelectInputList[]> {
    const owners = await this.appData.getSiteOwners();
    return owners.map(v => { return { label: v.Title ? v.Title + ' (' + v.Email + ')' : '', value: v.Id } });
  }

  async getStageNumbersList(stageType: string): Promise<SelectInputList[]> {
    const stages = await this.appData.getMasterStages(stageType);
    return stages.map(v => { return { label: v.Title, value: v.StageNumber } });
  }

  /** List of all therapies names (no indications related) */
  async getTherapiesList(): Promise<SelectInputList[]> {
    const indications = await this.appData.getMasterIndications();

    return indications
      .map(v => v.TherapyArea)
      .filter((value, index, self) => self.indexOf(value) === index)
      .map(v => { return { label: v, value: v } });
  }

  async getIndicationsList(therapy?: string): Promise<SelectInputList[]> {
    let indications = await this.appData.getMasterIndications(therapy);

    if (therapy) {
      return indications.map(el => { return { value: el.ID, label: el.Title } })
    }
    return indications.map(el => { return { value: el.ID, label: el.Title, group: el.TherapyArea } })
  }

  /** Accessible Geographies for the user (subfolders with read/write permission) */
  async getEntityAccessibleGeographiesList(entity: Opportunity, stageId?: number): Promise<SelectInputList[]> {
    const geographiesList = await this.appData.getEntityGeographies(entity.ID);

    let folder;
    if (stageId) {
      folder = `${FILES_FOLDER}/${entity.BusinessUnitId}/${entity.ID}/${stageId}/0`;
    } else {
      folder = `${FOLDER_WIP}/${entity.BusinessUnitId}/${entity.ID}/0/0`;
    }

    const geoFoldersWithAccess = await this.appData.getSubfolders(folder, true);
    return geographiesList.filter(mf => geoFoldersWithAccess.some((gf: any) => +gf.Name === mf.Id))
      .map(t => { return { value: t.Id, label: t.Title } });
  }
}
