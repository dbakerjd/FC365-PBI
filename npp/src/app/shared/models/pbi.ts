
export interface PBIReport {
  ID: number;
  name: string;
  GroupId: string;
  pageName: string;
  Title: string;
}

export interface PBIRefreshComponent {
  ComponentName: string;
  GroupId: string;
  ComponentType: string;
}

export interface PBIDataset {
  id: string;
  name: string;
  addRowsAPIEnabled: boolean;
  configuredBy: string;
  isRefreshable: true;
  isEffectiveIdentityRequired: boolean;
  isEffectiveIdentityRolesRequired: boolean;
  isOnPremGatewayRequired: boolean;
}

export interface PBIDatasetRefresh {
  endTime: Date;
  id: number;
  refreshType: string;
  requestId: string;
  startTime: Date;
  status: string;
}