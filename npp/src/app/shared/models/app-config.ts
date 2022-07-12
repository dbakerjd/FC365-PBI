export interface AppType {
    ID: number;
    Title: string;
}

export interface SelectInputList {
    label: string;
    value: any;
    group?: string;
}

export interface LicenseContext {
    host: string;
    entityId?: string;
    teamSiteDomain?: string;
}

export interface StringMapping {
    ID: number;
    Key: string;
    Title: string;
}