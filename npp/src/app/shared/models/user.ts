export interface User {
    Id: number;
    LoginName?: string;
    FirstName?: string;
    LastName?: string;
    Title?: string;
    Email?: string;
    profilePicUrl?: string;
    IsSiteAdmin?: boolean;
}

export interface GroupPermission {
    Id: number;
    Title: string;
    ListName: string;
    Permission: string;
    ListFilter: 'List' | 'Item';
  }