import { User } from "./user";

export interface NPPNotification {
    Id: number;
    Title: string;
    TargetUserId: number;
    TargetUser?: User;
}