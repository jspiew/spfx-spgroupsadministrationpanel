import { Draft } from "../utils/draft";

export interface ISpGroupSvc {
    GetGroups: (ids?: number[]) => Promise<Array<ISpGroup>>;
    GetUsersFromGroup: (groupId: number) => Promise<Array<ISpUser>>;
    UpdateGroup: (groupId: number, changes: Draft<ISpGroup>) => Promise<any>;
    AddGroupMembers: (groupId: number, users: Array<IUserSuggestion>) => Promise<any>;
    RemoveGroupMembers: (groupId: number, usersToRemove: Array<ISpUser>) => Promise<any>;
    AddGroup: (group: ISpGroup) => Promise<ISpGroup>;
    DeleteGroup: (groupId: number) => Promise<any>;
    GetGroupsForDropdown:() => Promise<{Id: number, Title: string}[]>;
}

export interface IUsersSvc {
    GetUsersSuggestions: (searchText: string, includeGroups: boolean) => Promise<Array<IUserSuggestion|ISpUser>>;
    EnsureUser: (suggestion: IUserSuggestion) => Promise<ISpUser>;
}

export interface ISpGroup {
    AllowMembersEditMembership: boolean;
    AllowRequestToJoinLeave: boolean;
    AutoAcceptRequestToJoinLeave: boolean;
    Description: string;
    Id: number;
    IsHiddenInUI: boolean;
    LoginName: string;
    OnlyAllowMembersViewMembership: boolean;
    OwnerTitle: string; //"Spiewak, Jacek"
    RequestToJoinLeaveEmailSetting: string;
    Owner: ISpUser;
    Title: string;
    Users: Array<ISpUser>;
}


export interface IUserSuggestion {
    Email: string;
    Title: string;
}

export interface ISpUser extends IUserSuggestion {
    Id: number;
}

export function isISpUser(arg: any): arg is ISpUser {
    return arg.Id !== undefined && arg.Email !== undefined && arg.Title !== undefined;
}





