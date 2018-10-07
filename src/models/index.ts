import { Draft } from "../utils/draft"

export interface ISpGroupSvc {
    GetGroups: () => Promise<Array<ISpGroup>>,
    UpdateGroup: (groupId: number, changes: Draft<ISpGroup>) => Promise<void>;
    AddGroupMembers: (groupId: number, users: Array<ISpUser>) => Promise<void>
    RemoveGroupMembers: (groupId: number, usersToRemove: Array<ISpUser>) => Promise<void>
    GetAllGroupMembers: (groupId: number) => Promise<Array<ISpGroup>>
    AddGroup: (group: ISpGroup) => Promise<ISpGroup>;
    DeleteGroup: (groupId: number) => Promise<void>
}

export interface IUsersSvc {
    GetUsersSuggestions: (searchText: string) => Promise<Array<IGraphUser>>
}

export interface ISpGroup {
    AllowMembersEditMembership: boolean
    AllowRequestToJoinLeave: boolean
    AutoAcceptRequestToJoinLeave: boolean
    Description: string
    Id: number
    IsHiddenInUI: boolean
    LoginName: string
    OnlyAllowMembersViewMembership: boolean
    OwnerTitle: string //"Spiewak, Jacek"
    RequestToJoinLeaveEmailSetting: string
    Owner: ISpUser
    Title: string
    Users: Array<ISpUser>
}

export interface ISpUser {
    Title: string,
    Id: number,
    Email: string
}

export interface IGraphUser {
    Email:string,
    Title: string,
    PhotoUrl: string
}


