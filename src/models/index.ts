import { Draft } from "../utils/draft"

export interface ISpGroupSvc {
    GetGroups: (ids?: number[]) => Promise<Array<ISpGroup>>
    GetUsersFromGroup: (groupId: number) => Promise<Array<ISpUser>>
    UpdateGroup: (groupId: number, changes: Draft<ISpGroup>) => Promise<void>
    UpdateGroupOwner: (groupId: number, owner: IUserSuggestion) => Promise<void>
    AddGroupMembers: (groupId: number, users: Array<IUserSuggestion>) => Promise<void>
    RemoveGroupMembers: (groupId: number, usersToRemove: Array<ISpUser>) => Promise<void>
    AddGroup: (group: ISpGroup) => Promise<ISpGroup>
    DeleteGroup: (groupId: number) => Promise<void>
    GetGroupsForDropdown:() => Promise<{Id: number, Title: string}[]>
}

export interface IUsersSvc {
    GetUsersSuggestions: (searchText: string) => Promise<Array<IUserSuggestion>>
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

export interface ISpUser extends IUserSuggestion {
    Id: number
}

export interface IUserSuggestion {
    Email:string,
    Title: string
}


