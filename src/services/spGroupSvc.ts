import { sp } from "@pnp/sp"
import { ISpGroupSvc, ISpGroup, ISpUser, IUserSuggestion } from "../models";
import { Draft } from "../utils/draft"
import { BaseWebPartContext } from "@microsoft/sp-webpart-base";


export class PnPSpGroupSvc implements ISpGroupSvc {
    AddGroupMembers: (groupId: number, users: IUserSuggestion[]) => Promise<void>;
    RemoveGroupMembers: (groupId: number, usersToRemove: ISpUser[]) => Promise<void>;
    GetAllGroupMembers: (groupId: number) => Promise<ISpGroup[]>;
    DeleteGroup: (groupId: number) => Promise<void>;


    private readonly groupEndpoint =
        {
            expandables: ["Owner"],
            selectables: [
                "Owner/Email",
                "Owner/Id",
                "Owner/Title",
                "AllowMembersEditMembership",
                "AllowRequestToJoinLeave",
                "AutoAcceptRequestToJoinLeave",
                "Description",
                "Id",
                "IsHiddenInUI",
                "LoginName",
                "OnlyAllowMembersViewMembership",
                "RequestToJoinLeaveEmailSetting",
                "Title"
            ]
        }



    constructor(ctx: BaseWebPartContext) {
        sp.setup({
            spfxContext: ctx
        })

    }
    async GetGroups() {
        let groups = await sp.web.siteGroups.expand(...this.groupEndpoint.expandables).select(...this.groupEndpoint.selectables).get<Array<ISpGroup>>();
        return groups;
    }
    async UpdateGroup(groupId: number, changes: Draft<ISpGroup>){
        await sp.web.siteGroups.getById(groupId).update(
            changes
        );
    }

    async GetUsersFromGroup(groupId: number){
        let selectables = ["Email","Title", "Id"]
        return sp.web.siteGroups.getById(groupId).users.select(...selectables).get()
    }

    async AddGroup(group:ISpGroup) {
        delete group.Id;
        let result = await sp.web.siteGroups.add(
            group
        )

        //TODO check if this returns what you think it returns
        return result.data;
    }
     

}
