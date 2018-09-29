import { sp } from "@pnp/sp"
import { ISpGroupSvc, ISpGroup, ISpUser } from "../models";
import { Draft } from "../utils/draft"
import { BaseWebPartContext } from "@microsoft/sp-webpart-base";


export class PnPSpGroupSvc implements ISpGroupSvc {


    private readonly groupEndpoint =
        {
            expandables: ["Owners", "Users"],
            selectables: [
                "Owner/Email",
                "Owner/Id",
                "Owner/Title",
                "Users/Email",
                "Users/Email",
                "Users/Id",
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
    // AddGroupMembers: (group: ISpGroup, users: ISpUser[]) => Promise<void>;
    // RemoveGroupMembers: (group: ISpGroup, usersToRemove: ISpUser[]) => Promise<void>;
    // GetAllGroupMembers: (group: ISpGroup) => Promise<ISpGroup[]>;
    // AddGroup: (group: ISpGroup) => Promise<ISpGroup>;
    // DeleteGroup: (group: ISpGroup) => Promise<void>;

}
