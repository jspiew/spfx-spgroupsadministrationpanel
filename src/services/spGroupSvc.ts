import { sp } from "@pnp/sp"
import { ISpGroupSvc, ISpGroup, ISpUser, IUserSuggestion } from "../models";
import { Draft } from "../utils/draft"
import { BaseWebPartContext } from "@microsoft/sp-webpart-base";


export class PnPSpGroupSvc implements ISpGroupSvc {
    
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
    async GetGroups(ids?:number[]) {
        if (ids) {
            let batch = sp.createBatch();
            let groupPromises = ids.map(i => { return sp.web.siteGroups.getById(i).expand(...this.groupEndpoint.expandables).select(...this.groupEndpoint.selectables).inBatch(batch).get<ISpGroup>()});
            await batch.execute();
            return (await Promise.all(groupPromises))
        } else  {
            let groups = await sp.web.siteGroups.expand(...this.groupEndpoint.expandables).select(...this.groupEndpoint.selectables).get<Array<ISpGroup>>();
            return groups;
        }
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

    async AddGroupMembers(groupId: number, users: IUserSuggestion[]) {

        let batch = sp.createBatch();
        let ensuredUsersBatch = sp.createBatch();
        let ensuredUsersPromises = users.map(u => sp.web.inBatch(ensuredUsersBatch).ensureUser(u.Email));
        await ensuredUsersBatch.execute();
        let ensuredUsers = await Promise.all(ensuredUsersPromises);

        ensuredUsers.forEach(u => { sp.web.siteGroups.getById(groupId).users.inBatch(batch).add(u.data.LoginName)})
        return batch.execute(); 
     }

    async RemoveGroupMembers(groupId: number, usersToRemove: ISpUser[]){
        let batch = sp.createBatch();
        usersToRemove.forEach(u => { sp.web.siteGroups.getById(groupId).users.inBatch(batch).removeById(u.Id) })
        return batch.execute(); 
    }

    async GetGroupsForDropdown() {
        return  (await sp.web.siteGroups.select("Title","Id").get()) as {Title: string, Id: number}[]
    }
     

}
