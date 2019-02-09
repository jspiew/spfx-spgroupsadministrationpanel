import { sp } from "@pnp/sp";
import { ISpGroupSvc, ISpGroup, ISpUser, IUserSuggestion } from "../models";
import { Draft } from "../utils/draft";
import { BaseWebPartContext } from "@microsoft/sp-webpart-base";


export class PnPSpGroupSvc implements ISpGroupSvc {
    
    public DeleteGroup: (groupId: number) => Promise<void>;


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
        };



    constructor(ctx: BaseWebPartContext) {
        sp.setup({
            spfxContext: ctx
        });

    }
    public async GetGroups(ids?:number[]) {
        if (ids) {
            let batch = sp.createBatch();
            let groupPromises = ids.map(i => { return sp.web.siteGroups.getById(i).expand(...this.groupEndpoint.expandables).select(...this.groupEndpoint.selectables).inBatch(batch).get<ISpGroup>();});
            await batch.execute();
            return (await Promise.all(groupPromises));
        } else  {
            let groups = await sp.web.siteGroups.expand(...this.groupEndpoint.expandables).select(...this.groupEndpoint.selectables).get<Array<ISpGroup>>();
            return groups;
        }
    }
    public async UpdateGroup(groupId: number, changes: Draft<ISpGroup>){
        if (changes.Owner) delete changes.Owner.Email;

        await sp.web.siteGroups.getById(groupId).update(
            changes
        );
    }

    public async GetUsersFromGroup(groupId: number){
        let selectables = ["Email","Title", "Id"];
        return sp.web.siteGroups.getById(groupId).users.select(...selectables).get();
    }

    public async AddGroup(group:ISpGroup) {
        delete group.Id;
        let result = await sp.web.siteGroups.add(
            group
        );

        //TODO check if this returns what you think it returns
        return result.data;
    }

    public async AddGroupMembers(groupId: number, users: IUserSuggestion[]) {

        if(users.length == 0){
            return Promise.resolve<void>();
        }
        // let batch = sp.createBatch();
        // let ensuredUsersBatch = sp.createBatch();
        // batching temporarily remvoed due to execution issues
        let ensuredUsersPromises = users.map(u => sp.web.ensureUser(u.Email));
        let ensuredUsers = await Promise.all(ensuredUsersPromises);

        ensuredUsers.forEach(u => { sp.web.siteGroups.getById(groupId).users.add(u.data.LoginName);});
        return Promise.all(ensuredUsers) as Promise<any>;
     }

    public async RemoveGroupMembers(groupId: number, usersToRemove: ISpUser[]){
        if (usersToRemove.length == 0) {
            return Promise.resolve<void>();
        }
        let batch = sp.createBatch();
        let removePromises = usersToRemove.map(u => { return sp.web.siteGroups.getById(groupId).users.removeById(u.Id); });
        return Promise.all(removePromises); 
    }

    public async GetGroupsForDropdown() {
        return  (await sp.web.siteGroups.select("Title","Id").get()) as {Title: string, Id: number}[];
    }
     

}
