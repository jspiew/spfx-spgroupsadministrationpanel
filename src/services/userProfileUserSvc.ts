import { sp, PrincipalType } from "@pnp/sp"
import { IUsersSvc, ISpGroup, ISpUser, IUserSuggestion } from "../models";
import { Draft } from "../utils/draft"
import { BaseWebPartContext } from "@microsoft/sp-webpart-base";


export class UserProfileUserSvc implements IUsersSvc {

    constructor(ctx: BaseWebPartContext) {
        sp.setup({
            spfxContext: ctx
        })

    }

    async GetUsersSuggestions(searchText: string) {
        let results = await sp.profiles.clientPeoplePickerSearchUser({
            MaximumEntitySuggestions: 3,
            QueryString: searchText,
            PrincipalType : PrincipalType.User & PrincipalType.SecurityGroup
        })

        console.log(JSON.stringify(results));

        return results.map<IUserSuggestion>(r => {
            return {
                Email: r.EntityData.Email,
                Title: r.DisplayText
            }
        })
    }

}