import { sp, PrincipalType } from "@pnp/sp";
import { IUsersSvc, ISpGroup, ISpUser, IUserSuggestion } from "../models";
import { Draft } from "../utils/draft";
import { BaseWebPartContext } from "@microsoft/sp-webpart-base";


export class SpUserSvc implements IUsersSvc {

    constructor(ctx: BaseWebPartContext) {
        sp.setup({
            spfxContext: ctx
        });

    }

    public async GetUsersSuggestions(searchText: string, includeSiteGroups = false) {
        
        
        let profileResultsPromise = sp.profiles.clientPeoplePickerSearchUser({
            MaximumEntitySuggestions: 3,
            QueryString: searchText,
            PrincipalType: includeSiteGroups ? PrincipalType.User & PrincipalType.SecurityGroup & PrincipalType.SharePointGroup : PrincipalType.User & PrincipalType.SecurityGroup
        });

        let siteGroupResults =includeSiteGroups ? await this.searchGroups(searchText) : [];
        

        let profileResults = await profileResultsPromise;

        let results = [...profileResults.map<IUserSuggestion>(r => {
            return {
                Email: r.EntityData.Email,
                Title: r.DisplayText
            };
        }), ...siteGroupResults];
        return results as (IUserSuggestion|ISpUser)[];
    }

    private async searchGroups(searchText: string) {
        let groups = await sp.web.siteGroups.select("Title,Id").usingCaching().get(); //default 30 seconds caching is perfectly fine in this case
        let matchedGroups = groups.filter(g => (g.Title as string).toLowerCase().indexOf(searchText) >= 0);
        return matchedGroups.map<ISpUser>(g => {return {
            Email: "",
            Id: g.Id,
            Title: g.Title
        };});
    }

    public async EnsureUser(suggestion: IUserSuggestion): Promise<ISpUser>{
        let ensuredUser = await sp.web.ensureUser(suggestion.Email);
        return {
            Email: ensuredUser.data.Email,
            Id: ensuredUser.data.Id,
            Title: ensuredUser.data.Title
        };
    }

}