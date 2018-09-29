import { sp } from "@pnp/sp"
import { Draft } from "../utils/draft"

export interface ISpGroupSvc {
    GetGroups: () => Promise<Array<ISpGroup>>,
    UpdateGroup: (group: Draft<ISpGroup>) => Promise<void>;
    AddGroup: (group: ISpGroup) => Promise<ISpGroup>;
    DeleteGroup: (group: ISpGroup) => Promise<void>
}

export interface ISpGroup {
    Title: string,
    Id: number,
    Members: Array<ISpUser>
    Owner: ISpUser

}

export interface ISpUser{
    Title: string,
    Id: number,
    Email: string
}



