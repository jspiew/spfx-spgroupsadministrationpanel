import { ISpGroupSvc, IUsersSvc } from "../../../models";
import { SPHttpClient } from "@microsoft/sp-http";
import {spGroupAdminPanelViewType} from "../SharePointGroupsAdminPanelWebPart"

export interface ISharePointGroupsAdminPanelProps {
  groupsSvc: ISpGroupSvc;
  usersSvc: IUsersSvc;
  spHttpClient: SPHttpClient;
  webAbsoluteUrl: string;
  viewType: spGroupAdminPanelViewType
}
