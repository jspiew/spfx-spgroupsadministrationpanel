import { ISpGroupSvc } from "../../../models";
import { SPHttpClient } from "@microsoft/sp-http";
import {spGroupAdminPanelViewType} from "../SharePointGroupsAdminPanelWebPart"

export interface ISharePointGroupsAdminPanelProps {
  groupsSvc: ISpGroupSvc;
  spHttpClient: SPHttpClient;
  webAbsoluteUrl: string;
  viewType: spGroupAdminPanelViewType
}
