import { ISpGroupSvc } from "../../../models";
import { SPHttpClient } from "@microsoft/sp-http";

export interface ISharePointGroupsAdminPanelProps {
  groupsSvc: ISpGroupSvc;
  spHttpClient: SPHttpClient;
  webAbsoluteUrl: string;
  extendedView: boolean
}
