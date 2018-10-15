import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { update, get, } from '@microsoft/sp-lodash-subset';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';

import { IComboBoxOption } from 'office-ui-fabric-react/lib/components/ComboBox';
import {PropertyPaneAsyncGroups} from "../../components/PropertyPaneAsyncGroups"

import * as strings from 'SharePointGroupsAdminPanelWebPartStrings';
import SharePointGroupsAdminPanel from './components/SharePointGroupsAdminPanel';
import { ISharePointGroupsAdminPanelProps } from './components/ISharePointGroupsAdminPanelProps';
import { PnPSpGroupSvc } from '../../services/spGroupSvc';
import { UserProfileUserSvc } from '../../services/userProfileUserSvc';


export enum spGroupAdminPanelViewType {
  Details,
  SimpleList,
  ExtendedList
}

export interface ISharePointGroupsAdminPanelWebPartProps {
  viewType: spGroupAdminPanelViewType
  groups: number[]
}

export default class SharePointGroupsAdminPanelWebPart extends BaseClientSideWebPart<ISharePointGroupsAdminPanelWebPartProps> {

  public render(): void {

    const webPartComponent: React.ReactElement<ISharePointGroupsAdminPanelProps> = React.createElement(
      SharePointGroupsAdminPanel,
      {
        groupsSvc: new PnPSpGroupSvc(this.context),
        selectedGroups: this.properties.groups,
        spHttpClient: this.context.spHttpClient,
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        viewType : this.properties.viewType,
        usersSvc: new UserProfileUserSvc(this.context)
      } as ISharePointGroupsAdminPanelProps
    );
    ReactDom.render(webPartComponent, this.domElement);
    
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private async _getGroupsForPropertyPane(): Promise<IComboBoxOption[]> {
    let svc = new PnPSpGroupSvc(this.context);
    let options = (await svc.GetGroupsForDropdown()).map<IComboBoxOption>(g => {return {
      key: g.Id,
      text: g.Title
    }});
    return options;
  }

  private _onGroupsChange(propertyPath: string, newValue: any): void {
    const oldValue: any = get(this.properties, propertyPath);
    // store new value in web part properties
    update(this.properties, propertyPath, (): any => { return newValue; });
    // refresh web part
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "WebPart configuration"
          },
          groups: [
            {
              groupName: "Display",
              groupFields: [
                PropertyPaneDropdown('viewType', {
                  label: "Display type",
                  selectedKey : this.properties.viewType,
                  options : [
                    {
                      key : spGroupAdminPanelViewType.SimpleList,
                      text : "Standard list"
                    },
                    {
                      key: spGroupAdminPanelViewType.ExtendedList,
                      text: "Extended list"
                    },
                    {
                      key: spGroupAdminPanelViewType.Details,
                      text: "Details"
                    }
                  ]
                }),
                new PropertyPaneAsyncGroups('groups',{
                  label: "Selected Groups",
                  loadOptions: this._getGroupsForPropertyPane.bind(this),
                  onPropertyChange: this._onGroupsChange.bind(this),
                  selectedKey: this.properties.groups || []
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
