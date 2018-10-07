import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';

import * as strings from 'SharePointGroupsAdminPanelWebPartStrings';
import SharePointGroupsAdminPanel from './components/SharePointGroupsAdminPanel';
import { ISharePointGroupsAdminPanelProps } from './components/ISharePointGroupsAdminPanelProps';
import { PnPSpGroupSvc } from '../../services/spGroupSvc';


export enum spGroupAdminPanelViewType {
  Details,
  SimpleList,
  ExtendedList
}

export interface ISharePointGroupsAdminPanelWebPartProps {
  viewType: spGroupAdminPanelViewType
}

export default class SharePointGroupsAdminPanelWebPart extends BaseClientSideWebPart<ISharePointGroupsAdminPanelWebPartProps> {

  public render(): void {

    const webPartComponent: React.ReactElement<ISharePointGroupsAdminPanelProps> = React.createElement(
      SharePointGroupsAdminPanel,
      {
        groupsSvc: new PnPSpGroupSvc(this.context),
        spHttpClient: this.context.spHttpClient,
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        viewType : this.properties.viewType
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
