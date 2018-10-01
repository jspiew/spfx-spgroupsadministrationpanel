import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SharePointGroupsAdminPanelWebPartStrings';
import SharePointGroupsAdminPanel from './components/SharePointGroupsAdminPanel';
import { ISharePointGroupsAdminPanelProps } from './components/ISharePointGroupsAdminPanelProps';
import { PnPSpGroupSvc } from '../../services/spGroupSvc';

export interface ISharePointGroupsAdminPanelWebPartProps {
  description: string;
}

export default class SharePointGroupsAdminPanelWebPart extends BaseClientSideWebPart<ISharePointGroupsAdminPanelWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISharePointGroupsAdminPanelProps > = React.createElement(
      SharePointGroupsAdminPanel,
      {
        groupsSvc: new PnPSpGroupSvc(this.context),
        spHttpClient: this.context.spHttpClient,
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl
      }
    );

    ReactDom.render(element, this.domElement);
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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
