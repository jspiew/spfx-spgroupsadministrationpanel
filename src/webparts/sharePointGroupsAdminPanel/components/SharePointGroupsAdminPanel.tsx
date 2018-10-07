import * as React from 'react';
import styles from './SharePointGroupsAdminPanel.module.scss';
import { ISharePointGroupsAdminPanelProps } from './ISharePointGroupsAdminPanelProps';
import { escape } from '@microsoft/sp-lodash-subset';
import GroupList from "./GroupsList"
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  IColumn,
  IDetailsList
} from 'office-ui-fabric-react/lib/DetailsList';
import { ISpGroup, ISpUser } from '../../../models';
import { autobind } from '@uifabric/utilities/lib';
import { spGroupAdminPanelViewType } from '../SharePointGroupsAdminPanelWebPart';
import {Spinner} from "office-ui-fabric-react/lib/Spinner"
import GroupsDetailsView from "./GroupsDetailsView"

export interface ISharePointGroupsAdminPanelState {
  groups: Array<ISpGroup>
  areGroupsLoading: boolean
}

export default class SharePointGroupsAdminPanel extends React.Component<ISharePointGroupsAdminPanelProps, ISharePointGroupsAdminPanelState> {
  

  
  constructor(props){
    super(props)
    this.state = {
      groups: [],
      areGroupsLoading: false
    }
  }

  public componentDidMount() {
    this._loadGroups();
  }

  public render(): React.ReactElement<ISharePointGroupsAdminPanelProps> {
    let groupDisplay: JSX.Element = null;

    switch(this.props.viewType) {
      case spGroupAdminPanelViewType.Details: 
        groupDisplay = <GroupsDetailsView
          groups={this.state.groups}
          spHttpClient={this.props.spHttpClient}
          webAbsoluteUrl={this.props.webAbsoluteUrl}
          updateGroup={this.props.groupsSvc.UpdateGroup}
        />
        break;
      case spGroupAdminPanelViewType.ExtendedList:
        groupDisplay = <GroupList
          groups={this.state.groups}
          spHttpClient={this.props.spHttpClient}
          webAbsoluteUrl={this.props.webAbsoluteUrl}
          updateGroup={this.props.groupsSvc.UpdateGroup}
          extendedView={true}
        />
        break;
      default:
        groupDisplay = <GroupList
          groups={this.state.groups}
          spHttpClient={this.props.spHttpClient}
          webAbsoluteUrl={this.props.webAbsoluteUrl}
          updateGroup={this.props.groupsSvc.UpdateGroup}
          extendedView={false}
        />
        break;
    }

    return (
      <div className={ styles.sharePointGroupsAdminPanel }>
      {this.state.areGroupsLoading && <Spinner label="Loading groups"/>}
      { this.state.groups && 
        groupDisplay
      }
      </div>
    );
  }


  @autobind
  private _loadGroups() {
    this.setState({
      areGroupsLoading: true
    });
    this.props.groupsSvc.GetGroups().then((groups) => {
      this.setState({
        groups: groups,
        areGroupsLoading: false
      });
    });
  }
}
