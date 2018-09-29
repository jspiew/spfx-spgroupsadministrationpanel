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
    return (
      <div className={ styles.sharePointGroupsAdminPanel }>
      {this.state.areGroupsLoading && "LOADING"}
      { this.state.groups && 
        <GroupList 
          groups = {this.state.groups}
        />
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
