import * as React from 'react';
import styles from './GroupsList.module.scss';
import {
    DetailsList,
    IColumn} from 'office-ui-fabric-react/lib/DetailsList';
import { ISpGroup } from '../../../models';
import {SpUserPersona, SpUsersFacepile} from "./small/userDisplays"
import UsersPanel from "./UsersPanel"
import { SPHttpClient } from '@microsoft/sp-http';

export interface IGroupsListState {
    openGroup: ISpGroup,
    isGroupEditPanelOpen: boolean
}

export interface IGroupsListProps {
    groups: Array<ISpGroup>,
    spHttpClient: SPHttpClient,
    webAbsoluteUrl: string
}

export default class GroupsList extends React.Component<IGroupsListProps, IGroupsListState> {

    private readonly _columns: Array<IColumn> = [
        {
            fieldName: "Title",
            minWidth: 100,
            key: "title",
            name: "Title"    ,
            isResizable: true,        
        },
        {
            fieldName: "Description",
            minWidth: 100,
            key: "description",
            name: "Description",
            isResizable: true,
        },
        {
            fieldName: "Owner",
            minWidth: 200,
            key: "owner",
            name: "Owner",
            onRender: (item: ISpGroup) => {
                return (
                    <SpUserPersona user = {item.Owner} />
                )
            },
            isResizable: true,
        },
        {
            fieldName: "Users",
            minWidth: 200,
            isResizable: true,
            key: "users",
            name: "Members",
            onRender: (item: ISpGroup) => {
                return (
                    <a href='#' onClick = {()=>{
                        this.setState({
                            openGroup: item,
                            isGroupEditPanelOpen: true
                        })
                    }}>
                        Edit users
                    </a>
                )
            },
        },
        {
            fieldName: "OnlyAllowMembersViewMembership",
            minWidth: 100,
            key: "OnlyAllowMembersViewMembership",
            name: "View members right",
            isResizable: true,
        },
        {
            fieldName: "AllowMembersEditMembership",
            minWidth: 100,
            key: "AllowMembersEditMembership",
            name: "Allow Members Edit Membership",
            isResizable: true,
        },
        {
            fieldName: "RequestToJoinLeaveEmailSetting",
            minWidth: 100,
            key: "RequestToJoinLeaveEmailSetting",
            name: "Requests email",
            isResizable: true,
        }
    ]

    constructor(props: IGroupsListProps) {
        super(props)
        this.state = {
            openGroup: null,
            isGroupEditPanelOpen: false
        }
    }

    public render(): React.ReactElement<IGroupsListProps> {
        let openGroup = this.state.openGroup;
        return (
            <div className={styles.groupsList}>
                <DetailsList
                    items={this.props.groups}
                    columns = {this._columns}
                />
                <UsersPanel 
                    groupTitle = {openGroup == null ? "Undefined" : openGroup.Title}
                    isOpen={this.state.isGroupEditPanelOpen}
                    users = {openGroup == null ? [] : openGroup.Users}
                    spHttpClient = {this.props.spHttpClient}
                    webAbsoluteUrl = {this.props.webAbsoluteUrl}
                />
            </div>
        );
    }


}
