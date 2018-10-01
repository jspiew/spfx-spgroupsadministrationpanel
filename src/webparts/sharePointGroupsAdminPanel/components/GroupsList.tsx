import * as React from 'react';
import styles from './GroupsList.module.scss';
import {
    DetailsList,
    IColumn} from 'office-ui-fabric-react/lib/DetailsList';
import { ISpGroup } from '../../../models';
import {SpUserPersona, SpUsersFacepile} from "./small/userDisplays"

export interface IGroupsListState {
    
}

export interface IGroupsListProps {
    groups: Array<ISpGroup>
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
                    <SpUsersFacepile users={item.Users} />
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

    constructor(props) {
        super(props)
    }

    public render(): React.ReactElement<IGroupsListProps> {
        return (
            <div className={styles.groupsList}>
                <DetailsList
                    items={this.props.groups}
                    columns = {this._columns}
                />
            </div>
        );
    }


}
