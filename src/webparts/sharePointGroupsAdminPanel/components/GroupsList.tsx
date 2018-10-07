import * as React from 'react';
import styles from './GroupsList.module.scss';
import {
    DetailsList,
    IColumn} from 'office-ui-fabric-react/lib/DetailsList';
import { ISpGroup } from '../../../models';
import {SpUserPersona, SpUsersFacepile} from "./small/userDisplays"
import UsersPanel from "./UsersPanel"
import { SPHttpClient } from '@microsoft/sp-http';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import {AbbrToggle} from "./small/abbrToggle"
import { Draft } from '../../../utils/draft';
import { Dialog } from '@microsoft/sp-dialog/lib/index';

export interface IGroupsListState {
    openGroup: ISpGroup,
    isGroupEditPanelOpen: boolean   
}

export interface IGroupsListProps {
    extendedView? : boolean,
    groups: Array<ISpGroup>,
    spHttpClient: SPHttpClient,
    webAbsoluteUrl: string,
    updateGroup: (groupId: number, changes: Draft<ISpGroup>) => Promise<any>
}

export default class GroupsList extends React.Component<IGroupsListProps, IGroupsListState> {

    private readonly _basicColumns: Array<IColumn> = [
        {
            fieldName: "Title",
            minWidth: 100,
            maxWidth: 200,
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
            minWidth: 80,
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
        }
    ]

    private readonly _extendedColumns = [
        {
            fieldName: "OnlyAllowMembersViewMembership",
            minWidth: 100,
            key: "OnlyAllowMembersViewMembership",
            name: "View members right",
            isResizable: true,
            onRender: (item:ISpGroup) => {
                return (
                    <AbbrToggle 
                        offAbbrText= "Anyone with site access can view group members"
                        onAbbrText= "Only group members can view other members"
                        defaultValue = {item.OnlyAllowMembersViewMembership}
                        onChanged = {(checked) => {
                            this.props.updateGroup(item.Id, {OnlyAllowMembersViewMembership : checked}).catch(() =>{
                                Dialog.alert(`There was an error while updating the "View members right" property of the "${item.Title}" group.`)
                            })
                        }}
                    />
                )
            }
        },
        {
            fieldName: "AllowMembersEditMembership",
            minWidth: 100,
            key: "AllowMembersEditMembership",
            name: "Allow Members Edit Membership",
            isResizable: true,
            onRender: (item: ISpGroup) => {
                return (
                    <AbbrToggle
                        offAbbrText="Only group owner can edit members"
                        onAbbrText="Members can edit other group members"
                        defaultValue={item.OnlyAllowMembersViewMembership}
                        onChanged={(checked) => {
                            this.props.updateGroup(item.Id, { AllowMembersEditMembership: checked }).catch(() => {
                                Dialog.alert(`There was an error while updating the "Allow Members Edit Membership" property of the "${item.Title}" group.`)
                            })
                        }}
                    />
                )
            }
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
                    columns = {this.props.extendedView ? [...this._basicColumns,...this._extendedColumns] : this._basicColumns}
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
