import * as React from 'react';
import styles from './GroupsList.module.scss';
import {
    DetailsList,
    IColumn} from 'office-ui-fabric-react/lib/DetailsList';
import { ISpGroup, IUsersSvc, ISpGroupSvc, IUserSuggestion } from '../../../models';
import {SpUserPersona, SpUsersFacepile} from "../../../components/small/userDisplays"
import UsersPanel from "./UsersPanel"
import { SPHttpClient } from '@microsoft/sp-http';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { AbbrToggle } from "../../../components/small/abbrToggle"
import { Draft } from '../../../utils/draft';
import { Dialog } from '@microsoft/sp-dialog/lib/index';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import EditableSpPersona from "../../../components/EditableSpPersona"
import EditableTextField from "../../../components/EditableTextField"
import { autobind } from '@uifabric/utilities/lib';
import { Item } from '@pnp/sp';

export interface IGroupsListState {
    openGroup: ISpGroup,
    isGroupEditPanelOpen: boolean   
}

export interface IGroupsListProps {
    extendedView? : boolean,
    groups: Array<ISpGroup>,
    usersSvc: IUsersSvc,
    groupsSvc: ISpGroupSvc
}

export default class GroupsList extends React.Component<IGroupsListProps, IGroupsListState> {

    private readonly _basicColumns: Array<IColumn> = [
        {
            fieldName: "Title",
            minWidth: 100,
            maxWidth: 200,
            key: "title",
            name: "Title",
            isResizable: true,        
            onRender: (item: ISpGroup) => {
                return (
                    <EditableTextField value={item.Title} onChanged={(t) => { return this._titleChanged(item, t)}} />
                )
            }
        },
        {
            fieldName: "Description",
            minWidth: 100,
            maxWidth: 200,
            key: "description",
            name: "Description",
            isResizable: true,
            onRender: (item: ISpGroup) => {
                return (
                    <EditableTextField value={item.Description} onChanged={(t) => { return this._descriptionChanged(item, t) }} />
                )
            }
        },
        {
            fieldName: "Owner",
            minWidth: 250,
            key: "owner",
            name: "Owner",
            onRender: (item: ISpGroup) => {
                return (
                    <EditableSpPersona user = {item.Owner} svc = {this.props.usersSvc} onChanged={(u) => {return this._ownerChanged(item,u)}} />
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
                    <div>
                        <a href='#' onClick = {()=>{

                            this.setState({
                                openGroup: item
                            });

                            this.props.groupsSvc.GetUsersFromGroup(item.Id).then(users => {
                                item.Users = users;
                                this.setState({
                                    isGroupEditPanelOpen: true
                                })
                            })
                        }}>
                            Edit users
                        </a>
                        {(this.state.openGroup && this.state.openGroup.Id == item.Id) && <Spinner />}
                    </div>
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
                            this.props.groupsSvc.UpdateGroup(item.Id, {OnlyAllowMembersViewMembership : checked}).catch(() =>{
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
                        defaultValue={item.AllowMembersEditMembership}
                        onChanged={(checked) => {
                            this.props.groupsSvc.UpdateGroup(item.Id, { AllowMembersEditMembership: checked }).catch(() => {
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
            onRender: (item: ISpGroup) => {
                return (
                    <EditableTextField value={item.RequestToJoinLeaveEmailSetting} onChanged={(t) => { return this._requestEmailChanged(item, t) }} />
                )
            }
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
        return (
            <div className={styles.groupsList}>
                <DetailsList
                    items={this.props.groups}
                    columns = {this.props.extendedView ? [...this._basicColumns,...this._extendedColumns] : this._basicColumns}
                />
                <UsersPanel 
                    group  = {this.state.openGroup}
                    isOpen={this.state.isGroupEditPanelOpen}
                    usersSvc = {this.props.usersSvc}
                    addUsersToGroup = {this.props.groupsSvc.AddGroupMembers}
                    removeUsersFromGroup = {this.props.groupsSvc.RemoveGroupMembers}
                    onClose = {() => {this.setState({
                        openGroup: null,
                        isGroupEditPanelOpen: false
                    })}}
                />
            </div>
        );
    }

    @autobind
    private async _ownerChanged(group: ISpGroup, owner: IUserSuggestion) {
        try{
            let newOwner = await this.props.usersSvc.EnsureUser(owner);
            await this.props.groupsSvc.UpdateGroup(group.Id, {
                Owner: newOwner
            });
        }
        catch(e){
            alert("err");
        }
    }

    @autobind
    private async _titleChanged(group: ISpGroup, title: string) {
        try{
            await this.props.groupsSvc.UpdateGroup(group.Id, { Title: title });
        } catch(e){
            Dialog.alert(`There was an error while updating the title of the "${group.Title}" group.`)
            throw e
        }
    }

    @autobind
    private async _descriptionChanged(group: ISpGroup, title: string) {
        try {
            await this.props.groupsSvc.UpdateGroup(group.Id, { Description: title });
        } catch (e) {
            Dialog.alert(`There was an error while updating the description of the "${group.Title}" group.`)
            throw e
        }
    }

    @autobind
    private async _requestEmailChanged(group: ISpGroup, title: string) {
        try {
            await this.props.groupsSvc.UpdateGroup(group.Id, { RequestToJoinLeaveEmailSetting: title });
        } catch (e) {
            Dialog.alert(`There was an error while updating the title of the "${group.Title}" group.`)
            throw e
        }
    }

}
