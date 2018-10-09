import * as React from 'react';
import styles from './UsersPanel.module.scss';
import { ISpUser, IUsersSvc, IUserSuggestion, ISpGroupSvc, ISpGroup } from '../../../models';
import { SpUserPersona } from "./small/userDisplays"
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import PeoplePicker from "../../../components/PeoplePicker"
import { SPHttpClient } from '@microsoft/sp-http';
import { autobind } from '@uifabric/utilities/lib/autobind';
import {TextField} from "office-ui-fabric-react/lib/TextField"
import {DefaultButton} from "office-ui-fabric-react/lib/Button"
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner"
import { PersonaSize } from 'office-ui-fabric-react/lib/Persona';
export interface IUsersPanelState {
    usersToAdd: IUserSuggestion[]
    usersToRemove: ISpUser[]
    originalUsers: ISpUser[]
    usersAreBeingAdded: boolean
}

export interface IUsersPanelProps {
    isOpen: boolean
    group: ISpGroup
    users: Array<ISpUser>
    usersSvc: IUsersSvc
    addUsersToGroup: (groupId: number, user: IUserSuggestion[]) => Promise<any>
    removeUsersFromGroup: (groupId: number, users: ISpUser[]) => Promise<any>
}

export default class UsersPanel extends React.Component<IUsersPanelProps, IUsersPanelState> {
    constructor(props){
        super(props);
        this.state = {
            usersToAdd : [],
            usersToRemove : [],
            originalUsers: this.props.users,
            usersAreBeingAdded : false
        }
    }
    public render(){
        if (this.props.group == null){
            return null;
        }
        return(
            <Panel
                className = {styles.usersPanel}
                isOpen={this.props.isOpen}
                type={PanelType.medium}
                headerText={`${this.props.group.Title} members"`}
            >
                <PeoplePicker
                    svc = {this.props.usersSvc}
                    onChanged = {this._peoplePickerChanged}
                    disabled = {this.state.usersAreBeingAdded}
                />

                {this.state.usersToAdd.length > 0 && <h4 className = {styles.userToBeAddedText}>Following users will be added</h4>}
                {this.state.usersToAdd.map(u => {
                    return <SpUserPersona user={u} />
                })}
                
                {this.state.usersToRemove.length > 0 && <h4 className={styles.userToBeRemovedText}>Following users will be removed</h4>}
                {this.state.usersToRemove.map(u => {
                    return <SpUserPersona user={u} />
                })}

                {this.state.originalUsers.length > 0 && <h4>Following users will remain in group</h4>}
                {this.state.originalUsers.map(u => {
                    return <SpUserPersona user={u} />
                })}
                
                {this.state.usersAreBeingAdded && <Spinner size = {SpinnerSize.small}/>}
            </Panel>
        )
    }

    @autobind
    private _peoplePickerChanged(items: IUserSuggestion[]){
        this.setState({
            usersToAdd: [...this.state.usersToAdd, ...items]
        })
    }

    @autobind
    private async _addPeopleButtonClicked() {
        this.setState({
            usersAreBeingAdded: true
        })
        
        this.props.addUsersToGroup(this.props.group.Id,this.state.usersToAdd)

        setTimeout(() => {
            this.setState({
                usersToAdd: [],
                usersAreBeingAdded: false
            })
        },1000)
        
    }

}