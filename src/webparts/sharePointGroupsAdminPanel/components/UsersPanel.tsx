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
            originalUsers: this.props.group ? this.props.group.Users: [],
            usersAreBeingAdded : false
        }
    }

    public componentWillReceiveProps(props:IUsersPanelProps){
        this.setState({
            originalUsers: props.group ? props.group.Users : [],
        })
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
                    return <SpUserPersona user={u} onDelete = {this._removeUserToAdd} />
                })}
                
                {this.state.usersToRemove.length > 0 && <h4 className={styles.userToBeRemovedText}>Following users will be removed</h4>}
                {this.state.usersToRemove.map(u => {
                    return <SpUserPersona user={u} onDelete = {this._removeUserToRemove} />
                })}

                {this.state.originalUsers.length > 0 && <h4>Following users will remain in group</h4>}
                {this.state.originalUsers.map(u => {
                    return <SpUserPersona user={u} onDelete = {this._removeOriginalUser}/>
                })}
                
                {this.state.usersAreBeingAdded && <Spinner size = {SpinnerSize.small}/>}
            </Panel>
        )
    }

    @autobind
    private _peoplePickerChanged(items: IUserSuggestion[]){
        let itemsToAdd = items.filter(i => {return !this.isUserSelected(i)})
        this.setState({
            usersToAdd: [...this.state.usersToAdd, ...itemsToAdd]
        })
    }

    @autobind
    private _removeOriginalUser(user: ISpUser){
        let newOriginalUsers = [...this.state.originalUsers].filter(u => { u.Email.toLowerCase() !== user.Email.toLowerCase() })
        let newUsersToRemove = [...this.state.usersToRemove]
        newUsersToRemove.push(user);

        this.setState({
            usersToRemove: newUsersToRemove,
            originalUsers: newOriginalUsers
        })
    }

    @autobind
    private _removeUserToRemove(user: ISpUser) {
        let newUsersToRemove = [...this.state.usersToRemove].filter(u => { u.Email.toLowerCase() !== user.Email.toLowerCase() })
        let newOriginalUsers = [...this.state.originalUsers]
        newOriginalUsers.push(user);

        this.setState({
            usersToRemove: newUsersToRemove,
            originalUsers: newOriginalUsers
        })
    }

    @autobind
    private  _removeUserToAdd(user: ISpUser){
        let newUsersToAdd = [...this.state.usersToAdd].filter(u => { u.Email.toLowerCase() !== user.Email.toLowerCase() })
        this.setState({
            usersToAdd: newUsersToAdd
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

    private isUserSelected(user:IUserSuggestion){
        let originalEmails = this.state.originalUsers.map(u => u.Email.toLowerCase());
        let newEmails = this.state.usersToAdd.map(u => u.Email.toLowerCase());
        let oldEmails = this.state.usersToRemove.map(u => u.Email.toLowerCase());

        let allEmails = [...originalEmails,...newEmails,...oldEmails];

        return allEmails.indexOf(user.Email.toLowerCase()) >= 0
        
    }

}