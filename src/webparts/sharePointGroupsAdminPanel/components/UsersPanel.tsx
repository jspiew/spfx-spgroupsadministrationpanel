import * as React from 'react';
import styles from './UsersPanel.module.scss';
import { ISpUser, IUsersSvc, IUserSuggestion, ISpGroupSvc, ISpGroup } from '../../../models';
import { SpUserPersona } from "../../../components/small/userDisplays";
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import PeoplePicker from "../../../components/PeoplePicker";
import { SPHttpClient } from '@microsoft/sp-http';
import { autobind } from '@uifabric/utilities/lib/autobind';
import {TextField} from "office-ui-fabric-react/lib/TextField";
import {DefaultButton} from "office-ui-fabric-react/lib/Button";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";
import { PersonaSize } from 'office-ui-fabric-react/lib/Persona';
export interface IUsersPanelState {
    usersToAdd: IUserSuggestion[];
    usersToRemove: ISpUser[];
    originalUsers: ISpUser[];
    usersAreBeingAdded: boolean;
    errorMessage: string;
}

export interface IUsersPanelProps {
    isOpen: boolean;
    group: ISpGroup;
    usersSvc: IUsersSvc;
    onClose: () => void;
    addUsersToGroup: (groupId: number, user: IUserSuggestion[]) => Promise<any>;
    removeUsersFromGroup: (groupId: number, users: ISpUser[]) => Promise<any>;
}

export default class UsersPanel extends React.Component<IUsersPanelProps, IUsersPanelState> {
    constructor(props){
        super(props);
        this.state = {
            usersToAdd : [],
            usersToRemove : [],
            originalUsers: this.props.group ? this.props.group.Users: [],
            usersAreBeingAdded : false,
            errorMessage: null
        };
    }

    public componentWillReceiveProps(props:IUsersPanelProps){
        this.setState({
            originalUsers: props.group ? props.group.Users : [],
        });
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
                onDismiss = {this.props.onClose}
            >
                {this.state.usersAreBeingAdded && 
                <div className={styles.submitOverlay}>
                    <Spinner size={SpinnerSize.large} label="Submitting" />
                </div>}
                {this.state.errorMessage &&
                <div>
                    <h4 style = {{color:"darkred"}}>
                        An error has occured: {this.state.errorMessage}
                    </h4>
                    <DefaultButton
                        onClick = {this.props.onClose} 
                        label="OK"/>
                </div>}
                <PeoplePicker
                    svc = {this.props.usersSvc}
                    includeGroups = {false}
                    onChanged = {this._peoplePickerChanged}
                    disabled = {this.state.usersAreBeingAdded}
                />

                {this.state.usersToAdd.length > 0 && <h4 className = {styles.userToBeAddedText}>Following users will be added</h4>}
                {this.state.usersToAdd.map(u => {
                    return <SpUserPersona user={u} key={u.Email} onDelete = {this._removeUserToAdd} />;
                })}

                <h4>Following users will remain in group</h4>
                {this.state.originalUsers && this.state.originalUsers.map(u => { //I don't like this check, original users should never be undefined, something's not right in the props passed
                    return <SpUserPersona key = {u.Email} user={u} onDelete = {this._removeOriginalUser}/>;
                })}

                {this.state.usersToRemove.length > 0 && <h4 className={styles.userToBeRemovedText}>Following users will be removed</h4>}
                {this.state.usersToRemove.map(u => {
                    return <SpUserPersona key={u.Email} user={u} onDelete={this._removeUserToRemove} />;
                })}
                
                <DefaultButton 
                    text="Submit" 
                    disabled={this.state.usersToAdd.length == 0 && this.state.usersToRemove.length == 0 &&this.state.usersAreBeingAdded } 
                    onClick={this._submitChanges} /> 
                {this.state.usersAreBeingAdded && <Spinner size={SpinnerSize.small} />}
            </Panel>
        );
    }

    @autobind
    private _peoplePickerChanged(items: IUserSuggestion[]){
        let itemsToAdd = items.filter(i => {return !this.isUserSelected(i);});
        this.setState({
            usersToAdd: [...this.state.usersToAdd, ...itemsToAdd]
        });
    }

    @autobind
    private _removeOriginalUser(user: ISpUser){
        let newOriginalUsers = [...this.state.originalUsers].filter(u => u.Email.toLowerCase() !== user.Email.toLowerCase() );
        let newUsersToRemove = [...this.state.usersToRemove];
        newUsersToRemove.push(user);

        this.setState({
            usersToRemove: newUsersToRemove,
            originalUsers: newOriginalUsers
        });
    }

    @autobind
    private _removeUserToRemove(user: ISpUser) {
        let newUsersToRemove = [...this.state.usersToRemove].filter(u =>  u.Email.toLowerCase() !== user.Email.toLowerCase() );
        let newOriginalUsers = [...this.state.originalUsers];
        newOriginalUsers.push(user);

        this.setState({
            usersToRemove: newUsersToRemove,
            originalUsers: newOriginalUsers
        });
    }

    @autobind
    private  _removeUserToAdd(user: ISpUser){
        let newUsersToAdd = [...this.state.usersToAdd].filter(u =>  u.Email.toLowerCase() !== user.Email.toLowerCase() );
        this.setState({
            usersToAdd: newUsersToAdd
        });
    }

    @autobind
    private async _submitChanges() {
        this.setState({
            usersAreBeingAdded: true
        });
        
        let addPromise = this.props.addUsersToGroup(this.props.group.Id,this.state.usersToAdd);
        let removePromise = this.props.removeUsersFromGroup(this.props.group.Id, this.state.usersToRemove);

        await Promise.all([addPromise,removePromise]);
        this.setState({
            usersToAdd: [],
            usersToRemove: [],
            usersAreBeingAdded: false
        });
        this.props.onClose();
        
    }

    private isUserSelected(user:IUserSuggestion){
        let originalEmails = this.state.originalUsers.map(u => u.Email.toLowerCase());
        let newEmails = this.state.usersToAdd.map(u => u.Email.toLowerCase());
        let oldEmails = this.state.usersToRemove.map(u => u.Email.toLowerCase());

        let allEmails = [...originalEmails,...newEmails,...oldEmails];

        return allEmails.indexOf(user.Email.toLowerCase()) >= 0;
        
    }

}