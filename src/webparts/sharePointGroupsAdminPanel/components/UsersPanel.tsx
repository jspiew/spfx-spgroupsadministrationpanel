import * as React from 'react';
import styles from './UsersPanel.module.scss';
import { ISpUser, IUsersSvc, IUserSuggestion } from '../../../models';
import { SpUserPersona } from "./small/userDisplays"
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import PeoplePicker from "../../../components/PeoplePicker"
import { SPHttpClient } from '@microsoft/sp-http';
import { autobind } from '@uifabric/utilities/lib/autobind';
import {TextField} from "office-ui-fabric-react/lib/TextField"
import {DefaultButton} from "office-ui-fabric-react/lib/Button"
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner"
export interface IUsersPanelState {
    selectedUsers: IUserSuggestion[],
    usersAreBeingAdded: boolean
}

export interface IUsersPanelProps {
    isOpen: boolean
    groupTitle: string
    users: Array<ISpUser>
    usersSvc: IUsersSvc
}

export default class UsersPanel extends React.Component<IUsersPanelProps, IUsersPanelState> {
    constructor(props){
        super(props);
        this.state = {
            selectedUsers : [],
            usersAreBeingAdded : false
        }
    }
    public render(){
        return(
            <Panel
                className = {styles.usersPanel}
                isOpen={this.props.isOpen}
                type={PanelType.medium}
                headerText={`${this.props.groupTitle} members"`}
            >
                <PeoplePicker
                    svc = {this.props.usersSvc}
                    onChanged = {this._peoplePickerChanged}
                    disabled = {this.state.usersAreBeingAdded}
                    selectedItems = {this.state.selectedUsers}
                />
                <DefaultButton 
                    text="Add"
                    onClick = {this._addPeopleButtonClicked}
                    className = {styles.addPeopleButton}
                    disabled = {this.state.selectedUsers.length == 0}
                    iconProps = {{
                        iconName: "PeopleAdd"
                    }}/>
                {this.state.usersAreBeingAdded && <Spinner size = {SpinnerSize.small}/>}
            </Panel>
        )
    }

    @autobind
    private _peoplePickerChanged(items: IUserSuggestion[]){
        this.setState({
            selectedUsers: items
        })
    }

    @autobind
    private _addPeopleButtonClicked() {
        this.setState({
            usersAreBeingAdded: true
        })
        setTimeout(() => {
            this.setState({
                selectedUsers: [],
                usersAreBeingAdded: false
            })
        },1000)
        
    }

}