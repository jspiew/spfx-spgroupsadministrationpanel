import * as React from 'react';
import styles from './UsersPanel.module.scss';
import { ISpUser, IUsersSvc } from '../../../models';
import { SpUserPersona } from "./small/userDisplays"
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import PeoplePicker from "../../../components/PeoplePicker"
import { SPHttpClient } from '@microsoft/sp-http';
import { autobind } from '@uifabric/utilities/lib/autobind';
import {TextField} from "office-ui-fabric-react/lib/TextField"
export interface IUsersPanelState {
    selectedUsers: any[]
}

export interface IUsersPanelProps {
    isOpen: boolean
    groupTitle: string
    users: Array<ISpUser>
    usersSvc: IUsersSvc
}

export default class UsersPanel extends React.Component<IUsersPanelProps, IUsersPanelState> {
    public render(){
        return(
            <Panel
                isOpen={this.props.isOpen}
                type={PanelType.medium}
                headerText={`Edit ${this.props.groupTitle} members"`}
            >
                <PeoplePicker
                    svc = {this.props.usersSvc}
                    onChanged = {(items) => {
                        console.log("success")
                    }}
                />
            </Panel>
        )
    }

}