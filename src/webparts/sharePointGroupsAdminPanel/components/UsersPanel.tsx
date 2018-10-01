import * as React from 'react';
import styles from './UsersPanel.module.scss';
import { ISpUser } from '../../../models';
import { SpUserPersona } from "./small/userDisplays"
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
export interface IUsersPanelState {

}

export interface IUsersPanelProps {
    isOpen: boolean
    groupTitle: string
    users: Array<ISpUser>
}

export default class UsersPanel extends React.Component<IUsersPanelProps, IUsersPanelState> {
    public render(){
        return(
            <Panel
                isOpen={this.props.isOpen}
                type={PanelType.medium}
                headerText={`Edit ${this.props.groupTitle} members"`}
            >
                {this.props.users.map(u => {return <SpUserPersona key={u.Email} user = {u} />})}
            </Panel>
        )
    }
}