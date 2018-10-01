import * as React from 'react';
import styles from './UsersPanel.module.scss';
import { ISpUser } from '../../../models';
import { SpUserPersona } from "./small/userDisplays"
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { SPHttpClient } from '@microsoft/sp-http';
import { autobind } from '@uifabric/utilities/lib/autobind';
export interface IUsersPanelState {
    selectedUsers: any[]
}

export interface IUsersPanelProps {
    isOpen: boolean
    groupTitle: string
    users: Array<ISpUser>
    webAbsoluteUrl: string,
    spHttpClient: SPHttpClient
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
                    context={({spHttpClient: this.props.spHttpClient, pageContext: {web : { absoluteUrl: this.props.webAbsoluteUrl}}}) as any}
                    titleText="People Picker"
                    personSelectionLimit={1}
                    showtooltip={true}
                    isRequired={true}
                    selectedItems={this._getPeoplePickerItems}
                    showHiddenInUI={false}
                    principleTypes={[PrincipalType.User, PrincipalType.SecurityGroup]}
                />
                {this.props.users.map(u => {return <SpUserPersona key={u.Email} user = {u} />})}
            </Panel>
        )
    }
    
    @autobind
    private _getPeoplePickerItems(items: any[]) {
        this.setState({
            selectedUsers: items
        })
        console.log(JSON.stringify(items));
    }
}