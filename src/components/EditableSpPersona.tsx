import { IUserSuggestion, IUsersSvc, ISpUser } from "../models/index"
import * as React from 'react';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { autobind } from '@uifabric/utilities/lib';
import {SpUserPersona} from "./small/userDisplays"
import styles from "./EditableSpPersona.module.scss"
import PeoplePicker from "./PeoplePicker"
export interface IEditableSpPersonaState {
    isEditMode: boolean;
}

export interface IEditableSpPersonaProps {
    onChanged: (newUser: IUserSuggestion) => any
    user: IUserSuggestion
    svc: IUsersSvc
}

export default class EditableSpPersona extends React.Component<IEditableSpPersonaProps, IEditableSpPersonaState> {
    constructor(props) {
        super(props);

        this.state = {
            isEditMode: false
        };
    }

    public render() {
        if(this.state.isEditMode) return this._renderEdit();

        return (
            <div className={styles.editableSpPersona} onClick={this.state.isEditMode ? () => { } : this._onPersonaClick }>
                <div className={styles.persona}>
                    <SpUserPersona user={this.props.user}/>
                </div>
                <div className={styles.editIcon}>
                    Edit <i className="ms-Icon ms-Icon--Edit" aria-hidden="true"></i>
                </div>
            </div>
        )
    }

    private _renderEdit(){
        return (
            <PeoplePicker 
                svc = {this.props.svc}
                onChanged = {(users) => {this.props.onChanged(users[0])}}
            />)
    }

    @autobind
    private _onPersonaClick(){
        this.setState({
            isEditMode : true
        })
    }
}