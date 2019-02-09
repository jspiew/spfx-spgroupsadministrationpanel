import { IUserSuggestion, IUsersSvc, ISpUser } from "../models/index";
import * as React from 'react';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { autobind } from '@uifabric/utilities/lib';
import {SpUserPersona} from "./small/userDisplays";
import styles from "./EditableSpPersona.module.scss";
import PeoplePicker from "./PeoplePicker";
export interface IEditableSpPersonaState {
    isEditMode: boolean;
    user: IUserSuggestion;
}

export interface IEditableSpPersonaProps {
    onChanged: (newUser: IUserSuggestion) => any;
    includeGroups: boolean;
    user: IUserSuggestion;
    svc: IUsersSvc;
}

export default class EditableSpPersona extends React.Component<IEditableSpPersonaProps, IEditableSpPersonaState> {
    constructor(props) {
        super(props);

        this.state = {
            isEditMode: false,
            user: this.props.user
        };
    }

    public componentDidMount(){
        this.setState({
            isEditMode: false,
            user: this.props.user
        });
    }

    public componentWillReceiveProps(nextProps: IEditableSpPersonaProps){
        this.setState({
            isEditMode: false,
            user: nextProps.user
        });
    }

    public render() {
        if(this.state.isEditMode) return this._renderEdit();

        return (
            <div className={styles.editableSpPersona} onClick={this.state.isEditMode ? () => { } : this._onPersonaClick }>
                <div className={styles.persona}>
                    <SpUserPersona user={this.state.user}/>
                </div>
                <div className={styles.editIcon}>
                    Edit <i className="ms-Icon ms-Icon--Edit" aria-hidden="true"></i>
                </div>
            </div>
        );
    }

    private _renderEdit(){
        return (
            <PeoplePicker 
                includeGroups = {this.props.includeGroups}
                svc = {this.props.svc}
                onChanged={(users) => { this._onChange(users[0]);}}
            />);
    }

    private async _onChange(u: IUserSuggestion){
        try{
            await this.props.onChanged(u);
            this.setState({
                user: u,
                isEditMode: false
            });
        } catch(e) {
            this.setState({
                isEditMode: false
            });
        }
    }

    @autobind
    private _onPersonaClick(){
        this.setState({
            isEditMode : true
        });
    }
}