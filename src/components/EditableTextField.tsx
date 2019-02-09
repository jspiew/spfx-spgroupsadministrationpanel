import { IUserSuggestion, IUsersSvc, ISpUser } from "../models/index";
import * as React from 'react';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import {TextField} from "office-ui-fabric-react/lib/TextField";
import { autobind } from '@uifabric/utilities/lib';
import {SpUserPersona} from "./small/userDisplays";
import styles from "./EditableTextField.module.scss";
import PeoplePicker from "./PeoplePicker";
export interface IEditableTextFieldState {
    isEditMode: boolean;
    text: string;
}

export interface IEditableTextFieldProps {
    onChanged: (newValue: string) => Promise<any>;
    value: string;
}

export default class EditableTextField extends React.Component<IEditableTextFieldProps, IEditableTextFieldState> {
    constructor(props) {
        super(props);

        this.state = {
            isEditMode: false,
            text: this.props.value
        };
    }

    public componentDidMount(){
        this.setState({
            isEditMode: false,
            text: this.props.value
        });
    }

    // public componentWillReceiveProps(nextProps: IEditableTextFieldProps){
    //     this.setState({
    //         isEditMode: false,
    //         text: nextProps.value
    //     });
    // }

    public render() {
        if(this.state.isEditMode) return this._renderEdit();

        return (
            <div className={styles.editableTextField} onClick={this.state.isEditMode ? () => { } : this._onFieldClick }>
                <div className={styles.text}>
                    <div className={styles.editIcon}>
                        <i className="ms-Icon ms-Icon--Edit" aria-hidden="true"></i>
                    </div>
                    {this.state.text}
                </div>
                
            </div>
        );
    }

    private _renderEdit(){
        return (
            <TextField
                defaultValue = {this.state.text}
                inputClassName={styles.editableTextBox}
                onChanged={(newVal: string) => { this.setState({
                    text: newVal
                });}}
                autoFocus= {true}
                selected = {true}
                onKeyDown = {this._onKeyDown}
                onBlur={this._onBlur}
            />);
    }

    @autobind
    private async _onBlur(){
        if (this.props.value !== this.state.text){
            try{
                await this.props.onChanged(this.state.text);
                this.setState({
                    isEditMode: false
                });
            } catch(e) {
                this.setState({
                    isEditMode: false,
                    text: this.props.value //if update failed, revert to initial value 
                });
            }
        } else {
            this.setState({
                isEditMode: false
            });
        }
    }

    @autobind
    private _onFieldClick(){
        this.setState({
            isEditMode : true
        });
    }

    @autobind
    private _onKeyDown(event:React.KeyboardEvent<HTMLInputElement|HTMLTextAreaElement>){
        switch(event.keyCode){
            case 13: this._onBlur();break;
            case 27: {
                this.setState({
                    isEditMode: false,
                    text: this.props.value //if update failed, revert to initial value 
                });
            }
        }
    }
}