import {
    CompactPeoplePicker,
    IBasePickerSuggestionsProps,
    IBasePicker,
    ListPeoplePicker,
    NormalPeoplePicker,
    ValidationState
} from 'office-ui-fabric-react/lib/Pickers';
import { IUserSuggestion, IUsersSvc } from "../models/index";
import * as React from 'react';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { autobind } from '@uifabric/utilities/lib';
export interface IPeoplePickerState {
    peopleList: IPersonaProps[];
    mostRecentlyUsed: IPersonaProps[];
    currentSelectedItems?: IPersonaProps[];
}

export interface IPeoplePickerProps {
    svc: IUsersSvc;
    includeGroups: boolean;
    onChanged: (selectedUsers: Array<IUserSuggestion>) => void;
    disabled?: boolean;
}

export default class PeoplePicker extends React.Component<IPeoplePickerProps, IPeoplePickerState> {
    constructor(props) {
        super(props);

        this.state = {
            peopleList: [],
            mostRecentlyUsed: [],
            currentSelectedItems: []
        };
    }

    public render() {
        return (
            <NormalPeoplePicker
                onResolveSuggestions={this._onFilterChanged}
                onEmptyInputFocus={this._returnMostRecentlyUsed}
                getTextFromItem={this._getTextFromItem}
                className={'ms-PeoplePicker'}
                disabled = {this.props.disabled}
                pickerSuggestionsProps={this._suggestionProps}
                key={'list'}
                selectedItems = {this.state.peopleList}
                inputProps = {{
                    placeholder : "Select user"
                }}
                onRemoveSuggestion={this._onRemoveSuggestion}
                onValidateInput={this._validateInput}
                resolveDelay={300}
                onInputChange={this._onInputChange}
                onChange = {(items) => {
                    this._onItemsChange(items); 
                    this.props.onChanged(items.map(i => {return {Email: i.secondaryText, Title: i.text};})); this._onRemoveSuggestion(items[0]);}}
            />
        );
    }

    @autobind
    private _onItemsChange(items: any[]){
        this.setState({
            currentSelectedItems: items
        });
    }

    @autobind
    private _onRemoveSuggestion (item: IPersonaProps) {
        const { peopleList, mostRecentlyUsed: mruState } = this.state;
        const indexPeopleList: number = peopleList.indexOf(item);
        const indexMostRecentlyUsed: number = mruState.indexOf(item);

        if (indexPeopleList >= 0) {
            const newPeople: IPersonaProps[] = peopleList
                .slice(0, indexPeopleList)
                .concat(peopleList.slice(indexPeopleList + 1));
            this.setState({ peopleList: newPeople });
        }

        if (indexMostRecentlyUsed >= 0) {
            const newSuggestedPeople: IPersonaProps[] = mruState
                .slice(0, indexMostRecentlyUsed)
                .concat(mruState.slice(indexMostRecentlyUsed + 1));
            this.setState({ mostRecentlyUsed: newSuggestedPeople });
        }
    }
    
    @autobind
    private async _onFilterChanged(
        filterText: string,
        currentPersonas: IPersonaProps[],
        limitResults?: number
    ) {
        if (filterText) {
            let filteredPersonas: IPersonaProps[] = (await this.props.svc.GetUsersSuggestions(filterText, this.props.includeGroups)).map<IPersonaProps>(s => {
                return {
                    text: s.Title,
                    secondaryText: s.Email,
                    imageUrl: `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${s.Email}&UA=0&size=HR64x64`
                };
            });

            filteredPersonas = this._removeDuplicates(filteredPersonas, currentPersonas);
            filteredPersonas = limitResults ? filteredPersonas.splice(0, limitResults) : filteredPersonas;
            return this._filterPromise(filteredPersonas);
        } else {
            return [];
        }
    }
    
    @autobind
    private _returnMostRecentlyUsed(currentPersonas: IPersonaProps[]) {
        let { mostRecentlyUsed } = this.state;
        mostRecentlyUsed = this._removeDuplicates(mostRecentlyUsed, currentPersonas);
        return this._filterPromise(mostRecentlyUsed);
    }


    private _filterPromise(personasToReturn: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> {
        return personasToReturn;

    }

    private _listContainsPersona(persona: IPersonaProps, personas: IPersonaProps[]) {
        if (!personas || !personas.length || personas.length === 0) {
            return false;
        }
        return personas.filter(item => item.text === persona.text).length > 0;
    }

    @autobind
    private _removeDuplicates(personas: IPersonaProps[], possibleDupes: IPersonaProps[]) {
        return personas.filter(persona => !this._listContainsPersona(persona, possibleDupes));
    }

    private _validateInput = (input: string): ValidationState => {
        if (input.indexOf('@') !== -1) {
            return ValidationState.valid;
        } else if (input.length > 1) {
            return ValidationState.warning;
        } else {
            return ValidationState.invalid;
        }
    }

    /**
     * Takes in the picker input and modifies it in whichever way
     * the caller wants, i.e. parsing entries copied from Outlook (sample
     * input: "Aaron Reid <aaron>").
     *
     * @param input The text entered into the picker.
     */
    private _onInputChange(input: string): string {
        const outlookRegEx = /<.*>/g;
        const emailAddress = outlookRegEx.exec(input);

        if (emailAddress && emailAddress[0]) {
            return emailAddress[0].substring(1, emailAddress[0].length - 1);
        }

        return input;
    }

    private _getTextFromItem(persona: IPersonaProps): string {
        return persona.text as string;
    }

    private _suggestionProps: IBasePickerSuggestionsProps = {
        suggestionsHeaderText: 'Suggested People',
        mostRecentlyUsedHeaderText: 'Suggested Contacts',
        noResultsFoundText: 'No results found',
        loadingText: 'Loading',
        showRemoveButtons: true,
        suggestionsAvailableAlertText: 'People Picker Suggestions available',
        suggestionsContainerAriaLabel: 'Suggested contacts'
    };
}