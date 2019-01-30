import * as React from 'react';
import styles from './GroupsDetailsView.module.scss';
import {
    DetailsList,
    IColumn
} from 'office-ui-fabric-react/lib/DetailsList';
import { ISpGroup, ISpUser } from '../../../models';
import { SpUserPersona, SpUsersFacepile } from "../../../components/small/userDisplays"
import UsersPanel from "./UsersPanel"
import { SPHttpClient } from '@microsoft/sp-http';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { AbbrToggle } from "../../../components/small/abbrToggle"
import { Draft } from '../../../utils/draft';
import { css, autobind } from '@uifabric/utilities/lib';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

export interface IGroupsDetailsViewState {
    selectedGroup: ISpGroup;
    usersToAdd: Array<ISpUser>;
}

export interface IGroupsDetailsViewProps {
    groups: Array<ISpGroup>,
    spHttpClient: SPHttpClient,
    webAbsoluteUrl: string,
    updateGroup: (groupId: number, changes: Draft<ISpGroup>) => Promise<any>
}

export default class GroupsDetailsView extends React.Component<IGroupsDetailsViewProps, IGroupsDetailsViewState> {
    constructor(props){
        super(props);
        this.state = {
            selectedGroup : null,
            usersToAdd : []
        }
    }

    private get hidePeoplePickerDialog() {
        return this.state.selectedGroup == null;
    }

    public render(): React.ReactElement<IGroupsDetailsViewProps> {
        return (
            <div>
                <div className={css("ms-Grid", styles.groupsDetailsView)}>
                    <div className="ms-Grid-row">
                        {this.props.groups.map((g) => {
                            return (
                                <div className="ms-Grid-col ms-sm4">
                                    <div>
                                        <h4 className={styles.groupTitle}>{g.Title}</h4>  
                                        <i 
                                            className={css("ms-Icon", "ms-Icon--PeopleAdd", styles.peopleAddIcon)} 
                                            onClick={() => {this._addPeopleIconClicked(g)}}
                                            aria-hidden="true"></i>
                                    </div>
                                    {g.Users.map(u => <SpUserPersona user={u}/>)}
                                </div>
                            )
                        })}
                    </div>
                </div>
                <Dialog
                    hidden = {this.hidePeoplePickerDialog}
                    onDismiss={this._closeDialog}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: `Add members to ${(this.state.selectedGroup|| {Title:"undefined"}).Title}`,
                    }}
                    modalProps={{
                        isBlocking: false,
                    }}
                >
                    <PeoplePicker
                        context={({ spHttpClient: this.props.spHttpClient, pageContext: { web: { absoluteUrl: this.props.webAbsoluteUrl } } }) as any}
                        titleText="Pick users"
                        personSelectionLimit = {20}

                        showtooltip={true}
                        selectedItems={this._getPeoplePickerItems}
                        showHiddenInUI={true}
                        principleTypes={[PrincipalType.User, PrincipalType.SecurityGroup]}
                    />
                    {this.state.usersToAdd.map(u => <SpUserPersona user={u} />)}
                </Dialog>
            </div>
        )
    }

    @autobind
    private _closeDialog(){
        this.setState({
            selectedGroup: null
        })
    }

    @autobind
    private _addPeopleIconClicked(group:ISpGroup) {
        this.setState({
            selectedGroup: group
        })
    }

    @autobind
    private _getPeoplePickerItems(items: any[]) {
        this.setState({
            usersToAdd: items.map<ISpUser>(u => {return {
                Email: u.secondaryText,
                Title: u.text,
                Id: u.id
            }})
        })
        console.log(JSON.stringify(items));
    }
}