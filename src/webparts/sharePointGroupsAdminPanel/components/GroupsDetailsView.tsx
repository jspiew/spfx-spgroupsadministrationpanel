import * as React from 'react';
import styles from './GroupsDetailsView.module.scss';
import {
    DetailsList,
    IColumn
} from 'office-ui-fabric-react/lib/DetailsList';
import { ISpGroup } from '../../../models';
import { SpUserPersona, SpUsersFacepile } from "./small/userDisplays"
import UsersPanel from "./UsersPanel"
import { SPHttpClient } from '@microsoft/sp-http';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { AbbrToggle } from "./small/abbrToggle"
import { Draft } from '../../../utils/draft';
import { Dialog } from '@microsoft/sp-dialog/lib/index';
import { css } from '@uifabric/utilities/lib';

export interface IGroupsDetailsViewState {
    
}

export interface IGroupsDetailsViewProps {
    groups: Array<ISpGroup>,
    spHttpClient: SPHttpClient,
    webAbsoluteUrl: string,
    updateGroup: (groupId: number, changes: Draft<ISpGroup>) => Promise<any>
}

export default class GroupsDetailsView extends React.Component<IGroupsDetailsViewProps, IGroupsDetailsViewState> {
    public render(): React.ReactElement<IGroupsDetailsViewProps> {
        return (
            <div className={css("ms-Grid", styles.groupsDetailsView)}>
                <div className="ms-Grid-row">
                    {this.props.groups.map((g) => {
                        return (
                            <div className="ms-Grid-col ms-sm4">
                                <div>
                                    <h4 className={styles.groupTitle}>{g.Title}</h4>  
                                    <i className={css("ms-Icon", "ms-Icon--PeopleAdd", styles.peopleAddIcon)} aria-hidden="true"></i>
                                </div>
                                {g.Users.map(u => <SpUserPersona user={u}/>)}
                            </div>
                        )
                    })}
                </div>
            </div>
        )
    }
}