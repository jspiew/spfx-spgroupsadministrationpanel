import * as React from 'react';
import styles from './GroupsList.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import {
    DetailsList,
    DetailsListLayoutMode,
    Selection,
    IColumn,
    IDetailsList
} from 'office-ui-fabric-react/lib/DetailsList';
import { IPersonaSharedProps, Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { IFacepileProps, Facepile, OverflowButtonType, IFacepilePersona } from 'office-ui-fabric-react/lib/Facepile';
import { ISpGroup, ISpUser } from '../../../models';
import { autobind } from '@uifabric/utilities/lib';
import { HoverCard, IExpandingCardProps } from 'office-ui-fabric-react/lib/HoverCard';

export interface IGroupsListState {
    
}

export interface IGroupsListProps {
    groups: Array<ISpGroup>
}

export default class GroupsList extends React.Component<IGroupsListProps, IGroupsListState> {

    private readonly _columns: Array<IColumn> = [
        {
            fieldName: "Title",
            minWidth: 100,
            key: "title",
            name: "Title"    ,
            isResizable: true,        
        },
        {
            fieldName: "Description",
            minWidth: 100,
            key: "description",
            name: "Description",
            isResizable: true,
        },
        {
            fieldName: "Owner",
            minWidth: 200,
            key: "owner",
            name: "Owner",
            onRender: (item: ISpGroup,index,column) => {
                return (
                    <Persona 
                        primaryText= {item.Owner.Title}
                        size = {PersonaSize.small}
                        imageUrl={`https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${item.Owner.Email}&UA=0&size=HR64x64`} 
                        onRenderSecondaryText = {() => {
                            return <a href={item.Owner.Email}>{item.Owner.Email}</a>
                        }} />
                )
            },
            isResizable: true,
        },
        {
            fieldName: "Users",
            minWidth: 200,
            isResizable: true,
            key: "users",
            name: "Members",
            onRender: (item: ISpGroup, index, column) => {
                return (
                    <Facepile
                        personas={item.Users.slice(0, 10).map<IFacepilePersona>(u => {
                            return { 
                                imageUrl: `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${u.Email}&UA=0&size=HR64x64`,
                                personaName: u.Title
                            }
                        })}
                        overflowPersonas={item.Users.slice(10).map<IFacepilePersona>(u => {
                            return {
                                imageUrl: `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${u.Email}&UA=0&size=HR64x64`,
                                personaName: u.Title
                            }
                        })}
                    />
                )
            },
        },
        {
            fieldName: "OnlyAllowMembersViewMembership",
            minWidth: 100,
            key: "OnlyAllowMembersViewMembership",
            name: "View members right",
            isResizable: true,
        },
        {
            fieldName: "AllowMembersEditMembership",
            minWidth: 100,
            key: "AllowMembersEditMembership",
            name: "Allow Members Edit Membership",
            isResizable: true,
        },
        {
            fieldName: "RequestToJoinLeaveEmailSetting",
            minWidth: 100,
            key: "RequestToJoinLeaveEmailSetting",
            name: "Requests email",
            isResizable: true,
        },
    ]

    constructor(props) {
        super(props)
    }

    public render(): React.ReactElement<IGroupsListProps> {
        return (
            <div className={styles.groupsList}>
                <DetailsList
                    items={this.props.groups}
                    columns = {this._columns}
                />
            </div>
        );
    }

    private _onRenderCompactCard
}
