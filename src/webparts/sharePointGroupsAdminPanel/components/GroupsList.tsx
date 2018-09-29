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
            name: "Title"            
        },
        {
            fieldName: "Description",
            minWidth: 100,
            key: "description",
            name: "Description"
        },
        {
            fieldName: "Owner",
            minWidth: 400,
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
            }
        },
        {
            fieldName: "Users",
            minWidth: 400,
            key: "users",
            name: "Users",
            onRender: (item: ISpGroup, index, column) => {
                return (

                    <div>
                        {item.Users.map(u => {return (
                            <div>
                                <HoverCard
                                    expandingCardProps = {{
                                        onRenderCompactCard: (props) => {
                                            return (
                                                <div>{props.renderData.Email}</div>
                                            )
                                        },
                                        onRenderExpandedCard: (props) => {
                                            return (
                                                <div>{props.renderData.Email}</div>
                                            )
                                        },
                                        renderData: u
                                    }}instantOpenOnClick = {true}>
                                    <Persona
                                        size={PersonaSize.small}
                                        imageUrl={`https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${u.Email}&UA=0&size=HR64x64`} />
                                </HoverCard>
                            </div>
                        )})}
                    </div>
                )
            }
        }
    ]

    constructor(props) {
        super(props)
    }

    public render(): React.ReactElement<IGroupsListProps> {
        return (
            <div>
                <DetailsList
                    items={this.props.groups}
                    columns = {this._columns}
                />
            </div>
        );
    }

    private _onRenderCompactCard
}
