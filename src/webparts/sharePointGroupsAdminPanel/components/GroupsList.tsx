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
import { ISpGroup, ISpUser } from '../../../models';
import { autobind } from '@uifabric/utilities/lib';

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
            minWidth: 100,
            key: "owner",
            name: "Owner",
            onRender: (item: ISpGroup,index,column) => {
                return (
                    <div>
                        <Persona 
                            primaryText= {item.Title}
                            size = {PersonaSize.small}
                            imageUrl={`https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${item.Owner.Email}&UA=0&size=HR64x64`} 
                            onRenderSecondaryText = {() => {
                                return <a href={item.Owner.Email}>{item.Owner.Email}</a>
                            }} />
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
}
