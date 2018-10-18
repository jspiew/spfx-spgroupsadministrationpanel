import * as React from 'react';
import { TagPicker, ITag } from 'office-ui-fabric-react/lib/components/pickers/TagPicker/TagPicker';
import { Spinner } from 'office-ui-fabric-react/lib/components/Spinner';
import { sortBy } from '@microsoft/sp-lodash-subset'


export interface IAsyncGroupsPickerState {
    loading: boolean;
    options: ITag[];
    selectedOptions: ITag[];
    error: string;
}
export interface IAsyncGroupsPickerProps {
    label: string;
    loadOptions: () => Promise<ITag[]>;
    onChanged: (options: ITag[]) => void;
    selectedKey:  number[];
    disabled: boolean;
    stateKey: string;
}

export default class AsyncGroupsPicker extends React.Component<IAsyncGroupsPickerProps, IAsyncGroupsPickerState> {

    constructor(props: IAsyncGroupsPickerProps, state: IAsyncGroupsPickerState) {
        super(props);

        this.state = {
            loading: false,
            options: [],
            selectedOptions: [],
            error: undefined
        };
    }

    public componentDidMount(): void {
        this.loadOptions();
    }

    public componentDidUpdate(prevProps: IAsyncGroupsPickerProps, prevState: IAsyncGroupsPickerState): void {
        if (this.props.disabled !== prevProps.disabled ||
            this.props.stateKey !== prevProps.stateKey) {
            this.loadOptions();
        }
    }

    private loadOptions(): void {
        this.setState({
            loading: true,
            error: undefined,
            options: []
        });

        this.props.loadOptions()
            .then((options: ITag[]): void => {
                this.setState({
                    loading: false,
                    error: undefined,
                    options: sortBy(options, o => o.name),
                    selectedOptions: options.filter( o => this.props.selectedKey.indexOf(parseInt(o.key)) >= 0)
                });
            }, (error: any): void => {
                this.setState((prevState: IAsyncGroupsPickerState, props: IAsyncGroupsPickerProps): IAsyncGroupsPickerState => {
                    prevState.loading = false;
                    prevState.error = error;
                    return prevState;
                });
            });
    }
/**
 *
 *
 * @returns {JSX.Element}
 * @memberof AsyncGroupsPicker
 */
public render(): JSX.Element {
        const loading: JSX.Element = this.state.loading ? <div><Spinner label={'Loading options...'} /></div> : <div />;
        const error: JSX.Element = this.state.error !== undefined ? <div className={'ms-TextField-errorMessage ms-u-slideDownIn20'}>Error while loading items: {this.state.error}</div> : <div />;

        return (
            <div>
                <TagPicker 
                    itemLimit = {20}
                    onEmptyInputFocus = {(selected) => {return this.state.options.filter(o => selected.indexOf(o) < 0)}}
                    selectedItems = {this.state.selectedOptions}
                    onResolveSuggestions = {(filter,selectedItems) => {
                        if (!filter) return this.state.options.filter(o => selectedItems.indexOf(o) < 0);
                        return this.state.options.filter(o => o.name.toLowerCase().indexOf(filter.toLowerCase()) >= 0 && selectedItems.indexOf(o) < 0)
                    }}
                    onChange = {this.onChanged.bind(this)}
                />
                {loading}
                {error}
            </div>
        );
    }

    private onChanged(options: ITag[]): void {
        this.setState({
            selectedOptions: [...options]
        });
        if (this.props.onChanged) {
            this.props.onChanged(options);
        }
    }
}