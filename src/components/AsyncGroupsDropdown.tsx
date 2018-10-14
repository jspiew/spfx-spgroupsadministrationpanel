import * as React from 'react';
import { ComboBox, IComboBoxOption } from 'office-ui-fabric-react/lib/components/ComboBox';
import { Spinner } from 'office-ui-fabric-react/lib/components/Spinner';
import { sortBy } from '@microsoft/sp-lodash-subset'

export interface IAsyncGroupsDropdownsState {
    loading: boolean;
    options: IComboBoxOption[];
    error: string;
}
export interface IAsyncGroupsDropdownsProps {
    label: string;
    loadOptions: () => Promise<IComboBoxOption[]>;
    onChanged: (options: IComboBoxOption[], index?: number) => void;
    selectedKey: string[] | number[];
    disabled: boolean;
    stateKey: string;
}

export default class AsyncGroupsDropdowns extends React.Component<IAsyncGroupsDropdownsProps, IAsyncGroupsDropdownsState> {
    private selectedKeys: string[] | number[];

    constructor(props: IAsyncGroupsDropdownsProps, state: IAsyncGroupsDropdownsState) {
        super(props);
        this.selectedKeys = props.selectedKey;

        this.state = {
            loading: false,
            options: undefined,
            error: undefined
        };
    }

    public componentDidMount(): void {
        this.loadOptions();
    }

    public componentDidUpdate(prevProps: IAsyncGroupsDropdownsProps, prevState: IAsyncGroupsDropdownsState): void {
        if (this.props.disabled !== prevProps.disabled ||
            this.props.stateKey !== prevProps.stateKey) {
            this.loadOptions();
        }
    }

    private loadOptions(): void {
        this.setState({
            loading: true,
            error: undefined,
            options: undefined
        });

        this.props.loadOptions()
            .then((options: IComboBoxOption[]): void => {
                this.setState({
                    loading: false,
                    error: undefined,
                    options: sortBy(options, o => o.text)
                });
            }, (error: any): void => {
                this.setState((prevState: IAsyncGroupsDropdownsState, props: IAsyncGroupsDropdownsProps): IAsyncGroupsDropdownsState => {
                    prevState.loading = false;
                    prevState.error = error;
                    return prevState;
                });
            });
    }

    public render(): JSX.Element {
        const loading: JSX.Element = this.state.loading ? <div><Spinner label={'Loading options...'} /></div> : <div />;
        const error: JSX.Element = this.state.error !== undefined ? <div className={'ms-TextField-errorMessage ms-u-slideDownIn20'}>Error while loading items: {this.state.error}</div> : <div />;

        return (
            <div>
                <ComboBox label={this.props.label}
                    disabled={this.props.disabled || this.state.loading || this.state.error !== undefined}
                    multiSelect = {true}
                    onChanged={this.onChanged.bind(this)}
                    selectedKey={this.selectedKeys}
                    options={this.state.options} />
                {loading}
                {error}
            </div>
        );
    }

    private onChanged(option: IComboBoxOption, index?: number): void {
        const options: IComboBoxOption[] = this.state.options;
        options.forEach((o: IComboBoxOption): void => {
            if (o.key === option.key) {
                o.selected = option.selected;
            }
        });
        this.setState((prevState: IAsyncGroupsDropdownsState, props: IAsyncGroupsDropdownsProps): IAsyncGroupsDropdownsState => {
            prevState.options = options;
            return prevState;
        });
        if (this.props.onChanged) {
            this.props.onChanged(options, index);
        }
    }
}