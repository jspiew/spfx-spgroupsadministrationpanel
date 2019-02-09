import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType
} from '@microsoft/sp-webpart-base';
import { IComboBoxOption } from 'office-ui-fabric-react/lib/components/ComboBox';
import AsyncGroupsPicker, { IAsyncGroupsPickerProps, IAsyncGroupsPickerState } from './AsyncGroupsPicker';
import { IPropertyPaneCustomFieldProps } from '@microsoft/sp-webpart-base';
import { ITag } from 'office-ui-fabric-react/lib/Pickers';

export interface IPropertyPaneAsyncGroupsProps {
    label: string;
    loadOptions: () => Promise<ITag[]>;
    onPropertyChange: (propertyPath: string, newValue: any) => void;
    selectedKey: number[];
    disabled?: boolean;
}

export interface IPropertyPaneAsyncGroupsInternalProps extends IPropertyPaneAsyncGroupsProps, IPropertyPaneCustomFieldProps {
}

export class PropertyPaneAsyncGroups implements IPropertyPaneField<IPropertyPaneAsyncGroupsProps> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyPaneAsyncGroupsInternalProps;
  private elem: HTMLElement;

  constructor(targetProperty: string, properties: IPropertyPaneAsyncGroupsProps) {
    this.targetProperty = targetProperty;
    this.properties = {
      key: properties.label,
      label: properties.label,
      loadOptions: properties.loadOptions,
      onPropertyChange: properties.onPropertyChange,
      selectedKey: properties.selectedKey,
      disabled: properties.disabled,
      onRender: this.onRender.bind(this)
    };
  }

  public render(): void {
    if (!this.elem) {
      return;
    }

    this.onRender(this.elem);
  }

  private onRender(elem: HTMLElement): void {
    if (!this.elem) {
      this.elem = elem;
    }

    const element: React.ReactElement<IAsyncGroupsPickerProps> = React.createElement(AsyncGroupsPicker, {
      label: this.properties.label,
      loadOptions: this.properties.loadOptions,
      onChanged: this.onChanged.bind(this),
      selectedKey: this.properties.selectedKey,
      disabled: this.properties.disabled,
      // required to allow the component to be re-rendered by calling this.render() externally
      stateKey: new Date().toString()
    });
    ReactDom.render(element, elem);
  }

    private onChanged(options: ITag[], index?: number): void {
        this.properties.onPropertyChange(this.targetProperty, options.map(o => parseInt(o.key)));
    }
}