import { IPropertyPaneCustomFieldProps, IPropertyPaneField, PropertyPaneFieldType } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { ColorPickerHost, IColorPickerHostProp } from './PropertyPaneColorPickerHost';

export interface IColorPickerProp {
  label: string; //Label
  onPropChange: (targetProperty: string, oldValue: any, newValue: any) => void; // On Property Change function
  properties: any; //Web Part properties
  color?: string;  //value
  key?: string;  //unique key
}

export interface IColorPickerPropInternal extends IPropertyPaneCustomFieldProps {
  targetProperty: string;
  label: string;
  onPropChange: (targetProperty: string, oldValue: any, newValue: any) => void;
  onRender: (elem: HTMLElement) => void;
  onDispose: (elem: HTMLElement) => void;
  properties: any;
  color: string;
}

export class ColorPickerBuilder implements IPropertyPaneField<IColorPickerPropInternal>{
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IColorPickerPropInternal;

  private label: string;
  private onPropChange: (targetProperty: string, oldValue: any, newValue: any) => void;
  private color: string;
  private customProperties: any;
  private key: string;

  constructor(targetProperty: string, prop: IColorPickerPropInternal) {
    this.targetProperty = prop.targetProperty;
    this.properties = prop;
    this.label = prop.label;
    this.customProperties = prop.properties;
    this.onPropChange = prop.onPropChange.bind(this);
    this.properties.onRender = this.render.bind(this);
    this.properties.onDispose = this.dispose.bind(this);
    this.color = prop.color;
  }

  public render(elem: HTMLElement): void {
    let element: React.ReactElement<IColorPickerHostProp> = React.createElement(ColorPickerHost, {
      targetProperty: this.targetProperty,
      label: this.label,
      properties: this.customProperties,
      onDispose: null,
      onRender: null,
      onPropChange: this.onPropChange.bind(this),
      color: this.color,
      key: this.key
    });
    ReactDom.render(element, elem);
  }

  private dispose(elem: HTMLElement): void {
  }
}

export function PropertyPaneColorPicker(targetProperty: string, properties: IColorPickerProp): IPropertyPaneField<IColorPickerPropInternal> {

  const colorPickerProp: IColorPickerPropInternal = {
    targetProperty: targetProperty,
    label: properties.label,
    properties: properties.properties,
    onDispose: null,
    onRender: null,
    onPropChange: properties.onPropChange.bind(this),
    color: properties.color,
    key: properties.key
  };
  return new ColorPickerBuilder(targetProperty, colorPickerProp);
}