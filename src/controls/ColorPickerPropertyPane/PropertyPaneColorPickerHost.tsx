import { ColorPicker, IconButton, Label, PrimaryButton, TextField } from 'office-ui-fabric-react';
import * as React from 'react';
import { IColorPickerPropInternal } from './PropertyPaneColorPicker';
export interface IColorPickerHostProp extends IColorPickerPropInternal {
}
export interface IColorPickerHostState {
  color?: any;
  showColorPicker?: boolean;
}

export class ColorPickerHost extends React.Component<IColorPickerHostProp, IColorPickerHostState>{
  constructor(prop: IColorPickerHostProp) {
    super(prop);
    this.state = ({
      color: prop.color,
      showColorPicker: false
    });
    this.onColorChanged = this.onColorChanged.bind(this);
    this.showColorPicker = this.showColorPicker.bind(this);
    this.closeColorPicker = this.closeColorPicker.bind(this);
    this.onTextColorChanged = this.onTextColorChanged.bind(this);
    this.updatePropValue = this.updatePropValue.bind(this);
  }

  public onColorChanged = (color: string): void => {
    this.setState({
      color: color
    });
  }

  public render(): JSX.Element {
    return (
      <div>
        <div style={{
          display: 'inline-block'
        }}>
          <Label>{this.props.label}</Label>
          <div style={{
            width: 115,
            float: 'left'
          }}>
            <TextField label='' value={this.state.color} onChanged={this.onTextColorChanged} />
          </div>
          <div style={{
            backgroundColor: this.state.color,
            width: 60,
            height: 25,
            float: 'left',
            marginLeft: 8
          }}></div>
          <IconButton
            iconProps={{ iconName: 'ColorSolid' }}
            onClick={this.showColorPicker}
            title='Open color picker'
            style={{
              float: 'left',
              marginLeft: 12,
              marginTop: -4
            }}
          />
        </div>

        {(this.state.showColorPicker) ? (
          <div style={{
            marginTop: 10
          }}>
            <PrimaryButton onClick={this.closeColorPicker}>OK</PrimaryButton>
            <ColorPicker color={this.state.color} onColorChanged={this.onColorChanged} />
          </div>
        ) : (null)}
      </div>
    );
  }

  private showColorPicker = (): void => {
    this.setState({
      showColorPicker: true
    });
  }

  private closeColorPicker = (): void => {
    this.setState({
      showColorPicker: false
    }, () => {
      this.updatePropValue();
    });
  }

  private onTextColorChanged = (newValue: string): void => {
    this.setState({
      color: newValue
    }, () => {
      this.updatePropValue();
    });
  }

  private updatePropValue = ():void => {
    let oldValues = this.props.properties[this.props.targetProperty];
    this.props.properties[this.props.targetProperty] = this.state.color;
    this.props.onPropChange(this.props.targetProperty, oldValues, this.state.color);
  }
}