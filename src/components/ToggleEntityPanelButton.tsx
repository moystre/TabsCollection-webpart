import { Icon } from 'office-ui-fabric-react';
import * as React from 'react';
import { CollapseExpandEntityPanel } from 'WorkPointStrings';
import styles from './ToggleEntityPanelButton.module.scss';

export interface IToggleEntityPanelButtonProps {
  onClick(event:any):void;
  shown:boolean;
}

export default class ToggleEntityPanelButton extends React.Component<IToggleEntityPanelButtonProps, null> {
  public render ():JSX.Element {

    const iconDirection:string = this.props.shown ? "ChevronUp" : "ChevronDown";

    return <div className={styles.toggleButton} onClick={this.props.onClick} title={CollapseExpandEntityPanel}><Icon iconName={iconDirection} /></div>;
  }
}