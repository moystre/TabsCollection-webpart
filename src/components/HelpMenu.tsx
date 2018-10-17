import { ActionButton } from 'office-ui-fabric-react/lib/components/Button/ActionButton/ActionButton';
import * as React from 'react';
import * as strings from 'WorkPointStrings';
import Storage from '../workPointLibrary/Storage';
import ContextMenu from './ContextMenu';
import styles from './HelpMenu.module.scss';

export interface IHelpMenuProps {
  solutionAbsoluteUrl: string;
}

export interface IHelpMenuState {
  opened: boolean;
}

export default class HelpMenu extends React.Component<IHelpMenuProps, IHelpMenuState> {

  public wrapperRef: HTMLElement;

  constructor () {
    super();
    this.state = {
      opened: false
    };
  }

  protected setWrapperRef = (node: HTMLElement): void => {
    this.wrapperRef = node;
  }

  private clearStorage = () => {
    Storage.recycleAll(this.props.solutionAbsoluteUrl);
    window.location.href = window.location.href;
  }

  protected openHelpMenu = ():void => {
    this.setState({
      opened: true
    });
  }

  protected closeHelpMenu = ():void => {
    this.setState({
      opened: false
    });
  }

  public render ():JSX.Element {
    
    const { opened } = this.state;

    const menuItems: JSX.Element[] = [
      
    ];

    return (
      <div className={styles.container} ref={this.setWrapperRef}>
        <p className={styles.needHelp} onClick={this.openHelpMenu}>{strings.NeedHelp}?</p>
        
        {opened && <ContextMenu target={this.wrapperRef} close={this.closeHelpMenu}>
          <ActionButton className={styles.helpMenuItem} iconProps={{iconName: "Help"}} text={strings.WorkPointSupportCenter} href="https://support.workpoint.dk/hc/en-us/categories/200285698-WorkPoint-365"></ActionButton>
          <ActionButton className={styles.helpMenuItem} onClick={this.clearStorage} iconProps={{iconName: "Clear"}} text={strings.ClearWorkPointBrowserCache}></ActionButton>
        </ContextMenu>}
      </div>
    );
  }
}