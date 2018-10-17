import * as React from 'react';
import * as strings from 'WorkPointStrings';
import menuStyles from './BreadcrumbMenu.module.scss';
import EntityBreadcrumbs from './EntityBreadcrumbs';
import HomeBreadcrumb from './HomeBreadcrumb';
import { INavbarConfigProps } from './WorkPointNavBar';

export interface IBreadcrumbMenuProps extends INavbarConfigProps {
  toggleHomeMenu(event:any): void;
  toggleEntityPanel(event:any): void;
  entityPanelShown:boolean;
  nonLicensedUser:boolean;
}

export interface IBreadcrumbMenuState {
  opened: boolean;
}

export default class BreadcrumbMenu extends React.Component<IBreadcrumbMenuProps, IBreadcrumbMenuState> {
  constructor(props: IBreadcrumbMenuProps) {
    super(props);

    this.state = {
      opened: false
    };
  }

  public render(): JSX.Element {

    const styleOverides: object = {
      position: "relative"
    };

    if (this.props.nonLicensedUser) {
      return (
      <div className={menuStyles.breadcrumbMenu}>
        <HomeBreadcrumb style={styleOverides} onMouseUp={() => {}} title={strings.Home} iconClass={"HomeSolid"} context={this.props.context} />
      </div>
      );
    }

    return (
      <div className={menuStyles.breadcrumbMenu}>
        <HomeBreadcrumb style={styleOverides} onMouseUp={this.props.toggleHomeMenu} title={strings.Home} iconClass={"HomeSolid"} context={this.props.context} />
        {this.props.currentEntity && <EntityBreadcrumbs {...this.props} />}
      </div>
    );
  }
}