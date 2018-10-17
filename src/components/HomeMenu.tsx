import { Layer } from 'office-ui-fabric-react/lib/Layer';
import * as React from 'react';
import { BusinessModuleHierarchy } from './BusinessModuleHierarchy';
import homeMenuStyles from './HomeMenu.module.scss';
import { RollIn } from './RollIn';
import { SitePageMenu } from './SitePageMenu';
import SolutionInfo from './SolutionInfo';
import { INavbarConfigProps } from './WorkPointNavBar';

export interface IHomeMenuData extends INavbarConfigProps {
  attachTarget: HTMLElement;
  closeHomeMenu(): void;
}

export interface IHomeMenuState {
  top: number;
  show: boolean;
}

export default class HomeMenu extends React.Component<IHomeMenuData, IHomeMenuState> {

  public state:IHomeMenuState = {
    top: 0,
    show: false
  };

  public wrapperRef: HTMLElement;

  protected setWrapperRef = (node: HTMLElement): void => {
    this.wrapperRef = node;
  }

  protected handleClickOutside = (event: MouseEvent) => {
    setTimeout(() => {
      if (this.wrapperRef && !this.wrapperRef.contains(event.target as HTMLElement)) {
        this.props.closeHomeMenu();
      }
    }, 50);
  }

  public componentWillUnmount(): void {
    window.removeEventListener('click', this.handleClickOutside);
  }

  public componentDidMount():void {

    const { attachTarget } = this.props;

    if (!attachTarget) {
      return;
    }

    let heightOffsets: ClientRect = attachTarget.getBoundingClientRect();

    let heightOffset: number;

    // Just below navigation bar
    //heightOffset = heightOffsets.top + attachTarget.clientHeight;

    // Just below Suite bar
    heightOffset = heightOffsets.top;

    this.setState({
      show: true,
      top: heightOffset
    });

    // Set timeout so we do not risk landing in this function when its initialized.
    setTimeout(() => {
      window.addEventListener('click', this.handleClickOutside);
    });
  }

  public render(): JSX.Element {

    // Destructuring properties that are not needed by all children.
    const { attachTarget, closeHomeMenu, ...requiredChildProps } = this.props;

    return (
      <Layer>
        <RollIn
          inProp={this.state.show}
          mountOnEnter={true}
          unmountOnExit={true}
          duration={200}
          origin="left"
          {...this.state}>
          <div ref={this.setWrapperRef} className={homeMenuStyles.overlayWrapper}>
            <div className={`${homeMenuStyles.homeMenu} ${homeMenuStyles.overlayContent}`}>
              <SitePageMenu {...requiredChildProps} closeHomeMenu={closeHomeMenu} />
              <BusinessModuleHierarchy businessModuleSettingsCollection={this.props.workPointSettingsCollection.businessModuleSettings} context={this.props.context} />
              <SolutionInfo  {...requiredChildProps} />
            </div>
          </div>
        </RollIn>
      </Layer>
    );
  }
}