import { detect as detectCurrentBrowser } from 'detect-browser';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import * as React from 'react';
import { storage } from 'sp-pnp-js/lib/pnp';
import * as strings from 'WorkPointStrings';
import { IBusinessModuleEntity } from '../workPointLibrary/BusinessModule';
import { UserLicenseStatus } from '../workPointLibrary/License';
import { ListObject } from '../workPointLibrary/List';
import { WorkPointSettingsCollection } from '../workPointLibrary/Settings';
import { WorkPointStorageKey } from '../workPointLibrary/Storage';
import BreadcrumbMenu from './BreadcrumbMenu';
import EntityPanel from './EntityPanel';
import ExpressPanel, { ExpressOpenButton } from './ExpressPanel';
import HomeMenu from './HomeMenu';
import ToolMenu from './ToolMenu';
import WizardDialog from './WizardDialog';
import styles from './WorkPointNavBar.module.scss';
import { IWorkPointBaseProps } from './WorkPointNavBarInterfaces';

const NoLicenseMessage:React.SFC = ():JSX.Element => {
  return (
    <div className={`ms-font-m ${styles.noLicense}`}>
      <span className={styles.message}>
        <i className="ms-Icon ms-Icon--Warning" aria-hidden="true"></i> <span>{strings.YouHaveNoWorkPoint365License}</span>
      </span>
    </div>
  );
};

const LoadingIndicator:React.SFC = ():JSX.Element => {
  return (
    <div className={styles.loadingIndicator}>
      <Spinner className={styles.spinner} size={SpinnerSize.small} />
    </div>
  );
};

export interface INavbarConfigProps extends IWorkPointBaseProps {
  workPointSettingsCollection: WorkPointSettingsCollection;
  rootLists: ListObject[];
  currentEntity: IBusinessModuleEntity;
  ready: boolean;
}

export interface IWorkPointNavbarState {
  showHomeMenu: boolean;
  showEntityPanel: boolean;
  width: number;
}

export default class WorkPointNavBar extends React.Component<INavbarConfigProps, IWorkPointNavbarState> {

  /**
   * This must match breakpoints in 'Base.module.scss'
   * @see Base.module.scss
   */
  private entityPanelCollapseThreshold:number = 768;

  private deprecatedBrowser: boolean;
  public navbarRef: HTMLElement;

  private setNavbarRef = (node: HTMLElement): void => {
    this.navbarRef = node;
  }

  constructor(props: any) {
    super(props);

    try {
      const browser = detectCurrentBrowser();
      if (browser.name === "ie" && browser.version < "12") {
        this.deprecatedBrowser = true;
      }
    } catch (exception) {
      console.warn("Could not determine browser name and version. Unexpected behavior might occur.");
    }

    /**
     * Solution independent preference, so no prefix for this.
     */
    const savedEntityPanelVisibility:boolean = storage.session.get(WorkPointStorageKey.entityPanelVisibility);

    // Determine entity panel visibility based on view port.
    let showEntityPanel:boolean = (window.innerWidth < this.entityPanelCollapseThreshold) ? false : true;

    // If user has manually interacted with entity panel folding, we accept his choice.
    if (savedEntityPanelVisibility !== null && savedEntityPanelVisibility !== undefined) {
      showEntityPanel = savedEntityPanelVisibility;
    }

    this.state = {
      showHomeMenu: false,
      showEntityPanel,
      // IE 11 does not cope well with our Navigation bar having "width: 100%" when using flexbox design, therefore we manually set the width on init and resize
      width: window.innerWidth
    };
  }

  /**
   * IE 11 does not cope well with our Navigation bar having "width: 100%" when using flexbox design, therefore we manually set the width on init and resize
   */
  private updateDimensions = ():void => {
    this.setState({
      width: window.innerWidth
    });
  }

  /**
   * Called on resize events and collapses entity panel if below a certain threshold.
   */
  private toggleEntityPanelOnResize = ():void => {
    const showEntityPanel = (window.innerWidth < this.entityPanelCollapseThreshold) ? false : true;

    this.setState({
      showEntityPanel
    });
  }

  public componentDidMount():void {
    if (this.deprecatedBrowser) {
      window.addEventListener("resize", this.updateDimensions);
    }

    window.addEventListener("resize", this.toggleEntityPanelOnResize);
  }

  public componentWillUnmount():void {
    if (this.deprecatedBrowser) {
      window.removeEventListener("resize", this.updateDimensions);
    }

    window.removeEventListener("resize", this.toggleEntityPanelOnResize);
  }

  /**
   * Closes home menu
   */
  protected closeHomeMenu = () => {
    this.setState({ showHomeMenu: false });
  }

  /**
   * Opens home menu
   */
  protected openHomeMenu = () => {
    this.setState({ showHomeMenu: true });
  }

  /**
   * Toggles home menu, reverses prior state.
   * Initiated manually through child component.
   * 
   * @see <HomeBreadcrumb>
   */
  protected toggleHomeMenu = (event: any) => {

    const mouseEvent = event as MouseEvent;

    // Has CTRL been active during this click, or did user click the middle mouse button?
    const newTab:boolean = mouseEvent.ctrlKey === true || (mouseEvent.button && mouseEvent.button === 1);

    if (newTab) {
      window.open(this.props.context.solutionAbsoluteUrl, "_blank");
      return null;
    }

    if (!this.state.showHomeMenu) {
      this.openHomeMenu();
    } else {
      this.closeHomeMenu();
    }
  }

  /**
   * Controls entity panels state. Seeps through to child components using it.
   * Initiated manually through child component.
   * 
   * @see <ToggleEntityPanelButton>
   */
  protected toggleEntityPanel = (event: any) => {
    const usersChoice:boolean = !this.state.showEntityPanel;

    storage.session.put(WorkPointStorageKey.entityPanelVisibility, usersChoice);

    this.setState({
      showEntityPanel: usersChoice
    });
  }

  public render(): JSX.Element {

    const style = this.deprecatedBrowser ? { width: this.state.width } : null;
    const { ready, currentEntity, context } = this.props;
    const { showEntityPanel, showHomeMenu } = this.state;
    const nonLicensedUser = this.props.context.userLicense.Status === UserLicenseStatus.None;

    return (
      <section ref={this.setNavbarRef} style={style} className={styles.workpointUI}>
        <section className={styles.navbar}>
          <BreadcrumbMenu {...this.props} toggleHomeMenu={this.toggleHomeMenu} nonLicensedUser={nonLicensedUser} entityPanelShown={showEntityPanel} toggleEntityPanel={this.toggleEntityPanel} />
          {!nonLicensedUser && <ToolMenu {...this.props} />}
          {!nonLicensedUser && <ExpressOpenButton />}
        </section>
        {!nonLicensedUser && currentEntity && showEntityPanel && <EntityPanel {...this.props} />}
        {!nonLicensedUser && showHomeMenu && <HomeMenu closeHomeMenu={this.closeHomeMenu} attachTarget={this.navbarRef} {...this.props} />}

        {!nonLicensedUser && ready && <ExpressPanel currentEntity={currentEntity} attachTarget={this.navbarRef} context={context} />}

        {!nonLicensedUser && ready && <WizardDialog context={context} />}

        {nonLicensedUser && <NoLicenseMessage />}

        { // Show loader when not ready.
          !ready && <LoadingIndicator />}
      </section>
    );
  }
}