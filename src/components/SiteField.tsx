import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import * as React from 'react';
import { spODataEntity, Web } from 'sp-pnp-js/lib/pnp';
import * as strings from 'WorkPointStrings';
import { BusinessModuleEntity } from '../workPointLibrary/BusinessModule';
import { BusinessModuleFields } from '../workPointLibrary/Fields';
import { normalizeURLValue } from '../workPointLibrary/Helper';
import { SiteStatus } from '../workPointLibrary/Site';
import styles from './SiteField.module.scss';

export interface ISiteFieldProps {
  urlValue: string;
  title: string;
  listId: string;
  solutionAbsoluteUrl: string;
  itemId: number;
}

export interface ISiteFieldState {
  site: string;
  errorMessage: string;
}

export default class SiteField extends React.Component<ISiteFieldProps, ISiteFieldState> {

  private web:Web;
  private listId:string;
  private itemId:number;

  /**
   * Use the mounting property to prevent unneccessary calls if component unmounts.
   */
  private mounted:boolean;
  private siteCreationCheckTimeout:number;

  constructor (props:ISiteFieldProps) {
    super(props);

    this.listId = props.listId;
    this.itemId = props.itemId;
    this.siteCreationCheckTimeout = 10000;

    this.state = {
      site: null,
      errorMessage: null
    };
  }

  /**
   * A call that checks for a specific items site creation. It will call itself every <n> milliseconds, corresponding to the 'siteCreationCheckTimeout' property.
   * 
   * This method could be improved so that it checks for multiple list items site creation at once.
   * 
   * @returns Promise<void> 
   */
  private getUpdatedSiteValue = async ():Promise<void> => {

    if (this.mounted) {

      let item:BusinessModuleEntity = null;
    
      try {
        item = await this.web.lists.getById(this.listId).items.getById(this.itemId).select(BusinessModuleFields.Site).getAs(spODataEntity(BusinessModuleEntity));
      } catch (error) {
        this.setState({
          site: SiteStatus.FrontEndQueryFailure,
          errorMessage: error
        });
        return null;
      }
  
      // Site value changed
      if (typeof item.wpSite === "string" && item.wpSite !== "" && item.wpSite !== SiteStatus.Provisioning) {
        this.setState({
          site: item.wpSite
        });
      } else {
        setTimeout(this.getUpdatedSiteValue, this.siteCreationCheckTimeout);
      }
    }
  }

  /**
   * Will the contexts page context be enough if the wpSite field customizer is being shown in a web part on a different site collection?
   * Will it need the WP 365 Solution URL instead of 'this.props.context.pageContext.site.absoluteUrl'?
   */
  public componentDidMount():void {
    this.mounted = true;
    const { urlValue } = this.props;

    if (urlValue === SiteStatus.Provisioning) {
      this.web = new Web(this.props.solutionAbsoluteUrl);
      setTimeout(this.getUpdatedSiteValue, this.siteCreationCheckTimeout);
    }
  }

  public componentWillUnmount():void {
    this.mounted = false;
  }

  /**
   * Support for touch devices, as SharePoint cancels href clicks in favor of row selection in built in lists.
   */
  private onTouchStart = (urlValue:string):void => {
    window.location.href = urlValue;
  }

  public render(): JSX.Element {

    const { urlValue } = this.props;
    const { site } = this.state;

    /**
     * If an updated site value is fetched, use that instead of inherited property value.
     * @see getUpdatedSiteValue
     */
    let wpSiteValue:string = site ? site : urlValue;

    if (typeof wpSiteValue !== "string" ) {
      return null;
    }

    if (wpSiteValue === "" || wpSiteValue === SiteStatus.IgnoreCreateSiteOnItemAddedEvent) {
      return null;
    }

    switch (wpSiteValue) {

      /**
       * Failed site creation
       */
      case SiteStatus.Failed: {
        return <i title={strings.SiteCouldNotBeCreated} className={`ms-Icon ms-Icon--Warning ${styles.backendWarning}`} aria-hidden="true"></i>;
      }

      /**
       * Site is being provisioned
       */
      case SiteStatus.Provisioning: {
        return <div title={strings.SiteCreationIsUnderway} className={styles.field}><Spinner size={ SpinnerSize.xSmall } /></div>;  
      }

      /**
       * Site creation status could not be checked (client/backend communication error)
       */
      case SiteStatus.FrontEndQueryFailure: {
        return <i title={`${strings.SiteCreationStatusCouldNotBeCheckedDueToTheFollowingError}: ${this.state.errorMessage}`} className={`ms-Icon ms-Icon--Warning ${styles.frontendWarning}`} aria-hidden="true"></i>;        
      }

      /**
       * Proper Site value being shown
       */
      default: {

        const normalizedURLValue:string = normalizeURLValue(wpSiteValue);
        return (
          <a title={`${strings.GoToTheSiteFor} '${this.props.title}'`} className={styles.field} href={normalizedURLValue} target="_top" onTouchStart={() => this.onTouchStart(normalizedURLValue)}>
            <i className="ms-Icon ms-Icon--Favicon" aria-hidden="true"></i>
          </a>
        );
      }
    }
  }
}
