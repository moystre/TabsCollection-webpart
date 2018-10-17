
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { Moment } from 'moment';
import * as React from 'react';
import { ShowAllRelatedEvents } from 'WorkPointStrings';
import { getRootSiteCollectionUrl, getWorkPointAppLaunchParameters, IWorkPointAppLaunchParameters } from '../workPointLibrary/service';
import { getAspNetTicksFromDate } from '../workPointLibrary/Time';
import styles from './LogLink.module.scss';

export interface ILogLinkProps {
  loggingScopeId: string;
  logEndTime: Moment;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
}

export default class LogLink extends React.Component<ILogLinkProps, null> {

  private navigateToLogItem = async (event:any):Promise<void> => {

    try {
      
      const mouseEvent = event as MouseEvent;
      
      // Has CTRL been active during this click, or did user click the middle mouse button?
      const newTab:boolean = mouseEvent.ctrlKey === true || (mouseEvent.button && mouseEvent.button === 1);

      const { loggingScopeId, logEndTime, pageContext, spHttpClient } = this.props;

      const solutionAbsoluteUrl:string = await getRootSiteCollectionUrl(pageContext.site.absoluteUrl);
      const appLaunchParameters:IWorkPointAppLaunchParameters = await getWorkPointAppLaunchParameters(solutionAbsoluteUrl, pageContext, spHttpClient);

      let appWebArguments: string[] = [];

      /**
       * Get .Net epoch ticks from Moment date
       */
      let dateTicks:number = null;

      if (logEndTime) {
        dateTicks = getAspNetTicksFromDate(logEndTime.toDate());
      }

      appWebArguments.push(`LoggingScopeId=${loggingScopeId}${dateTicks ? `&loggingEndTime=${dateTicks}`:``}`);

      appWebArguments.push(`SPHostUrl=${solutionAbsoluteUrl}`);
      appWebArguments.push(`SPLanguage=${pageContext.cultureInfo.currentUICultureName}`);
      appWebArguments.push(`SPAppWebUrl=${appLaunchParameters.appWebFullUrl}`);

      const argumentString: string = appWebArguments.join("&");
      const loggingPageURL:string = `${appLaunchParameters.appLaunchUrl}/Logging?${argumentString}`;

      if (newTab) {
        window.open(loggingPageURL, "_blank");
      } else {
        window.location.href = loggingPageURL;
      }

    } catch (exception) {}

    return null;
  }

  public render(): JSX.Element {
    return <a className={styles.link} onMouseUp={this.navigateToLogItem} title={ShowAllRelatedEvents}>{ShowAllRelatedEvents}</a>;
  }
}