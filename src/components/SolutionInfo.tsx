
import SPPermission from '@microsoft/sp-page-context/lib/SPPermission';
import { IconButton } from 'office-ui-fabric-react';
import * as React from 'react';
import * as strings from 'WorkPointStrings';
import { UserLicenseStatus } from '../workPointLibrary/License';
import HelpMenu from './HelpMenu';
import styles from './SolutionInfo.module.scss';
import { INavbarConfigProps } from './WorkPointNavBar';

export default class SolutionInfo extends React.Component<INavbarConfigProps, null> {
  
  public render():JSX.Element {

    const { context } = this.props;
    const { Version, Status } = context.userLicense;

    let licenseDisplayText: string = null;

    switch (Status) {
      case UserLicenseStatus.Full:
        licenseDisplayText = strings.LicensedUserDescription;
        break;
      case UserLicenseStatus.Limited:
        licenseDisplayText = strings.LimitedLicenseDescription;
        break;
      case UserLicenseStatus.None:
        licenseDisplayText = strings.NoLicenseDescription;
        break;
      case UserLicenseStatus.External:
        licenseDisplayText = strings.ExternalUserLicenseDescription;
        break;
    }

    let administrationArguments: string[] = [];

    administrationArguments.push(`SPHostUrl=${context.solutionAbsoluteUrl}`);
    administrationArguments.push(`SPLanguage=${context.sharePointContext.pageContext.cultureInfo.currentUICultureName}`);
    administrationArguments.push(`SPAppWebUrl=${context.appWebFullUrl}`);

    const argumentString: string = administrationArguments.join("&");
    const administrationLink:string = `${context.appLaunchUrl}?${argumentString}`;
    const isUserAdmin:boolean = context.sharePointContext.pageContext.web.permissions.hasPermission(SPPermission.manageWeb);

    return (
      <div className={styles.solutionInfoContainer}>
        <div className={styles.solution}>
          <div className={styles.brandContainer}>
            <div className={styles.brand}></div>
            {isUserAdmin && <IconButton className={styles.administrationLink} href={administrationLink} title={`${strings.GoTo} ${strings.WorkPoint365Administration}`} iconProps={{iconName: "Settings"}}></IconButton>}
          </div>
          {Version && <div className={styles.version}>{strings.Version}: {Version}</div>}
          {licenseDisplayText && <div className={styles.version}>{licenseDisplayText}</div>}
        </div>
        
        <HelpMenu solutionAbsoluteUrl={context.solutionAbsoluteUrl} />
      </div>
    );
  }
}