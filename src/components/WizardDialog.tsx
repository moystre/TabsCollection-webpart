import { DialogType } from 'office-ui-fabric-react';
import { Dialog } from 'office-ui-fabric-react/lib/Dialog';
import * as React from 'react';
import * as strings from 'WorkPointStrings';
import { IWizardMessageData, IWizardStartMessage, wizardResultEventTypes } from '../workPointLibrary/Event';
import { getClientProgramOpenUrl } from '../workPointLibrary/Helper';
import styles from './WizardDialog.module.scss';
import { IWorkPointBaseProps } from './WorkPointNavBarInterfaces';

interface ISizeObject {
  width: number;
  height: number;
}

interface IWizardDialogContentProps {
  url: string;
  closeWizard: () => void;
  iframeOnLoad?: (iframe: any) => void;
  width: string|number;
  height: string|number;
}

class WizardDialogContent extends React.Component<IWizardDialogContentProps> {
  private _iframe: any;

  constructor(props: IWizardDialogContentProps) {
    super(props);
  }

  public render(): JSX.Element {
    return (
      <div className={styles.wizardDialog} style={{ width: this.props.width }}>
        <iframe
          ref={(iframe) => { this._iframe = iframe; }}
          frameBorder={0}
          src={this.props.url}
          style={{
            width: this.props.width,
            height: this.props.height
          }} />

        <i title={strings.CloseThisDialog} onClick={this.props.closeWizard} className={`ms-Icon ms-Icon--ChromeClose ${styles.closeButton}`} aria-hidden="true"></i>
      </div>
    );
  }
}

export interface IWizardDialogState {
  hidden: boolean;
  width: number;
  height: number;
  url: string;
}

export default class WizardDialog extends React.Component<IWorkPointBaseProps, IWizardDialogState> {

  constructor(props:IWorkPointBaseProps) {
    super(props);
    const sizeObject = this.calculateWizardSize();
    this.state = {...sizeObject, url: null, hidden: true };
  }

  private calculateWizardSize = ():ISizeObject => {

    let width: number;
    let height: number;

    const windowWidth:number = window.innerWidth;
    const windowHeight:number = window.innerHeight;
    
    // Wizard default desired aspect ratio
    let desiredWizardRatio:number = 0.84;

    // How much of the horizontal space the Wizard dialog should use for calculating its size
    let maxAvailableSizeRatio:number = 1;

    // Set ratios for tablets
    if (windowWidth >= 768) {
      desiredWizardRatio = 0.72;
      maxAvailableSizeRatio = 0.9;
    }

    // Set ratio for desktops
    if (windowWidth >= 1024) {
      desiredWizardRatio = 0.64;
      maxAvailableSizeRatio = 0.75;
    }

    const containerWidth = Math.floor(windowWidth * maxAvailableSizeRatio);
    const containerHeight = Math.floor(windowHeight * maxAvailableSizeRatio);

    const projectedHeight = containerWidth * desiredWizardRatio;
    const projectedWidth = containerHeight / desiredWizardRatio;

    if (projectedHeight > containerHeight) {
      width = projectedWidth;
      height = projectedWidth * desiredWizardRatio;
    } else if (projectedWidth > containerWidth) {
      height = projectedHeight;
      width = projectedHeight / desiredWizardRatio;
    }

    return { width, height };
  }

  private handleWizardResultMessage = (event: MessageEvent):void => {

    if (event.origin === this.props.context.appLaunchUrl) {
      let messageData:IWizardMessageData = null;
      
      try {
        messageData = JSON.parse(event.data);
      } catch (exception) {
        messageData = event.data;
      }
      
      // Only accept valid Wizard event types
      if (typeof messageData.type === "string" && messageData.type !== "" && wizardResultEventTypes.indexOf(messageData.type) !== -1) {

        switch (messageData.action) {
          /**
           * Used as end action from many places
           */
          case "closeAndRefresh": {
            this.close();    
            window.location.reload();
            break;
          }
          /**
           * Return message from Document provisioning Wizard
           */
          case "openOnline": {
            this.close();
            window.location.href = messageData.redirectUrl;
            break;
          }
          /**
           * Return message from Document provisioning Wizard
           */
          case "openClient": {
            this.close();

            const openUrl:string = getClientProgramOpenUrl(messageData.redirectUrl);
            window.location.href = openUrl;
            break;
          }
          /**
           * Template library management dialog
           */
          case "closeDialogAndOpenFile": {
            this.close();
            const templateAbsoluteUrl:string = messageData.templateServerRelativeUrl.replace(this.props.context.sharePointContext.pageContext.site.serverRelativeUrl, this.props.context.solutionAbsoluteUrl);
            const openUrl:string = getClientProgramOpenUrl(templateAbsoluteUrl);
            window.location.href = openUrl;
          }
        }
      }
    }
  }

  private handleWizardStartMessage = (event: MessageEvent):void => {

    // Listen for start messages from both SharePoint and WorkPoint Addin-web
    if (event.origin === `${window.location.protocol}//${window.location.host}` || event.origin === this.props.context.appLaunchUrl) {
      const messageData:IWizardStartMessage = event.data;

      if (messageData.type === "workpointwizard") {

        let dialogArguments: string[] = [];
        const wizardUrl: string = messageData.url;
        let argumentString: string = null;

        dialogArguments.push(`SPHostUrl=${this.props.context.solutionAbsoluteUrl}`);
        dialogArguments.push(`SPLanguage=${this.props.context.sharePointContext.pageContext.cultureInfo.currentUICultureName}`);
        dialogArguments.push(`SPAppWebUrl=${this.props.context.appWebFullUrl}`);

        argumentString = dialogArguments.join("&");

        // TODO: Fear of url encoded "?" and "&". Check if its valid!
        if (wizardUrl.indexOf("?") !== -1) {
          argumentString = `&${argumentString}`;
        } else {
          argumentString = `?${argumentString}`;
        }

        const completeDialogUrl:string = `${this.props.context.appLaunchUrl}/${wizardUrl}${argumentString}`;

        const wizardArguments = {
          url: completeDialogUrl,
          hidden: false
        };

        const size = this.calculateWizardSize();

        this.setState({...wizardArguments, ...size});

        window.addEventListener("message", this.handleWizardResultMessage);
        window.addEventListener("resize", this.resize);
        window.removeEventListener("message", this.handleWizardStartMessage);
      }
    }
  }

  private resize = ():void => {
    this.setState(this.calculateWizardSize());
  }

  public componentDidMount():void {
    window.addEventListener("message", this.handleWizardStartMessage);
  }

  public componentWillUnmount():void {
    window.removeEventListener("resize", this.resize);
    window.removeEventListener("message", this.handleWizardResultMessage);
    window.removeEventListener("message", this.handleWizardStartMessage);
  }

  protected close = ():void => {

    this.setState({
      url: null,
      hidden: true
    });

    window.removeEventListener("resize", this.resize);
    window.removeEventListener("message", this.handleWizardResultMessage);
    window.addEventListener("message", this.handleWizardStartMessage);
  }

  public render(): JSX.Element {

    if (this.state.hidden) {
      return null;
    }

    const { width, height } = this.state;

    /* Resizing could be useful at some point
    const wizardContainerStyleString = `.${styles.wizardContainer} {
      height: ${height};
      width: ${width};
    }`;*/

    return (
      <Dialog
        hidden={this.state.hidden}
        dialogContentProps={{
          type: DialogType.close,
          className: styles.dialogContentOverrides
        }}
        modalProps={{
          isBlocking: true,
          containerClassName: styles.wizardContainer
        }}>
        <WizardDialogContent
          url={this.state.url}
          width={`${width}px`}
          height={`${height}px`}
          closeWizard={this.close} />
      </Dialog>
    );
  }
}