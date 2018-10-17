import { Layer } from 'office-ui-fabric-react/lib/Layer';
import Resizable, { NumberSize, ResizableDirection } from 're-resizable';
import * as React from 'react';
import { storage, util } from 'sp-pnp-js/lib/pnp';
import { IBusinessModuleEntity } from '../workPointLibrary/BusinessModule';
import { IIncomingExpressPanelMessage, IOutGoingExpressPanelMessage, IWorkPointMessageData } from '../workPointLibrary/Event';
import { parseSlashDividedDateValue } from '../workPointLibrary/Helper';
import WorkPointStorage, { IRawStorageObject, WorkPointStorageKey } from '../workPointLibrary/Storage';
import styles from './ExpressPanel.module.scss';
import { RollIn } from './RollIn';
import { IWorkPointBaseProps } from './WorkPointNavBarInterfaces';

interface IExpressButtonProps {
  onClick: () => void;
  type: "close"|"open";
}

const ExpressButton:React.SFC<IExpressButtonProps> = ({ onClick, type }):JSX.Element => {
  return (
    <div className={styles[`${type}Button`]} onClick={onClick}>
      <img className={styles.expressLogo} src="https://workpoint.azureedge.net/Images/workpoint-express-icon.png" />
    </div>
  );
};

class ExpressCloseButton extends React.Component {

  private closeExpressPanel = ():void => {
    const message:IIncomingExpressPanelMessage = {
      action: "close",
      type: "expressPanel"
    };

    window.postMessage(message, window.location.href);
  }
  
  public render ():JSX.Element {
    return <ExpressButton onClick={this.closeExpressPanel} type="close" />;
  }
}

export class ExpressOpenButton extends React.Component {

  protected openExpressPanel = ():void => {
    const message:IIncomingExpressPanelMessage = {
      action: "open",
      type: "expressPanel"
    };

    window.postMessage(message, window.location.href);
  }
  
  public render ():JSX.Element {
    return <ExpressButton onClick={this.openExpressPanel} type="open" />;
  }
}

export interface IExpressPanelProps extends IWorkPointBaseProps {
  attachTarget: HTMLElement;
  currentEntity: IBusinessModuleEntity;
}

export interface IExpressPanelState {
  top: number;
  show: boolean;
  initialized: boolean;
  width: number;
}

export default class ExpressPanel extends React.Component<IExpressPanelProps, IExpressPanelState> {
  private _iframe: any;

  private frameUrl:string;
  private defaultMinimumWidth:number;

  constructor (props:IExpressPanelProps) {
    super(props);

    this.defaultMinimumWidth = 300;

    const width:number = this.getExpressPanelWidth();

    let expressPanelArguments:string[] = [];

    expressPanelArguments.push(`SPHostUrl=${props.context.solutionAbsoluteUrl}`);
    expressPanelArguments.push(`SPLanguage=${props.context.sharePointContext.pageContext.cultureInfo.currentUICultureName}`);
    expressPanelArguments.push(`SPAppWebUrl=${props.context.appWebFullUrl}`);

    if (props.currentEntity && typeof props.currentEntity.ItemLocation === "string" && props.currentEntity.ItemLocation !== "") {
      expressPanelArguments.push(`wpItemLocation=${props.currentEntity.ItemLocation}`);
    }

    expressPanelArguments.push(`webTitle=${props.context.sharePointContext.pageContext.web.title}`);

    this.frameUrl = `${props.context.appLaunchUrl}/ExpressPanel/?${expressPanelArguments.join("&")}`;

    this.state = {
      top: 0,
      show: false,
      initialized: false,
      width: width || this.defaultMinimumWidth
    };
  }

  private open = ():void => {
    
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
      top: heightOffset,
      initialized: true
    });
  }

  private close = ():void => {
    this.setState({
      show: false
    });
  }

  private toggle = ():void => {
    this.state.show ? this.close() : this.open();
  }

  private handleExpressPanelMessage = (event: MessageEvent):void => {

    if (event.origin === `${window.location.protocol}//${window.location.host}` || event.origin === this.props.context.appLaunchUrl) {

      const messageData:IWorkPointMessageData = event.data;

      // Only accept valid Express Panel event types
      if (typeof messageData.type === "string" && messageData.type === "expressPanel") {

        const expressPanelMessageData:IIncomingExpressPanelMessage = messageData as IIncomingExpressPanelMessage;

        switch (expressPanelMessageData.action) {
          case "open": {
            this.open();
            break;
          }
          case "close": {
            this.close();
            break;
          }
          /**
           * This message was bein used for other purposes in WorkPoint classic, now its being used for setting the first focus.
           * It is only called the first time the Express Panel is loaded (the apps page load), which is exactly the time we want to set focus in the Express Panel search field.
           */
          case "signalReady": {
            // We fake an F2 click here
            this.eventToggle(113, false);
            break;
          }
          case "toggle": {
            this.toggle();
            break;
          }
          case "openUrl": {
            window.location.href = expressPanelMessageData.redirectUrl;
            break;
          }
          case "checkIfWorkPointSettingsAreOutOfSync": {
            this.checkIfWorkPointSettingsAreOutOfSync(expressPanelMessageData.SettingsChangedUTCTimeStamp);
          }
        }
      }

      /* Not implemented as we are not sure how the Favorite flow will be in Modern UI.
      if (typeof messageData.type === "string" && messageData.type === "favorites") {

        const favoritesMessageData:IIncomingFavoriteMessage = messageData as IIncomingFavoriteMessage;

        switch (favoritesMessageData.action) {
          case "AddedFavorite":
          case "RemovedFavorite":
          case "ClearFavoriteCache":
            break;
        }
      }*/
    }
  }

  /**
   * Checks for expired WorkPointSettings cache and flushes if neccessary.
   * 
   * @argument dateCandidate Slash divided string UTC date.
   * @see parseSlashDividedDateValue
   */
  private checkIfWorkPointSettingsAreOutOfSync = (dateCandidate:string):void => {

    const settingsChanged:Date = parseSlashDividedDateValue(dateCandidate);

    // No valid date, so exit
    if (!settingsChanged) {
      return;
    }

    const workPointSettingStorageKey:string = WorkPointStorage.getKey(WorkPointStorageKey.workPointSettings, this.props.context.solutionAbsoluteUrl);
    const rootListsStorageKey:string = WorkPointStorage.getKey(WorkPointStorageKey.rootLists, this.props.context.solutionAbsoluteUrl);

    // Manually get expiration of WorkPointSettings cache key
    const workPointSettingsStorageItem: IRawStorageObject = JSON.parse(localStorage.getItem(workPointSettingStorageKey));

    if (!workPointSettingsStorageItem) {
      return;
    }

    const now: Date = new Date();
    const settingsExpiration = new Date(workPointSettingsStorageItem.expiration);
    const settingsCacheInitiation: Date = util.dateAdd(settingsExpiration, "day", -1);

    // Settings have changed since we cached this item
    if (settingsExpiration > now && settingsChanged > settingsCacheInitiation) {
      storage.local.delete(rootListsStorageKey);
      storage.local.delete(workPointSettingStorageKey);
    }
  }

  private getExpressPanelWidth = ():number => {

    // We have a valid state
    if (this.state && typeof this.state.width === "number") {

      if (this.state.width >= window.innerWidth) {
        return window.innerWidth * 0.9;
      } else if (this.state.width < window.innerWidth) {
        return this.state.width;
      }

    } else {

      // No state width has been set yet

      // First try to load saved user width
      const lastSavedExpressPanelWidth:number = storage.local.get(WorkPointStorageKey.expressPanelWidth);

      // If saved user width is sane, then apply it.
      if (lastSavedExpressPanelWidth && lastSavedExpressPanelWidth >= this.defaultMinimumWidth && lastSavedExpressPanelWidth < window.innerWidth) {
        return lastSavedExpressPanelWidth;
      }
    }

    return this.defaultMinimumWidth;
  }

  private handleResizeAndRotation = ():void => {
    const desiredWidth:number = this.getExpressPanelWidth();
    this.setState({width: desiredWidth});
  }

  private handleKeyUp = (event:KeyboardEvent):void => {

    // Listen for F2 clicks
    if (event.which === 113) {
      this.eventToggle(113);
    }
  }

  /**
   * Sending messages to the Express Panel whether it should be shown or not. It posts messages back to this component, handling its view state.
   * 
   * @param eventKey Key 'which' value of a Keyboard event
   * @param showOverride Optional. Override for telling the Express Panel what state it should this it is in.
   */
  private eventToggle = (eventKey: number, showOverride?:boolean):void => {

    // If Express Panel is not initialized, start it up. We will control focus on 'signalReady' message
    if (!this.state.initialized) {
      this.open();
    } else if (this._iframe && this._iframe.contentWindow) {
      // Toggle the Express Panel

      let expressPanelVisibleState:boolean = this.state.show;

      if (showOverride !== undefined && showOverride !== null) {
        expressPanelVisibleState = showOverride;
      }
  
      const message: IOutGoingExpressPanelMessage = {
        method: "ToggleExpress",
        expressVisibleInWorkPoint: expressPanelVisibleState,
        tab: null,
        eventKey: eventKey
      };
      this._iframe.contentWindow.postMessage(message, this.props.context.appLaunchUrl);
    }
  }

  public componentDidMount(): void {
    window.document.addEventListener("keyup", this.handleKeyUp);
    window.addEventListener("message", this.handleExpressPanelMessage);
    window.addEventListener("resize", this.handleResizeAndRotation);
  }

  public componentWillUnmount(): void {
    window.document.removeEventListener("keyup", this.handleKeyUp);
    window.removeEventListener("message", this.handleExpressPanelMessage);
    window.removeEventListener("resize", this.handleResizeAndRotation);
  }

  private onResizeStop = (event:MouseEvent|TouchEvent, direction: ResizableDirection, ref:HTMLElement, delta:NumberSize):void => {
    this.setState((prevState) => {
      const newWidth:number = prevState.width + delta.width;
      storage.local.put(WorkPointStorageKey.expressPanelWidth, newWidth, util.dateAdd(new Date(), "year", 1));
      return { width: newWidth };
    });
  }
  
  public render ():JSX.Element {

    const { show } = this.state;

    return (
      <Layer>
        <div className={styles.panel}>
          <RollIn
            inProp={show} duration={200}
            origin="right"
            unmountOnExit={false}
            mountOnEnter={false}
            top={this.state.top}
          >
            <Resizable
              enable={{
                left: show,
                top:false,
                right:false,
                bottom:false,
                topRight:false,
                bottomRight:false,
                bottomLeft:false,
                topLeft:false
              }}
              size={{
                width: this.state.width,
                height: "100%"
              }}
              minWidth={this.defaultMinimumWidth}
              handleStyles={{left: {cursor: "ew-resize"}}}
              onResizeStop={this.onResizeStop}
            >
              {this.state.initialized && <div className={show ? styles.activeFrameContainer : styles.frameContainer}>
                <iframe
                  ref={(iframe) => {this._iframe = iframe; }}
                  frameBorder={0}
                  src={this.frameUrl}
                  style={{
                    height: "100%",
                    width: "100%"
                  }}
                  className={styles.iframeElement}
                />
                <ExpressCloseButton />
              </div>
              }
            </Resizable>
          </RollIn>
        </div>
      </Layer>
    );
  }
}