import Resizable, { NumberSize, ResizableDirection } from 're-resizable';
import * as React from 'react';
import { storage, util } from 'sp-pnp-js/lib/pnp';
import { getListByIdOrUrl } from '../workPointLibrary/Helper';
import { WorkPointContentType } from '../workPointLibrary/List';
import { IStageHistory } from '../workPointLibrary/Stage';
import { WorkPointStorageKey } from '../workPointLibrary/Storage';
import ActivityControl from './Activity';
import { EntityDetails } from './EntityDetails';
import styles from './EntityPanel.module.scss';
import StageControl from './Stage';
import { INavbarConfigProps } from './WorkPointNavBar';

export interface IEntityPanelState {
  height: number | string;
  resizingEnabled: boolean;
}

export default class EntityPanel extends React.Component<INavbarConfigProps, IEntityPanelState> {

  /**
   * This must match breakpoints in 'Base.module.scss'
   * @see Base.module.scss
   */
  private entityPanelCollapseThreshold:number = 768;

  private stagingEnabled: boolean;

  private businessModuleContentTypes: WorkPointContentType[];

  constructor (props:INavbarConfigProps) {
    super(props);

    /**
     * This must match breakpoints in 'Base.module.scss'
     * @see Base.module.scss
     */
    const resizingEnabled = (window.innerWidth < this.entityPanelCollapseThreshold) ? false : true;

    const height:number = storage.local.get(WorkPointStorageKey.entityPanelHeight);
    this.stagingEnabled = props.currentEntity.Settings.StageSettings.Enabled;

    if (this.stagingEnabled) {
      this.businessModuleContentTypes = getListByIdOrUrl(props.rootLists, props.currentEntity.ListId).ContentTypes;
    }

    this.state = {
      height: height || 100,
      resizingEnabled
    };
  }

  /**
   * TODO: Debounce this
   */
  private onResizeStop = (event:MouseEvent|TouchEvent, direction: ResizableDirection, ref:HTMLElement, delta:NumberSize):void => {
    
    this.setState((prevState) => {
      const newHeight:number = parseInt(prevState.height.toString()) + delta.height;
      storage.local.put(WorkPointStorageKey.entityPanelHeight, newHeight, util.dateAdd(new Date(), "year", 1));
      return { height: newHeight };
    });
  }

  private setIsResizingEnabled = ():void => {
    const resizingEnabled = (window.innerWidth < this.entityPanelCollapseThreshold) ? false : true;

    this.setState({
      resizingEnabled
    });
  }

  public componentDidMount():void {
    window.addEventListener("resize", this.setIsResizingEnabled);
  }

  public componentWillUnmount():void {
    window.removeEventListener("resize", this.setIsResizingEnabled);
  }

  public render ():JSX.Element {

    const { resizingEnabled, height } = this.state;

    const enabledResizingHandles:any = {
      bottom: resizingEnabled,
      top:false, right:false, left:false, topRight:false, bottomRight:false, bottomLeft:false, topLeft:false
    };
    
    const resizeSize:any = {
      height: resizingEnabled ? height : "100%",
      width: "100%"
    };

    const handleStyles:any = {
      bottom: {
        cursor: "ns-resize"
      }
    };

    let stageHistory:IStageHistory[] = null;
    
    if (this.stagingEnabled) {
      try {
        stageHistory = JSON.parse(this.props.currentEntity.wp_stageHistory);
      } catch (exception) {}
    }

    return (
      <Resizable className={styles.panel} enable={enabledResizingHandles} size={resizeSize} handleStyles={handleStyles} onResizeStop={this.onResizeStop} minHeight={50}>
        <div className={styles.columnLeftBlock}>
          <EntityDetails fieldValues={this.props.currentEntity.FieldValues} context={this.props.context} />
        </div>
        <div className={styles.rightBlock}>
          <ActivityControl context={this.props.context} currentEntity={this.props.currentEntity} />
          {this.stagingEnabled && <StageControl 
            currentEntity={this.props.currentEntity}
            context={this.props.context}
            stageHistory={stageHistory}
            stageSettings={this.props.currentEntity.Settings.StageSettings}
            businessModuleContentTypes={this.businessModuleContentTypes}
          />}
        </div>
      </Resizable>
    );
  }
}