import { Callout } from 'office-ui-fabric-react/lib/Callout';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import { IContextualMenuItem } from 'office-ui-fabric-react/lib/components/ContextualMenu/ContextualMenu.Props';
import * as React from 'react';
import { storage, Util } from 'sp-pnp-js/lib/pnp';
import * as strings from 'WorkPointStrings';
import { IBusinessModuleEntity } from '../workPointLibrary/BusinessModule';
import { IAddStageFilterMessage, IBasicStageFilterMessage, IWizardStartMessage } from '../workPointLibrary/Event';
import { isValidDate } from '../workPointLibrary/Helper';
import { WorkPointContentType } from '../workPointLibrary/List';
import { getOrderedStageMap, getStageConstraints, IConstraint, IHistoryConstraintStatus, IStageHistory, IStageSettings, StageObject, StageStatus, StageType } from '../workPointLibrary/Stage';
import WorkPointStorage, { WorkPointStorageKey } from '../workPointLibrary/Storage';
import styles from './Stage.module.scss';
import { IWorkPointBaseProps } from './WorkPointNavBarInterfaces';

export interface IStageBaseProps extends IWorkPointBaseProps {
  deadline: Date;
  stage: StageObject;
  isFiltered: boolean;
  stageSettings: IStageSettings;
}

interface IStageDetailsProps extends IStageBaseProps {
  changeStage(stageId:string):void;
  toggleStageFilter():void;
  stageElement: JSX.Element;
  deadlineOverdue: boolean;
  deadlineComingSoon: boolean;
}

class StageDetails extends React.Component<IStageDetailsProps, null> {

  public render ():JSX.Element {

    const { isFiltered, stage, stageSettings, stageElement, deadline, deadlineComingSoon, deadlineOverdue, context } = this.props;

    let stageActions:IContextualMenuItem[] = [];

    const filterText:string = isFiltered ? strings.RemoveStageFilter : strings.ApplyStageFilter;
    const filterIcon:string = isFiltered ? "ClearFilter" : "Filter";

    // If stage has valid history (entity has been in this stage), let users add it as a stage filter
    if ((stage.history || stage.filterable) && (stage.status === StageStatus.History || stage.status === StageStatus.Active)) {
      stageActions.push(
        {
          key: 'stageFilter',
          name: filterText,
          icon: filterIcon,
          checked: isFiltered,
          onClick: this.props.toggleStageFilter
        }
      );
    }

    // Can user transition to this stage? If so, let him change to it
    if (stage.transitionable) {
      stageActions.push(
        {
          key: 'changeStage',
          name: strings.ChangeStage,
          icon: "GitGraph",
          onClick: (e:any) => {
            this.setState({
              calloutVisible: false
            });
            this.props.changeStage(stage.id);
          }
        }
      );
    }

    // Show stage link, if available
    if (typeof stage.link === "string" && stage.link !== "") {
      stageActions.push(
        {
          key: 'stageLink',
          name: strings.OpenStageLink,
          icon: "Website",
          href: stage.link,
          target: "_blank",
          title: `${strings.OpenLinkTo}: ${stage.link}`
        }
      );
    }

    let historyElement:JSX.Element = null;
    let warningElement:JSX.Element = null;

    // Element detailing history of this stage change
    if (stage.history && (stage.status === StageStatus.History || stage.status === StageStatus.Active)) {

      let historyConstraints:IHistoryConstraintStatus[] = [];

      // Does this history have any constraints?
      if (stage.history.constraints) {

        const constraintNames:string[] = Object.keys(stage.history.constraints);

        if (constraintNames.length > 0) {

          // Available constraints for this stage.
          const constraints:IConstraint[] = getStageConstraints(stage.id, stageSettings.Constraints);

          // Find all matching constraints and filter away null entries
          historyConstraints = constraintNames.map(constraintName => {
            let historyConstraint:IHistoryConstraintStatus = null;
            let matchingConstraint:IConstraint = null;
            
            try {
              matchingConstraint = constraints.filter(constraint => constraint.Title === constraintName)[0];
            } catch (exception) {}

            if (matchingConstraint) {
              historyConstraint = { upholded: stage.history.constraints[constraintName], ...matchingConstraint };
            }

            return historyConstraint;
          }).filter(historyConstraint => historyConstraint !== null);
        }
      }

      if (deadlineOverdue || deadlineComingSoon) {

        const localeFormattedDateString:string = deadline.toLocaleDateString(context.sharePointContext.pageContext.cultureInfo.currentUICultureName);
  
        if (deadlineOverdue) {
  
          warningElement = (
            <p className={`${styles.stageDetailsContent} ${styles.deadlineOverdueColor}`}>
              {strings.DeadlineExceeded}. {strings.DeadlineWas}: <span className="ms-fontWeight-semibold">{localeFormattedDateString}</span>
            </p>
          );
        } else if (deadlineComingSoon) {
          
          warningElement = (
            <p className={`${styles.stageDetailsContent} ${styles.deadlineSoonColor}`}>
              {strings.DeadlineSoon}. {strings.DeadlineIs}: {localeFormattedDateString}
            </p>
          );
        }
      }

      historyElement = (
        <div>
          <p className={styles.stageDetailsContent}>
            {strings.StartedOn}: {stage.history.time.toLocaleString(context.sharePointContext.pageContext.cultureInfo.currentUICultureName)}
          </p>
          <p className={styles.stageDetailsContent}>
            {historyConstraints.length > 0 ? (
              <div className={styles.stageDetailsConstraints}>
                <label className="ms-font-s-plus">{strings.ConstraintsPassedWhenStarted}:</label>
                <ul className={styles.stageDetailsConstraintsList}>
                  {historyConstraints.map(constraint => {
                    return <li className={styles.stageDetailsConstraint}><i className={`ms-Icon ms-Icon--${constraint.upholded ? "CheckMark" : "Warning" }`}></i> {constraint.Title}</li>;
                  })}
                </ul>
              </div>
            ) : null}
          </p>
        </div>
      );
    }

    return (
      <div className={styles.stageDetailsContainer}>
        <div className={styles.stageDetailsHeader}>
          <p className={styles.stageDetailsTitle}>{stageElement}</p>
        </div>
        
        <div className={styles.stageDetailsInner}>
          {warningElement ? warningElement : null}
          {historyElement}
        </div>
        <CommandBar items={stageActions} />
      </div>
    );
  }
}

interface IStageProps extends IStageBaseProps {
  currentEntity: IBusinessModuleEntity;
  changeStage(stageId:string):void;
  addStageFilter(stageId:string):void;
  removeStageFilter():void;
  limitTitleLength:boolean;
}

interface IStageState {
  calloutVisible: boolean;
}

class Stage extends React.Component<IStageProps, IStageState> {

  protected wrapperRef: HTMLElement;

  constructor(props:IStageProps) {
    super(props);

    this.state = {
      calloutVisible: false
    };
  }

  protected setWrapperRef = (node: HTMLElement): void => {
    this.wrapperRef = node;
  }

  private showCallout = ():void => {
    this.setState({
      calloutVisible:true
    });
  }

  private _onDismiss = (ev: any):void => {
    this.setState({
      calloutVisible: false
    });
  }

  protected toggleStageFilter = ():void => {
    if (this.props.isFiltered) {
      this.props.removeStageFilter();
    } else {
      this.props.addStageFilter(this.props.stage.id);
    }
  }

  public render(): JSX.Element {

    let statusClass: string = null;
    let iconClass: string = null;
    let iconModifierClass:string = null;
    let deadlineOverdue:boolean = false;
    let deadlineComingSoon:boolean = false;
    let localeFormattedDateString:string = null;
    let stageName:string = null;
    let stageTitle:string = null;
    let isStageClickable:boolean = false;

    const { isFiltered, stage, currentEntity, deadline, context, limitTitleLength } = this.props;
    const { status } = stage;
    const { calloutVisible } = this.state;

    // Handle shown stage name
    stageName = stage.name;
    stageTitle = stage.name;

    // Limit length of stage name to three characters if rendering in limited space
    if (limitTitleLength) {
      stageName = `${stage.name.charAt(0)}${stage.name.charAt(1)}${stage.name.charAt(2)}`;
    }

    // Can user do anything with this stage?
    if (stage.transitionable || stage.link || stage.history || stage.filterable) {
      isStageClickable = true;
    }

    // Styling controls
    switch (status) {
      case StageStatus.History: {

        statusClass = styles.history;

        if (stage.history || stage.filterable) {
          iconClass = "Accept";
        }
        break;
      }
      case StageStatus.Active: {

        // If stage flow has ended, discontinue normal active rendering
        if (stage.type === StageType.End) {
          statusClass = `${styles.end} ${styles.endColor} ${styles.endBorder}`;
          iconClass = "Accept";
          iconModifierClass = styles.endColor;
          break;
        }

        statusClass = styles.active;
        iconClass = "GitGraph";
        iconModifierClass = styles.activeStageIcon;

        if (deadline) {
          const now = new Date();
          const deadlineMinusOneWeek:Date = Util.dateAdd(deadline, "week", -1);
          localeFormattedDateString = deadline.toLocaleDateString(context.sharePointContext.pageContext.cultureInfo.currentUICultureName);

          // Are we nearing a deadline?
          if (now > deadlineMinusOneWeek) {
            deadlineOverdue = false;
            deadlineComingSoon = true;
          }

          // Have we exceeded the deadline?
          if (now > this.props.deadline) {
            deadlineOverdue = true;
            deadlineComingSoon = false;
          }
        }
        break;
      }
      case StageStatus.Future:
        statusClass = styles.future;
        break;
    }

    // Deadline is overdue
    if (deadlineOverdue) {
      statusClass = `${statusClass} ${styles.deadlineOverdueColor} ${styles.deadlineOverdueBorder}`;
      iconClass = "Error";
      iconModifierClass = styles.deadlineOverdueColor;
      stageTitle = `${strings.DeadlineExceeded}. ${strings.DeadlineWas}: ${localeFormattedDateString}`;
    } else if (deadlineComingSoon) { // Deadline is soon
      statusClass = `${statusClass} ${styles.deadlineSoonColor} ${styles.deadlineSoonBorder}`;
      iconClass = "Clock";
      iconModifierClass = styles.deadlineSoonColor;
      stageTitle = `${strings.DeadlineSoon}. ${strings.DeadlineIs}: ${localeFormattedDateString}`;
    }

    // Show Filter icon for filtered stage
    if (isFiltered) {
      statusClass = `${statusClass} ${styles.filteredBorder}`;
      iconClass = `Filter ${styles.filteredStageIcon}`;
    }

    const stageIconElement:JSX.Element = iconClass ? <i className={`${styles.stageIcon} ms-Icon ms-Icon--${iconClass} ${iconModifierClass}`}></i> : null;
    const stageNameElement:JSX.Element = <span className={styles.stageName}>{stageName}</span>;

    // The combined element to use in the stage details view
    const stageElement:JSX.Element = <div title={stageTitle} className={`${styles.basicStage} ${statusClass}`}>{stageIconElement} {stageNameElement}</div>;

    return (
      <div 
        ref={this.setWrapperRef}
        title={stageTitle}
        onClick={() => { if (isStageClickable) { this.showCallout(); }}}
        className={`${isStageClickable ? styles.stageClickable : styles.unclickableStage} ${statusClass}`}>{stageIconElement} {stageNameElement}{isStageClickable && calloutVisible ? (
        <Callout
          role={'alertdialog'}
          gapSpace={3}
          target={this.wrapperRef}
          onDismiss={this._onDismiss}
        >
          <StageDetails
            stage={stage}
            isFiltered={isFiltered}
            changeStage={this.props.changeStage}
            toggleStageFilter={this.toggleStageFilter}
            context={this.props.context}
            stageSettings={this.props.stageSettings}
            deadline={this.props.deadline}
            stageElement={stageElement}
            deadlineComingSoon={deadlineComingSoon}
            deadlineOverdue={deadlineOverdue}
          />
        </Callout>
      ) : (null)}
      </div>
    );
  }
}

export interface IStageControlProps extends IWorkPointBaseProps {
  stageHistory: IStageHistory[];
  currentEntity: IBusinessModuleEntity;
  businessModuleContentTypes: WorkPointContentType[];
  stageSettings: IStageSettings;
}

export interface IStageControlState {
  filteredStage: string;
  limitStageTitleLength: boolean;
}

const StageDivider:React.SFC = (): JSX.Element => <div className={styles.stageDivider}></div>;

export default class StageControl extends React.Component<IStageControlProps, IStageControlState> {

  private stages: StageObject[];
  protected wrapperRef: HTMLElement;
  protected deadlineDate:Date;

  protected setWrapperRef = (node: HTMLElement): void => {
    this.wrapperRef = node;
  }

  constructor(props: IStageControlProps) {
    super(props);

    this.deadlineDate = null;    

    // Get deadline date for this entity
    const { wp_stageDeadline } = props.currentEntity;
    let deadLineDate:Date = null;

    if (wp_stageDeadline) {
      deadLineDate = new Date(props.currentEntity.wp_stageDeadline);

      if (isValidDate(deadLineDate)) {
        this.deadlineDate = deadLineDate;
      }
    }

    this.stages = getOrderedStageMap(
      props.businessModuleContentTypes,
      props.stageSettings.Stages,
      props.stageSettings.Transitions,
      props.stageSettings.Constraints,
      props.currentEntity.ContentTypeId,
      props.stageHistory
    );

    const filteredStageId:string = storage.session.get(WorkPointStorage.getKey(WorkPointStorageKey.stageFilter, props.context.solutionAbsoluteUrl, props.currentEntity.ListId, props.currentEntity.Id.toString()));

    this.state = {
      filteredStage: filteredStageId || null,
      limitStageTitleLength: false
    };
  }

  /**
   * Sets a stage id mark to be used for filtering items in other controls (webparts, lists etc.)
   * Stage id mark is set in session state, for the current entity and should be cleared when moving away from this entity.
   */
  protected addStageFilter = (stageId:string):void => {

    const storageKey:string = WorkPointStorage.getKey(
      WorkPointStorageKey.stageFilter,
      this.props.context.solutionAbsoluteUrl,
      this.props.currentEntity.ListId,
      this.props.currentEntity.Id.toString()
    );

    // Remove prior stage filters
    storage.session.delete(storageKey);

    // Persisted stage filter in session storage.
    storage.session.put(storageKey, stageId);

    // Post a message across the browser window to let controls know they need to filter on this stage id
    const stageFilterMessage: IAddStageFilterMessage = {
      type: "stage",
      action: "addStageFilter",
      entityListId: this.props.currentEntity.ListId,
      entityItemId: this.props.currentEntity.Id,
      stage: stageId
    };

    this.setState({
      filteredStage: stageId
    });

    window.postMessage(stageFilterMessage, this.props.context.solutionAbsoluteUrl);
  }

  /**
   * Remove stage filtering for a specific business module entity.
   */
  protected removeStageFilter = ():void => {

    const storageKey:string = WorkPointStorage.getKey(
      WorkPointStorageKey.stageFilter,
      this.props.context.solutionAbsoluteUrl,
      this.props.currentEntity.ListId,
      this.props.currentEntity.Id.toString()
    );

    // Remove prior stage filters
    storage.session.delete(storageKey);

    // Post a message across the browser window to let controls know they need to filter on this stage id
    const stageFilterMessage: IBasicStageFilterMessage = {
      type: "stage",
      action: "removeStageFilter",
      entityListId: this.props.currentEntity.ListId,
      entityItemId: this.props.currentEntity.Id
    };

    this.setState({
      filteredStage: null
    });

    window.postMessage(stageFilterMessage, this.props.context.solutionAbsoluteUrl);
  }

  /**
   * Launches the change stage wizard with this stages id selected for the current entity.
   */
  protected changeStage = (stageId:string):void => {

    let dialogArguments:string[] = [];

    const { ListId, Id} = this.props.currentEntity;

    dialogArguments.push(`entityListId=${ListId}&entityItemId=${Id}&stageId=${stageId}`);

    const dialogUrl:string = `ChangeStage?${dialogArguments.join("&")}`;

    const wizardMessage: IWizardStartMessage = {
      type: "workpointwizard",
      url: dialogUrl
    };
    window.postMessage(wizardMessage, this.props.context.solutionAbsoluteUrl);
  }

  protected handleResize = () => {
    if (this.wrapperRef.offsetWidth < this.wrapperRef.scrollWidth) {
      this.setState({
        limitStageTitleLength: true
      });
    }
  }

  public componentDidMount():void {
    window.addEventListener("resize", this.handleResize);
    this.handleResize();
  }

  public componentWillUnmount():void {
    window.removeEventListener("resize", this.handleResize);
  }

  public render(): JSX.Element {

    return (
      <article
        className={styles.stageMenu}
        ref={this.setWrapperRef}
      >
        {this.stages.map((stage: StageObject, index:number) => {

          const isFiltered:boolean = this.state.filteredStage === stage.id;

          return ([
            (index > 0) ? <StageDivider/>: null,
            <Stage
              stage={stage}
              isFiltered={isFiltered}
              changeStage={this.changeStage}
              addStageFilter={this.addStageFilter}
              removeStageFilter={this.removeStageFilter}
              currentEntity={this.props.currentEntity}
              stageSettings={this.props.stageSettings}
              context={this.props.context}
              deadline={this.deadlineDate}
              limitTitleLength={this.state.limitStageTitleLength}
            />
          ]);
        })}
      </article>
    );
  }
}