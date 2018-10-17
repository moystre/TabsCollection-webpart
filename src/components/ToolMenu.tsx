import * as React from 'react';
import { PermissionKind, storage } from 'sp-pnp-js/lib/pnp';
import { spODataEntity, Web } from 'sp-pnp-js/lib/sharepoint';
import * as strings from 'WorkPointStrings';
import { BusinessModuleEntity, IBusinessModuleEntity } from '../workPointLibrary/BusinessModule';
import { Orientation } from '../workPointLibrary/Common';
import { IWizardStartMessage } from '../workPointLibrary/Event';
import { addFavorite, deleteFavorite, FavoriteTypes, getEntityFavoriteStatus, IFavorite } from '../workPointLibrary/Favorite';
import { actionAllowed, getListByIdOrUrl, getWebAbsoluteUrl } from '../workPointLibrary/Helper';
import { UserLicenseStatus } from '../workPointLibrary/License';
import { ListObject, WorkPointContentType } from '../workPointLibrary/List';
import { IMyToolsAdvancedWizardButtonSetting, IMyToolsButtonSetting, IMyToolsCurrentItemActionButtonSetting, IMyToolsCustomScriptButtonSetting, IMyToolsGroupSetting, IMyToolsItemCreationButtonSetting, IMyToolsLinkButtonSetting, IMyToolsNewListItemButtonSetting, IMyToolsNewParentListItemButtonSetting, IMyToolsRelationWizardButtonSetting, IMyToolsSettings, MyToolsButtonType, MyToolsCurrentItemActionType } from '../workPointLibrary/MyTools';
import { getEntityHierarchy, recycleEntity } from '../workPointLibrary/service';
import { WorkPointSettingsCollection } from '../workPointLibrary/Settings';
import { IStage, IStageSettings, StageObject, StageType } from '../workPointLibrary/Stage';
import WorkPointStorage, { WorkPointStorageKey } from '../workPointLibrary/Storage';
import ContextMenu, { IOnTopElement } from './ContextMenu';
import { Icon, IconSize } from './Icon';
import styles from './ToolMenu.module.scss';
import { INavbarConfigProps } from './WorkPointNavBar';
import { IFieldAndValueObject, IWorkPointBaseProps } from './WorkPointNavBarInterfaces';

interface IToolProps {
  className: string;
  iconUrl: string;
  title: string;
}

const NotAllowedTool: React.SFC<IToolProps> = (props): JSX.Element => {

  return (
    <div className={`${props.className} ${styles.notAllowed}`} title={strings.YouDoNotHavePermissionToPerformThisAction}>
      <div className={styles.iconAndText}>
        <Icon size={IconSize.tiny} iconUrl={props.iconUrl} />
        <span className={styles.toolText}>{props.title}</span>
      </div>
    </div>
  );
};

interface ILinkToolsProps extends IToolProps {
  href: string;
  target?: string;
}

const LinkTool: React.SFC<ILinkToolsProps> = (props): JSX.Element => {
  return (
    <a href={props.href} className={props.className} target={props.target} title={props.title}>
      <div className={styles.iconAndText}>
        <Icon size={IconSize.tiny} iconUrl={props.iconUrl} />
        <span className={styles.toolText}>{props.title}</span>
      </div>
    </a>
  );
};

interface IScriptActionToolProps extends IToolProps {
  onClick: any;
}

const ScriptActionTool: React.SFC<IScriptActionToolProps> = (props): JSX.Element => {
  return (
    <div onClick={props.onClick} className={props.className} title={props.title}>
      <div className={styles.iconAndText}>
        <Icon size={IconSize.tiny} iconUrl={props.iconUrl} />
        <span className={styles.toolText}>{props.title}</span>
      </div>
    </div>
  );
};

export interface IFavoriteToggleToolProps extends INavbarConfigProps {
  className: string;
  iconUrl: string;
  title: string;
}

export interface IFavoriteToggleToolState {
  favorite: IFavorite;
}

export class FavoriteToggleTool extends React.Component<IFavoriteToggleToolProps, IFavoriteToggleToolState> {

  constructor (props:IFavoriteToggleToolProps) {
    super(props);

    this.state = {
      favorite: {
        Id: null,
        Title: props.currentEntity.Title,
        RelativeSitePath: props.currentEntity.Title,
        ItemId: props.currentEntity.Id,
        ListId: props.currentEntity.ListId,
        ListName: props.currentEntity.Settings.Title,
        UserEmail: props.context.sharePointContext.pageContext.user.email,
        FavoriteType: FavoriteTypes.Entity,
        RelativeSiteUrl: props.context.sharePointContext.pageContext.web.serverRelativeUrl,
        IsMailList: false
      }
    };
  }

  /**
   * If current entity has a valid favorite id, we know that the toggle operation should remove the favorite.
   * Otherwise, add it.
   */
  protected _toggleFavorite = async ():Promise<void> => {

    const { favorite } = this.state;

    if (favorite.Id) {
      await deleteFavorite(favorite.Id, this.props.context.sharePointContext.serviceScope, this.props.context.solutionAbsoluteUrl);
    } else {
      await addFavorite(favorite, this.props.context.sharePointContext.serviceScope, this.props.context.solutionAbsoluteUrl);      
    }

    return Promise.resolve(null);
  }

  public render ():JSX.Element {

    const { className, title } = this.props;
    const { favorite } = this.state;
    const favorited:boolean = (favorite.Id !== null && favorite.Id !== undefined);

    const statusIcon:string = favorited ? "FavoriteStarFill" : "FavoriteStar";
    let style = favorited ? {color: "#fbb606"} : null;
    
    return (<div onClick={this._toggleFavorite} className={className} title={title}>
      <div className={styles.iconAndText}>
        <Icon style={style} size={IconSize.tiny} iconClass={statusIcon} />
        <span className={styles.toolText}>{title}</span>
      </div>
    </div>);
  }

  /**
   * When this button is mounted, lookup the current entity in the favorite list, to see if we should show it as favorited.
   * This is only run when this button is visible, so a favorite button hidden away in a group will not make this call until it is visible to the user.
   */
  public async componentDidMount ():Promise<void> {
    const { ListId, Id } = this.props.currentEntity;
    const { solutionAbsoluteUrl, solutionRelativeUrl } = this.props.context; 
    const { email } = this.props.context.sharePointContext.pageContext.user;

    const favoriteId:number = await getEntityFavoriteStatus(ListId, Id, solutionAbsoluteUrl, solutionRelativeUrl, email);
    if (favoriteId > 0) {

      this.setState({
        favorite: {...this.state.favorite, Id:favoriteId}
      });
    }
  }
}

interface IToolsFoldoutData {
  icon: string;
  title: string;
  orientation: Orientation;
}

interface IToolFoldoutState {
  opened: boolean;
}

export interface IWizardDialogButtonProps extends IWorkPointBaseProps {
  url: string;
  title: string;
  className: string;
  iconUrl: string;
}

export class WizardDialogButton extends React.Component<IWizardDialogButtonProps, {}> {

  constructor(props: IWizardDialogButtonProps) {
    super(props);
  }

  protected showDialog = (): void => {
    const wizardMessage: IWizardStartMessage = {
      type: "workpointwizard",
      url: this.props.url
    };
    window.postMessage(wizardMessage, this.props.context.solutionAbsoluteUrl);
  }

  public render(): JSX.Element {
    return <ScriptActionTool onClick={this.showDialog} title={this.props.title} className={this.props.className} iconUrl={this.props.iconUrl} />;
  }
}

class ToolsFoldout extends React.Component<IToolsFoldoutData, IToolFoldoutState> {
  constructor(props: IToolsFoldoutData) {
    super(props);

    this.state = {
      opened: false
    };
  }

  public wrapperRef: HTMLElement;

  protected setWrapperRef = (node: HTMLElement): void => {
    this.wrapperRef = node;
  }

  public handleFoldoutClick = (event: React.SyntheticEvent<HTMLElement>): void => {

    if (this.state.opened) {
      this.close();
    } else {
      this.open();
    }
  }

  protected open = (): void => {
    this.setState({
      opened: true
    });
  }

  protected close = (): void => {
    this.setState({
      opened: false
    });
  }

  public render(): JSX.Element {

    const { orientation, title, icon, children } = this.props;
    const { opened } = this.state;

    let className: string;
    let onTopElement: IOnTopElement = null;

    if (orientation === Orientation.Horizontal) {
      className = styles.horizontalBaseItem;
      onTopElement = { title: `${title}...`, iconUrl: icon };
    } else {
      className = styles.verticalBaseItem;
    }

    return (
      <div ref={this.setWrapperRef} className={className}>
        <div onClick={this.handleFoldoutClick} className={styles.iconAndText}>
          <Icon size={IconSize.tiny} iconUrl={icon} />
          <span className={styles.toolText}>{`${title} ...`}</span>
        </div>
        {children && opened &&
          <ContextMenu
            target={this.wrapperRef}
            close={this.close}
            onTopElement={onTopElement}
          >
            {children}
          </ContextMenu>
        }
      </div>
    );
  }
}

interface IToolControlProps extends INavbarConfigProps {
  button: IMyToolsButtonSetting;
  orientation: Orientation;
}

class ToolControl extends React.Component<IToolControlProps> {

  private list: ListObject;
  private type: MyToolsButtonType;
  private button: any;
  private allowed: boolean;

  constructor(props: IToolControlProps) {
    super(props);

    // If were on the root web, then just set weblists = rootLists
    const webLists: ListObject[] = props.currentEntity ? props.currentEntity.Lists : props.rootLists;
    const rootLists: ListObject[] = props.rootLists;
    const { currentEntity, button } = props;

    this.list = null;
    this.type = MyToolsButtonType[button.Type];
    this.allowed = true;

    let listPermissionChecks: PermissionKind[] = [];
    let itemPermissionChecks: PermissionKind[] = [];


    try {

      // Find and set the associated list for this button, if any
      switch (this.type) {
        case MyToolsButtonType.NewRootListItem:
        case MyToolsButtonType.NewBusinessModuleEntity:
        case MyToolsButtonType.NewSiteListItem: {

          this.button = button as IMyToolsNewListItemButtonSetting;
          const listsResource: ListObject[] = (this.type === MyToolsButtonType.NewSiteListItem) ? webLists : rootLists;

          this.list = getListByIdOrUrl(listsResource, null, this.button.ListUrl);

          listPermissionChecks.push(PermissionKind.AddListItems);
          break;
        }
        case MyToolsButtonType.CurrentItemAction: {

          this.button = button as IMyToolsCurrentItemActionButtonSetting;
          this.list = getListByIdOrUrl(rootLists, currentEntity.ListId);

          // Different cases of modifying the current entity
          switch (this.button.ActionType) {
            case MyToolsCurrentItemActionType.Edit:
            case MyToolsCurrentItemActionType.ChangeStage:
              itemPermissionChecks.push(PermissionKind.EditListItems);
              break;
            case MyToolsCurrentItemActionType.Delete:
              itemPermissionChecks.push(PermissionKind.DeleteListItems);
              break;
            case MyToolsCurrentItemActionType.View:
              itemPermissionChecks.push(PermissionKind.ViewListItems);
          }
          break;
        }

        default:
          this.button = button;
      }

      // Run all permission checks for list or for item
      if (this.list && listPermissionChecks.length > 0) {
        this.allowed = actionAllowed(this.list.EffectiveBasePermissions, listPermissionChecks, this.list);
      } else if (currentEntity && itemPermissionChecks.length > 0) {
        this.allowed = actionAllowed(currentEntity.EffectiveBasePermissions, itemPermissionChecks, currentEntity as BusinessModuleEntity);
      }

      // Check for limited user
      if (props.context.userLicense.Status === UserLicenseStatus.Limited) {
        this.allowed = false;
      }

    } catch (exception) {
      this.allowed = false;
    }
  }

  /**
  * Gets a SharePoint form given the default newForm url. Support for redirecting the user afterwards if 'returnUrl' is provided.
  *
  * @param defaultFormUrl The form URL to use from list. Supports both NewForm and EditForm.
  * @param contentType Optionally provide a content type string guid for the new form.
  * @param returnUrl Optionally set a 'redirect on success / error' URL for SharePoint.
  * @param listItemId Optional specfic list item id to edit: number|string
  * @param additionalFormArguments Optional. Support for custom fields and values for the URL in the form of IFieldAndValueObject[].
  *
  * @returns A SharePoint form URL string.
  *
  */
  private getFormUrl = (defaultFormUrl: string, contentType?: string, returnUrl?: string, listItemId?: number | string, additionalFormArguments?: IFieldAndValueObject[]): string => {

    let formUrlParts: string[] = [];

    if (contentType && contentType !== "") {
      formUrlParts.push(`ContentTypeId=${contentType}`);
    }

    // Support for custom 'url after action'
    if (returnUrl && returnUrl !== "") {
      formUrlParts.push(`Source=${returnUrl}`);
    }

    // Used in conjunction with editing list items
    if (listItemId !== null) {

      if ((typeof listItemId === "number" && listItemId > 0) || (typeof listItemId === "string" && listItemId !== "")) {
        formUrlParts.push(`ID=${listItemId}`);
      }
    }

    // Add custom fields and values to url
    if (additionalFormArguments && additionalFormArguments.length > 0) {
      additionalFormArguments.forEach((fieldAndValueObject) => {
        formUrlParts.push(`${fieldAndValueObject.field}=${fieldAndValueObject.value}`);
      });
    }

    return `${defaultFormUrl}${(formUrlParts.length > 0) ? `?${formUrlParts.join("&")}` : ''}`;
  }

  public render(): JSX.Element {

    const { orientation, currentEntity, context } = this.props;

    let toolClass: string = null;

    if (orientation === Orientation.Horizontal) {
      toolClass = styles.horizontalTool;
    } else if (orientation === Orientation.Vertical) {
      toolClass = styles.verticalTool;
    }

    if (!this.allowed) {
      const button = this.button as IMyToolsButtonSetting;
      return <NotAllowedTool className={toolClass} title={button.Title} iconUrl={button.Icon} />;
    }

    /**
    * Switches the different action types and returns a React element
    */
    switch (this.type) {

      /**
      * New list items, either on this web or on the root web
      */
      case MyToolsButtonType.NewRootListItem:
      case MyToolsButtonType.NewSiteListItem: {

        const button = this.button as IMyToolsNewListItemButtonSetting;

        const newFormUrl: string = this.getFormUrl(this.list.DefaultNewFormUrl, button.ContentType);

        return (
          <LinkTool
            href={newFormUrl}
            className={toolClass}
            iconUrl={button.Icon}
            title={button.Title}
          />
        );
      }

      /**
      * New list item on parent web
      */
      case MyToolsButtonType.NewParentListItem: {

        const button = this.button as IMyToolsNewParentListItemButtonSetting;

        const loadHierarchy = () => {
          return getEntityHierarchy(
            currentEntity, this.props.context.solutionAbsoluteUrl,
            this.props.workPointSettingsCollection.businessModuleSettings
          );
        };

        const onClickAction = async (event: MouseEvent) => {

          try {

            // Has CTRL been active during this click, or did user click the middle mouse button?
            const newTab: boolean = event.ctrlKey === true || (event.button && event.button === 1);
            const storageKey: string = WorkPointStorage.getKey(
              WorkPointStorageKey.parents,
              this.props.context.solutionAbsoluteUrl,
              currentEntity.ListId,
              currentEntity.Id.toString()
            );

            const entityHierarchy: IBusinessModuleEntity[] = await storage.session.getOrPut(storageKey, loadHierarchy);

            const targetParentEntityMatches: IBusinessModuleEntity[] = entityHierarchy.filter((entity: IBusinessModuleEntity) => entity.ListId === button.ParentBmId);

            if (targetParentEntityMatches.length < 1) {
              throw "No parent matches this buttons desired target business module.";
            }

            const targetParentEntity: IBusinessModuleEntity = targetParentEntityMatches[0];

            const webAbsoluteUrl: string = getWebAbsoluteUrl(
              this.props.context.solutionAbsoluteUrl,
              targetParentEntity.wpSite
            );
            const web = new Web(webAbsoluteUrl);
            const listUrl: string = `${targetParentEntity.wpSite}/${button.ListUrl}`;

            const list: ListObject = await web.getList(listUrl).select("DefaultNewFormUrl").getAs(spODataEntity(ListObject));

            const newFormUrl: string = this.getFormUrl(list.DefaultNewFormUrl, button.ContentType);
            if (newTab) {
              window.open(newFormUrl, "_blank");
            } else {
              window.location.href = newFormUrl;
            }

          } catch (exception) {
            // TODO: Error handle
          }
        };

        return (
          <ScriptActionTool
            onClick={onClickAction}
            className={toolClass}
            title={button.Title}
            iconUrl={button.Icon}
          />
        );
      }

      /**
      * Creation of new business module entities
      */
      case MyToolsButtonType.NewBusinessModuleEntity: {

        const button: IMyToolsNewListItemButtonSetting = this.button;

        // Parent information
        let additionalFormArguments: IFieldAndValueObject[] = [];

        if (currentEntity) {
          additionalFormArguments.push({
            field: "selectedid",
            value: currentEntity.Id.toString()
          });
        }

        // TODO: The selectedid (wpParent) information is now available in the new form url, so we need to adhere to it in new forms, maybe in the wpParent field customizer?
        const newFormUrl: string = this.getFormUrl(this.list.DefaultNewFormUrl, button.ContentType, null, null, additionalFormArguments);
        const stageSettings: IStageSettings = this.props.workPointSettingsCollection.businessModuleSettings.getSettingsForBusinessModule(this.list.Id).StageSettings;

        // Is staging enabled, are there any stages configured?
        if (stageSettings.Enabled && stageSettings.Stages && stageSettings.Stages.length > 0) {

          const listContentTypes: WorkPointContentType[] = this.list.ContentTypes;

          // TODO: Overhead of using array iterators?
          const stages: StageObject[] = stageSettings.Stages
            .map((stage: IStage) => {

              // Not start stages
              if (stage.Type === StageType.End || stage.Type === StageType.Standard) {
                return null;
              }

              const contentTypes: WorkPointContentType[] = listContentTypes.filter(contentType => contentType.StringId === stage.ContentTypeId);

              if (contentTypes.length < 0) {
                return null;
              }

              const stageContentType: WorkPointContentType = contentTypes[0];

              return new StageObject(
                stageContentType.StringId,
                stage.Type,
                stageContentType.Name
              );
            })
            .filter(stage => stage !== null);

          return (
            <ToolsFoldout orientation={orientation} title={button.Title} icon={button.Icon}>
              {stages && stages.length > 0 && stages.map((stage: StageObject) => {

                const contentTypeNewFormUrl: string = this.getFormUrl(this.list.DefaultNewFormUrl, stage.id, null, null, additionalFormArguments);

                return <LinkTool href={contentTypeNewFormUrl} className={styles.verticalTool} title={stage.name} iconUrl={button.Icon} />;
              }
              )
              }
            </ToolsFoldout>
          );

        } else {
          return <LinkTool href={newFormUrl} className={toolClass} title={button.Title} iconUrl={button.Icon} />;
        }
      }

      /**
      * Handling of customized script actions
      */
      case MyToolsButtonType.CustomScript: {

        const button: IMyToolsCustomScriptButtonSetting = this.button;

        const customActionFunction = () => {

          const apiPackage: any = {
            ...currentEntity && {
              entity: {
                id: currentEntity.Id,
                listId: currentEntity.ListId,
                title: currentEntity.Title,
                webRelativeUrl: currentEntity.wpSite,
                effectiveBasePermissions: currentEntity.EffectiveBasePermissions,
                parentId: currentEntity.wpParentId,
                parentListId: currentEntity.Settings.Parent
              }
            },
            sharePointContext: {
              web: {
                absoluteUrl: context.sharePointContext.pageContext.web.absoluteUrl,
                relativeUrl: context.sharePointContext.pageContext.web.serverRelativeUrl
              },
              site: {
                absoluteUrl: context.sharePointContext.pageContext.site.absoluteUrl,
                relativeUrl: context.sharePointContext.pageContext.site.serverRelativeUrl
              }
            },
            workPointContext: {
              solutionAbsoluteUrl: this.props.context.solutionAbsoluteUrl,
              solutionRelativeUrl: this.props.context.solutionRelativeUrl,
              appLaunchUrl: this.props.context.appLaunchUrl,
              appWebFullUrl: this.props.context.appWebFullUrl
            }
          };

          const codeToExecute: string = button.Code;

          try {
            Function(`"use strict";try{${codeToExecute}} catch(exception) {alert(\`This buttons custom code failed with the following error: \${exception}\`);}`)(apiPackage);
          } catch (exception) {
            alert(`This buttons custom code could not be create. It failed with the following error: ${exception}`);
          }
        };

        return <ScriptActionTool onClick={customActionFunction} className={toolClass} title={button.Title} iconUrl={button.Icon} />;
      }

      /**
      * Handling actions for the current business module entity
      *
      * Open either SharePoint forms, opens relation wizard dialog or deletes an item and reloads the page.
      */
      case MyToolsButtonType.CurrentItemAction: {

        const button: IMyToolsCurrentItemActionButtonSetting = this.button;

        // Different cases of modifying the current entity
        switch (button.ActionType) {

          /**
          * Edit entity
          *
          * Navigates to the list items edit form.
          */
          case MyToolsCurrentItemActionType.Edit: {

            // We want to return to this page when were done editing this entity
            const editFormUrl: string = this.getFormUrl(
              this.list.DefaultEditFormUrl,
              null,
              context.sharePointContext.pageContext.web.absoluteUrl,
              currentEntity.Id
            );

            return <LinkTool href={editFormUrl} className={toolClass} title={button.Title} iconUrl={button.Icon} />;
          }

          /**
          * Change stage action
          *
          * Launches the change stage wizard dialog.
          */
          case MyToolsCurrentItemActionType.ChangeStage: {

            let dialogArguments: string[] = [];

            // If no entity is present, we cannot show the change stage wizard dialog, therefore dont render the MyTools button.
            if (currentEntity === null) {
              return null;
            }

            const { ListId, Id } = currentEntity;

            dialogArguments.push(`entityListId=${ListId}&entityItemId=${Id}`);

            const dialogUrl: string = `ChangeStage?${dialogArguments.join("&")}`;

            return <WizardDialogButton url={dialogUrl} className={toolClass} iconUrl={button.Icon} title={button.Title} context={this.props.context} />;
          }

          /**
          * Delete this entity
          *
          * Doesn't actually delete it, it recycles it.
          */
          case MyToolsCurrentItemActionType.Delete: {
            const onMouseUpAction = async () => {
              if (confirm(`${strings.ReallyRecycleThisEntity} ${strings.IfYouRegretRecyclingEntityItCanBeRestored}`)) {

                try {

                  await recycleEntity(this.props.context.solutionAbsoluteUrl, this.list.Id, currentEntity.Id);

                  if (confirm(`${strings.TheEntityHasBeenRecycled}. ${strings.IfYouRegretRecyclingEntityItCanBeRestored}. ${strings.NavigateToTheBusinessModuleWhereEntityWasRecycled}`)) {
                    window.location.href = this.list.DefaultViewUrl;
                  }

                } catch (exception) {
                  alert(`${strings.TheEntityCouldNotBeRecycled}.`);
                }
              }
            };
            return <ScriptActionTool onClick={onMouseUpAction} className={toolClass} title={button.Title} iconUrl={button.Icon} />;
          }

          /**
          * View
          *
          * This is a theoretical possibility, which is not implemented in the backend. So we skip it for now
          */
          case MyToolsCurrentItemActionType.View: {
            return null;
          }

          /**
           * No matching MyToolsCurrentItemActionType could be found, so the button type is malformed or not yet implemented.
           */
          default: {
            console.warn(`MyTools rendering of unknown MyToolsCurrentItemActionType denied. Type: '${button.ActionType}'.`);
            return null;
          }
        }
      }

      /**
      * Link type buttons
      *
      * Navigates to URL on click.
      */
      case MyToolsButtonType.Link: {
        const button: IMyToolsLinkButtonSetting = this.button;
        return <LinkTool href={button.Url} target={button.Target} className={toolClass} title={button.Title} iconUrl={button.Icon} />;
      }

      /**
      * Add relation buttons
      *
      * Will show relation wizard dialog on click
      */
      case MyToolsButtonType.AddRelation: {

        const button: IMyToolsRelationWizardButtonSetting = this.button;

        let dialogArguments: string[] = [];

        // If no entity is present, we cannot show the wizard dialog correctly, therefore dont render the MyTools button.
        if (!currentEntity) {
          return null;
        }

        const { ListId, Id } = currentEntity;

        dialogArguments.push(`entityListId=${ListId}&entityItemId=${Id}`);
        dialogArguments.push(`relationTypeA=${button.RelationTypeA}`);
        dialogArguments.push(`relationTypeB=${button.RelationTypeB}`);
        dialogArguments.push(`targetBusinessModuleId=${button.BusinessModule}`);

        if ((typeof button.FilterField === "string" && button.FilterField !== "") && (typeof button.FilterValue === "string" && button.FilterValue !== "")) {
          dialogArguments.push(`filterField=${button.FilterField}`);
          dialogArguments.push(`filterValue=${button.FilterValue}`);
        }

        const dialogUrl: string = `Relations?${dialogArguments.join("&")}`;

        return <WizardDialogButton url={dialogUrl} className={toolClass} iconUrl={button.Icon} title={button.Title} context={this.props.context} />;
      }

      /**
      * General open Wizard dialog buttons.
      * Can be either opened using the old Wizard framework (DocumentProvisioning) or the new Wizard framework (Document set or Entity from Email).
      *
      * button.Wizard property can be either a name (eg. DocumentProvisioning) or a GUID (1d6ad226-ac83-473b-be01-cb8978608236). This determines what Wizard framework is being used. Ie. GUID = new framework.
      */
      case MyToolsButtonType.OpenWizard: {
        // Old document wizard
        if (this.button.Wizard === "DocumentProvisioning") {

          const button = this.button as IMyToolsButtonSetting;

          let dialogArguments: string[] = [];

          if (currentEntity !== null) {
            const { ListId, Id } = currentEntity;
            dialogArguments.push(`entityListId=${ListId}&entityItemId=${Id}`);
          }

          const dialogUrl: string = `DocumentProvisioning?${dialogArguments.join("&")}`;

          return <WizardDialogButton url={dialogUrl} className={toolClass} iconUrl={button.Icon} title={button.Title} context={this.props.context} />;

        } else {

          // New Wizard framework
          const button = this.button as IMyToolsAdvancedWizardButtonSetting;

          let dialogArguments: string[] = [];

          if (currentEntity !== null) {
            dialogArguments.push(`BMAId=${currentEntity.ListId}&ItemAId=${currentEntity.Id}`);
          }

          dialogArguments.push(`ButtonId=${button.Id}`);

          const dialogUrl: string = `Wizard?${dialogArguments.join("&")}`;

          return <WizardDialogButton url={dialogUrl} className={toolClass} iconUrl={button.Icon} title={button.Title} context={this.props.context} />;
        }
      }

      /**
      * Dynamic item creation.
      * Opens a Wizard to handle further interactions.
      */
      case MyToolsButtonType.ItemCreationTrigger: {

        const button = this.button as IMyToolsItemCreationButtonSetting;

        let dialogArguments:string[] = [];

        dialogArguments.push(`triggerId=${button.TriggerId}`);

        if (currentEntity !== null) {
          const { ListId, Id } = currentEntity;
          dialogArguments.push(`businessModuleId=${ListId}&entityId=${Id}`);
        }

        const dialogUrl: string = `ItemCreation?${dialogArguments.join("&")}`;

        return <WizardDialogButton url={dialogUrl} className={toolClass} iconUrl={button.Icon} title={button.Title} context={this.props.context} />;
      }
      
      /**
       * Favorite button.
       * Toggles favorite status of current entity (lists as well?)
       */
      case MyToolsButtonType.Favorite: {
        const button = this.button as IMyToolsButtonSetting;
        return <FavoriteToggleTool className={toolClass} iconUrl={button.Icon} title={button.Title} {...this.props} />;
      }

      /**
       * No matching MyTools type could be found, so the button type is malformed or not yet implemented.
       */
      default: {
        console.warn(`MyTools rendering of unknown MyToolsButtonType denied. Type: '${this.type}'.`);
        return null;
      }
    }
  }
}

interface IGroupProps extends INavbarConfigProps {
  group: IMyToolsGroupSetting;
}

class Group extends React.Component<IGroupProps, IOpenableMenuState> {

  constructor(props: IGroupProps) {
    super(props);

    this.state = {
      opened: false
    };
  }

  public wrapperRef: HTMLElement;

  protected setWrapperRef = (node: HTMLElement): void => {
    this.wrapperRef = node;
  }

  protected handleToggleMenuClick = (event: React.SyntheticEvent<HTMLElement>): void => {
    this.setState({ opened: !this.state.opened });
  }

  protected open = (): void => {
    this.setState({
      opened: true
    });
  }

  protected close = (): void => {
    this.setState({
      opened: false
    });
  }

  public render(): JSX.Element {

    const { Buttons, GroupTitle } = this.props.group;
    const { opened } = this.state;

    if (!Array.isArray(Buttons) || Buttons.length < 1) {
      return null;
    }

    return (
      <div onClick={this.handleToggleMenuClick} ref={this.setWrapperRef} className={styles.verticalBaseItem}>
        <div className={styles.itemContent}>
          <span>{GroupTitle}</span>
          <div className={styles.showMore}>
            <Icon iconClass="ChevronRight" />
          </div>
        </div>
        {opened && Buttons && Buttons.length > 0 &&
          <ContextMenu
            target={this.wrapperRef}
            close={this.close}
          >
            {Buttons.map(btn => <ToolControl orientation={Orientation.Vertical} button={btn} {...this.props} />)}
          </ContextMenu>
        }
      </div>
    );
  }
}

export interface IAllToolsProps extends INavbarConfigProps {
  baseGroupButtons: IMyToolsButtonSetting[];
  groups: IMyToolsGroupSetting[];
}

export interface IOpenableMenuState {
  opened: boolean;
}

export class RemainingTools extends React.Component<IAllToolsProps, IOpenableMenuState> {

  constructor(props: IAllToolsProps) {
    super(props);

    this.state = {
      opened: false
    };
  }

  public wrapperRef: HTMLElement;

  protected setWrapperRef = (node: HTMLElement): void => {
    this.wrapperRef = node;
  }

  protected handleToggleMenuClick = () => {
    this.setState({
      opened: !this.state.opened
    });
  }

  protected open = (): void => {
    this.setState({
      opened: true
    });
  }

  protected close = (): void => {
    this.setState({
      opened: false
    });
  }

  public render(): JSX.Element {
    const { groups, baseGroupButtons } = this.props;
    const { opened } = this.state;

    if (groups.length === 0 && baseGroupButtons.length === 0) {
      return null;
    }

    return (
      <div
        onClick={this.handleToggleMenuClick}
        ref={this.setWrapperRef}
        className={styles.remainingTools}
      >

        <div className={styles.iconAndText}>
          <span>{strings.AllActions}</span>
          <span style={{ marginLeft: "5px", fontSize: '0.7em' }} className="ms-Icon ms-Icon--ChevronDown" />
        </div>

        {opened && groups && groups.length > 0 &&
          <ContextMenu
            target={this.wrapperRef}
            close={this.close}
            onTopElement={{ title: strings.AllActions }}
          >
            {baseGroupButtons.map(button => (
              <ToolControl orientation={Orientation.Vertical} button={button} {...this.props} />
            ))}
            {groups.map(group => (
              <Group group={group} {...this.props} />
            ))}
          </ContextMenu>
        }
      </div>
    );
  }
}

export interface IToolMenuState {
  visibleButtons: IMyToolsButtonSetting[];
  remainingGroups: IMyToolsGroupSetting[];
}

export default class ToolMenu extends React.Component<INavbarConfigProps, IToolMenuState> {
  constructor(props: INavbarConfigProps) {
    super(props);
    this.state = this.getMyToolsSettingsState(props.workPointSettingsCollection, props.currentEntity);
  }

  /**
  * TODO: Rewrite it to use 'getDerivedStateFromProps' when SPFx gets to React version 16.
  * @deprecated This will be deprecated in React 17
  * @see https://reactjs.org/docs/react-component.html#static-getderivedstatefromprops
  *
  * @param nextProps INavbarConfigProps
  */
  public componentWillReceiveProps(nextProps: INavbarConfigProps): void {
    this.setState(this.getMyToolsSettingsState(nextProps.workPointSettingsCollection, nextProps.currentEntity));
  }

  public render(): JSX.Element {

    // Filter so only viewport visible buttons are in this group
    const { visibleButtons, remainingGroups } = this.state;

    return (
      <div className={styles.toolMenu}>
        <div className={styles.toolsContainer}>
          {visibleButtons.map(button => {
            return <ToolControl orientation={Orientation.Horizontal} button={button} {...this.props} />;
          })}
        </div>
        <RemainingTools groups={remainingGroups} baseGroupButtons={visibleButtons} {...this.props} />
      </div>
    );
  }

  /**
  * Fetch the MyTools settings given a WorkPointSettings collection and an entity.
  *
  * @param settingsCollection WorkPoint settings collection to work on
  * @param currentEntity The current entity, or null if none
  *
  * @returns IMyToolsSettings or null
  */
  private getMyToolsSettingsState(settingsCollection: WorkPointSettingsCollection, currentEntity: IBusinessModuleEntity): IToolMenuState {

    let myToolsSettings: IMyToolsSettings;

    try {
      if (currentEntity) {
        myToolsSettings = settingsCollection.myToolsSettings.getSettingsForBusinessModule(currentEntity.ListId);
      } else {
        myToolsSettings = settingsCollection.myToolsSettings.getSettingsForSolution();
      }

      if (myToolsSettings.Groups.length > 0 && myToolsSettings.Groups[0].Buttons.length > 0) {

        return {
          visibleButtons: myToolsSettings.Groups[0].Buttons,
          remainingGroups: myToolsSettings.Groups.slice(1)
        };

      } else {
        throw "MyTools settings are not valid for this use.";
      }

    } catch (exception) {

      return {
        visibleButtons: [],
        remainingGroups: []
      };
    }
  }
}