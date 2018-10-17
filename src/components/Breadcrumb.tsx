import * as React from 'react';
import { IBusinessModule, IBusinessModuleEntity } from '../workPointLibrary/BusinessModule';
import { getListByIdOrUrl, getListIcon, getWebAbsoluteUrl } from '../workPointLibrary/Helper';
import { ListObject } from '../workPointLibrary/List';
import { getWebLists } from '../workPointLibrary/service';
import styles from './Breadcrumb.module.scss';
import ContextMenu, { IOnTopElement } from './ContextMenu';
import FoldoutItem from './FoldoutItem';
import { Icon } from './Icon';
import { FoldoutType, IFoldoutItemData, IWorkPointBaseProps } from './WorkPointNavBarInterfaces';

export interface IBreadcrumbData extends IWorkPointBaseProps {
  entity: IBusinessModuleEntity;
  title: string;
  text?: string;
  iconClass?: string;
  iconUrl?: string;
  currentEntity: IBusinessModuleEntity;
  subModules: IBusinessModule[];
  rootLists: ListObject[];
}

export interface IBreadcrumbState {
  opened: boolean;
  menuItems: IFoldoutItemData[];
}

export default class Breadcrumb extends React.Component<IBreadcrumbData, IBreadcrumbState> {

  constructor(props: IBreadcrumbData) {
    super(props);

    this.state = {
      opened: false,
      menuItems: []
    };
  }

  public wrapperRef: HTMLElement;

  public getSiteLists = (webUrl: string): Promise<ListObject[]> => {

    const { currentEntity, entity } = this.props;

    // Current entity already has lists loaded, so resolve them directly
    if (currentEntity.ListId === entity.ListId &&
      currentEntity.Id === entity.Id &&
      currentEntity.Lists &&
      currentEntity.Lists.length > 0) {

        let promise = new Promise<ListObject[]>((resolve) => {
          resolve(currentEntity.Lists);
        });

        return promise;
    }

    const webAbsoluteUrl: string = getWebAbsoluteUrl(this.props.context.solutionAbsoluteUrl, webUrl);
    
    return getWebLists(webAbsoluteUrl);
  }

  protected setWrapperRef = (node: HTMLElement): void => {
    this.wrapperRef = node;
  }

  protected navigateToEntity = (event:any):void => {
    const mouseEvent = event as MouseEvent;
    const { entity } = this.props;
    if (entity.wpSite) {
      window.location.href = entity.wpSite;
    }
  }

  protected handleToggleMenuOnMouseUp = (event:any):void => {

    const mouseEvent = event as MouseEvent;

    // Has CTRL been active during this click, or did user click the middle mouse button?
    const newTab:boolean = mouseEvent.ctrlKey === true || (mouseEvent.button && mouseEvent.button === 1);

    const { menuItems, opened } = this.state;
    const { entity, context, rootLists } = this.props;

    if (newTab && entity.wpSite) {
      window.open(entity.wpSite, "_blank");
      return null;
    }

    if (menuItems && menuItems.length > 0) {

      if (opened) {
        this.hide();
      } else {
        this.show();
      }
    } else {

      this.show();

      this.getSiteLists(entity.wpSite).then((lists) => {

        // Map submodules for this business module
        const businessModuleMenuItems: IFoldoutItemData[] = this.props.subModules.map(businessModule => {

          const rootBusinessModuleList:ListObject = getListByIdOrUrl(rootLists, businessModule.id);

          return {
            text: businessModule.title,
            title: businessModule.title,
            type: FoldoutType.businessModuleViewCollection,
            iconUrl: businessModule.iconUrl,
            list: {
              id: rootBusinessModuleList.Id,
              defaultViewUrl: rootBusinessModuleList.DefaultViewUrl,
              parentWebUrl: context.solutionAbsoluteUrl,
              baseTemplate: rootBusinessModuleList.BaseTemplate
            },
            entity,
            context
          };
        });

        // Map all this sites lists
        const listMenuItems: IFoldoutItemData[] = lists.map((list) => ({
          text: list.Title,
          title: list.Title,
          type: list.BaseTemplate === 119 ? FoldoutType.listItemCollection : FoldoutType.listViewCollection,
          iconClass: getListIcon(list.BaseTemplate),
          list: {
            id: list.Id,
            defaultViewUrl: list.DefaultViewUrl,
            parentWebUrl: getWebAbsoluteUrl(
              context.solutionAbsoluteUrl,
              list.ParentWebUrl
            ),
            baseTemplate: list.BaseTemplate
          },
          entity,
          context
        }));

        const concatenatedArray:IFoldoutItemData[] = [...businessModuleMenuItems, ...listMenuItems].sort((a:IFoldoutItemData, b:IFoldoutItemData) => {
          if (a.list.baseTemplate === 119) { // Sort up the SitePages gallery
            return -1;
          } else if (b.list.baseTemplate === 119) {
            return 1;
          } else {
            return 0;
          }
        });

        this.setState({ menuItems: concatenatedArray });
      });
    }
  }

  protected show = (): void => {
    this.setState({ opened: true });
  }

  protected hide = (): void => {
    this.setState({ opened: false });
  }

  public render(): JSX.Element {
    const opened = this.state.opened;
    const menuItems = this.state.menuItems;

    const { text, title, iconClass, iconUrl, context, currentEntity, entity } = this.props;
    let className: string = styles.breadcrumb;

    // Is current element?
    if (currentEntity && entity && 
      currentEntity.ListId === entity.ListId &&
      currentEntity.Id === entity.Id) {
      className = styles.breadcrumbActive;
    }

    // Element to be sent to the context menu
    const onTopElement:IOnTopElement = {
      title: text,
      iconUrl,
      iconClass,
      url : entity.wpSite
    };

    return (
      <div
        ref={this.setWrapperRef}
        className={className}
        title={title}
      >
        <div
          className={styles.iconAndText}
          onDoubleClick={this.navigateToEntity}
          onMouseUp={this.handleToggleMenuOnMouseUp}
        >
          <Icon iconClass={iconClass} iconUrl={iconUrl} />
          {text &&
            <span className={styles.text}>{text}</span>
          }
        </div>

        {opened && 
          <ContextMenu 
            target={this.wrapperRef}
            onTopElement={onTopElement}
            close={this.hide}
          >
            {menuItems && menuItems.length > 0 && menuItems.map(item => (
              <FoldoutItem {...item} context={context} />
            ))}
          </ContextMenu>
        }
      </div>
    );
  }
}