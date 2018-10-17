import * as React from 'react';
import { appendFilterValueToUrl } from '../workPointLibrary/Helper';
import { getSitePagesByIdAndUrl, getViewsByIdAndUrl } from '../workPointLibrary/service';
import ContextMenu from './ContextMenu';
import styles from './FoldoutItem.module.scss';
import { Icon } from './Icon';
import { FoldoutType, IFoldoutItemData } from './WorkPointNavBarInterfaces';


export interface IFoldoutItemState {
  opened: boolean;
  menuItems: IFoldoutItemData[];
}

export default class FoldoutItem extends React.Component<IFoldoutItemData, IFoldoutItemState> {

  constructor(props: IFoldoutItemData) {
    super(props);

    this.state = {
      opened: false,
      menuItems: []
    };
  }

  public wrapperRef: HTMLElement;

  protected setWrapperRef = (node: HTMLElement): void => {
    this.wrapperRef = node;
  }

  public handleFoldoutClick = (event: React.SyntheticEvent<HTMLElement>): void => {

    // Stop link from navigating to list
    event.preventDefault();

    const { menuItems, opened } = this.state;
    const { type, list, entity, context } = this.props;

    if (menuItems && menuItems.length > 0) {

      if (opened) {
        this.hide();
      } else {
        this.show();
      }
    } else if (!opened) {

      this.show();

      switch (type) {
        case FoldoutType.listViewCollection:
        case FoldoutType.businessModuleViewCollection: {

          // List view requests
  
          let viewType: FoldoutType;
  
          if (type === FoldoutType.listViewCollection) {
            viewType = FoldoutType.listView;
          } else if (type === FoldoutType.businessModuleViewCollection) {
            viewType = FoldoutType.businessModuleView;
          }
  
          getViewsByIdAndUrl(list.id, list.parentWebUrl).then((listViewResult) => {
  
            const listViews: IFoldoutItemData[] = listViewResult.map((listView) => (
              {
                text: listView.Title,
                title: listView.Title,
                type: viewType,
                url: listView.ServerRelativeUrl,
                entity,
                context
              }
            ));
  
            this.setState({
              menuItems: listViews
            });
          });
          break;
        }

        /**
         * For now this is only used for site pages, but could well be showing other list items soon.
         */
        case FoldoutType.listItemCollection: {

          getSitePagesByIdAndUrl(list.id, list.parentWebUrl).then((sitePageListResult) => {

            const sitePages: IFoldoutItemData[] = sitePageListResult.map((sitePage) => {

              let title = sitePage.Title;

              if (!title) {
                title = sitePage.FieldValuesAsText.FileLeafRef.replace(".aspx", "");
              }

              return {
                text: title,
                title,
                type: FoldoutType.listItem,
                url: sitePage.FieldValuesAsText.FileLeafRef,
                list,
                entity,
                context
              };
            });
  
            this.setState({
              menuItems: sitePages
            });
          });
        }
      }
    }

    event.stopPropagation();
    event.nativeEvent.stopImmediatePropagation();
  }

  protected show = (): void => {
    this.setState({ opened: true });
  }

  protected hide = (): void => {
    this.setState({ opened: false });
  }

  public render(): JSX.Element {

    const { menuItems, opened } = this.state;
    const { text, title, type, iconUrl, iconClass, url, list, entity, context } = this.props;

    let renderElement: JSX.Element = null;

    // Link type item (view, item etc.)
    if (type && url && (type === FoldoutType.listView || type === FoldoutType.businessModuleView)) {
      renderElement = (
        <a ref={this.setWrapperRef} href={this.appendParentFilterToBusinessModuleLists(url, entity.Title, type)} className={styles.foldoutItem} title={title}>
          <div className={styles.foldoutItemContainer}>
            <Icon iconClass={iconClass} iconUrl={iconUrl} />
            {text &&
              <span className={styles.foldoutItemText}>{text}</span>
            }
          </div>
        </a>
      );

      // Show items from this list
    } else if (type === FoldoutType.listItemCollection) {
      renderElement = (
        
        <div ref={this.setWrapperRef} className={styles.foldoutItem} title={title}>
          <div className={styles.foldoutItemContainer}  onClick={this.handleFoldoutClick}>
            <Icon iconClass={iconClass} iconUrl={iconUrl} />
            <span className={styles.foldoutItemText}>{text} ...</span>
          </div>
          {opened && menuItems && menuItems.length > 0 &&
            <ContextMenu
              target={this.wrapperRef}
              close={this.hide}
            >
              {menuItems.map((menuItem) => (
                <FoldoutItem {...menuItem} context={context} />
              ))}
            </ContextMenu>
          }
        </div>
      );

      // List item
    } else if (type === FoldoutType.listItem) {

      const sitePageUrl:string = `${list.parentWebUrl}/SitePages/${url}`;

      renderElement = (
        <a href={sitePageUrl} className={styles.foldoutItem} title={title}>
          <div className={styles.foldoutItemContainer}>
            <Icon iconClass={iconClass} iconUrl={iconUrl} />
            {text &&
              <span className={styles.foldoutItemText}>{text}</span>
            }
          </div>
        </a>
      );
      
      // Business module view item
    } else {

      renderElement = (
        <div ref={this.setWrapperRef} className={styles.foldoutItem} title={title}>
          <a href={this.appendParentFilterToBusinessModuleLists(list.defaultViewUrl, entity.Title, type)} className={styles.foldoutItemContainer}>
            <Icon iconClass={iconClass} iconUrl={iconUrl} />
            {text &&
              <span className={styles.foldoutItemText}>{text}</span>
            }
            <div className={styles.showMore} onClick={this.handleFoldoutClick}>
              <Icon iconClass="ChevronRight" />
            </div>
          </a>
          {opened && menuItems && menuItems.length > 0 &&
            <ContextMenu
              target={this.wrapperRef}
              close={this.hide}
            >
              {menuItems.map((menuItem) => (
                <FoldoutItem {...menuItem} context={context} />
              ))}
            </ContextMenu>
          }
        </div>
      );
    }

    return renderElement;
  }

  private appendParentFilterToBusinessModuleLists = (url: string, parentLookupText: string, type: FoldoutType): string => {
    if (type === FoldoutType.businessModuleViewCollection || type === FoldoutType.businessModuleItemCollection || type === FoldoutType.businessModuleView) {
      return appendFilterValueToUrl(url, "wpParent", parentLookupText);
    } else {
      return url;
    }
  }
}