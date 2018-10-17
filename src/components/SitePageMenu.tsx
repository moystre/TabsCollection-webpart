import * as React from 'react';
import * as strings from 'WorkPointStrings';
import { ListObject } from '../workPointLibrary/List';
import { getSitePagesByIdAndUrl } from '../workPointLibrary/service';
import ContextMenu from './ContextMenu';
import { Icon, IconSize } from './Icon';
import sitePageMenuStyles from './SitePageMenu.module.scss';
import { INavbarConfigProps } from './WorkPointNavBar';

export interface ISitePageMenuItem {
  url: string;
  title: string;
}

export interface ISitePageMenuState {
  opened: boolean;
  menuItems: ISitePageMenuItem[];
}

export interface ISitePageMenuProps extends INavbarConfigProps {
  closeHomeMenu(): void;
}

export class SitePageMenu extends React.Component<ISitePageMenuProps, ISitePageMenuState> {

  private listId: string;
  private listTitle: string;

  constructor(props: ISitePageMenuProps) {
    super(props);

    this.state = {
      opened: false,
      menuItems: []
    };

    const sitePagesListCandidates: ListObject[] = props.rootLists.filter((list) => {
      return list.BaseTemplate === 119;
    });

    if (sitePagesListCandidates.length === 0) {
      return null;
    }

    const sitePagesList: ListObject = sitePagesListCandidates[0];

    this.listId = sitePagesList.Id;
    this.listTitle = sitePagesList.Title;
  }

  public wrapperRef: HTMLElement;

  protected setWrapperRef = (node: HTMLElement): void => {
    this.wrapperRef = node;
  }

  public handleFoldoutClick = async (event: React.SyntheticEvent<HTMLElement>): Promise<void> => {

    const { menuItems, opened } = this.state;

    if (menuItems && menuItems.length > 0) {

      if (opened) {
        this.hide();
      } else {
        this.show();
      }
    } else if (!opened) {

      this.show();

      const sitePages = await getSitePagesByIdAndUrl(this.listId, this.props.context.solutionAbsoluteUrl);
      
      const baseSitePageUrl: string = `${this.props.context.solutionAbsoluteUrl}/SitePages`;

      const sitePageMenuItems: ISitePageMenuItem[] = sitePages.map(sitePage => {
        const title = sitePage.Title;
        return {
          title: title,
          url: `${baseSitePageUrl}/${sitePage.FieldValuesAsText.FileLeafRef}`
        };
      });

      this.setState({
        menuItems: sitePageMenuItems
      });
    }
  }

  protected show = (): void => {
    this.setState({
      opened: true
    });
  }

  protected hide = (): void => {
    this.setState({
      opened: false
    });
  }

  public render(): JSX.Element {

    const { menuItems, opened } = this.state;
    const { context } = this.props;

    // The site pages list could not be found, so we will not render this component
    if (!this.listId) {
      return null;
    }

    return (
      <div ref={this.setWrapperRef} className={sitePageMenuStyles.solutionItem} title={strings.Dashboards}>

        <a href={this.props.context.solutionAbsoluteUrl} className={sitePageMenuStyles.itemContainer}>
          <Icon size={IconSize.tiny} iconClass="HomeSolid" />
          <span className={sitePageMenuStyles.text}>{strings.MyWorkPointSolution}</span>
        </a>
        <div
          className={sitePageMenuStyles.showMore}
          onMouseDown={this.handleFoldoutClick}
        >
          <Icon iconClass="ChevronRight" />
        </div>
        {menuItems && menuItems.length > 0 && opened &&

          <ContextMenu
            target={this.wrapperRef}
            close={this.hide}
          >
            {menuItems.map(item => {
              return (
                <a href={item.url} className={sitePageMenuStyles.item} title={item.title}>
                  <div className={sitePageMenuStyles.itemContainer}>
                    <span className={sitePageMenuStyles.text}>{item.title}</span>
                  </div>
                </a>
              );
            })}
          </ContextMenu>
        }
      </div>
    );
  }
}