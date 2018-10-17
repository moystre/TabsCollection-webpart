import * as React from 'react';
import { BusinessModuleSettingsCollection, IBusinessModule } from '../workPointLibrary/BusinessModule';
import { BasicViewObject } from '../workPointLibrary/List';
import { getViewsByIdAndUrl } from '../workPointLibrary/service';
import businessModuleHierarchyStyles from './BusinessModuleHierarchy.module.scss';
import ContextMenu from './ContextMenu';
import { Icon, IconSize } from './Icon';
import { IWorkPointBaseProps } from './WorkPointNavBarInterfaces';

export interface IBusinessModuleHierarchyElementProps extends IWorkPointBaseProps {
  businessModule: IBusinessModuleHierarchyElement;
}

export interface IBusinessModuleHierarchyElementState {
  opened: boolean;
  menuItems: BasicViewObject[];
}

export interface IBusinessModuleHierarchyElement extends IBusinessModule {
  indentation: string;
}

export class BusinessModuleHierarchyElement extends React.Component<IBusinessModuleHierarchyElementProps, IBusinessModuleHierarchyElementState> {

  constructor(props: IBusinessModuleHierarchyElementProps) {
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
    const { businessModule, context } = this.props;

    if (menuItems && menuItems.length > 0) {

      if (opened) {
        this.hide();
      } else {
        this.show();
      }

    } else if (!opened) {

      this.show();
      getViewsByIdAndUrl(businessModule.id, context.solutionAbsoluteUrl).then(listViewResult => {

        this.setState({
          menuItems: listViewResult
        });
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

    const { businessModule, context } = this.props;
    const { opened } = this.state;
    const menuItems = this.state.menuItems as BasicViewObject[];

    return (
      <div ref={this.setWrapperRef} className={businessModuleHierarchyStyles.foldoutItem} title={businessModule.title}>
        <a href={`${context.solutionAbsoluteUrl}/${businessModule.listUrl}`} className={businessModuleHierarchyStyles.itemContainer}>
          {businessModule.indentation ? (<span>{businessModule.indentation}<Icon size={IconSize.microscopic} iconClass="ChevronLeft" className={businessModuleHierarchyStyles.hierarchyChildIcon} /></span>) : null}
          <Icon size={IconSize.tiny} iconUrl={businessModule.iconUrl} />
          <span className={businessModuleHierarchyStyles.text}>{businessModule.title}</span>
          <div className={businessModuleHierarchyStyles.showMore} onClick={this.handleFoldoutClick}>
            <Icon iconClass="ChevronRight" />
          </div>
        </a>
        {menuItems && menuItems.length > 0 && opened &&

          <ContextMenu
            target={this.wrapperRef}
            close={this.hide}
          >
            {menuItems.map(item => {
              return (
                <a href={item.ServerRelativeUrl} className={businessModuleHierarchyStyles.foldoutItem} title={item.Title}>
                  <div className={businessModuleHierarchyStyles.itemContainer}>
                    <span className={businessModuleHierarchyStyles.text}>{item.Title}</span>
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

export interface IBusinessModuleHierarchyProps extends IWorkPointBaseProps {
  businessModuleSettingsCollection: BusinessModuleSettingsCollection;
}

export interface IBusinessModuleHierarchyState {
  businessModuleHierarchy: IBusinessModuleHierarchyElement[];
}

export class BusinessModuleHierarchy extends React.Component<IBusinessModuleHierarchyProps, IBusinessModuleHierarchyState> {

  constructor(props:IBusinessModuleHierarchyProps) {
    super(props);

    this.state = {
      businessModuleHierarchy: []
    };
  }

  public render(): JSX.Element {

    return (
      <div className={businessModuleHierarchyStyles.container}>
        {this.state.businessModuleHierarchy.map(bm => <BusinessModuleHierarchyElement context={this.props.context} businessModule={bm} />)}
      </div>
    );
  }

  public async componentDidMount():Promise<void> {

    let businessModuleHierarchy: IBusinessModuleHierarchyElement[] = [];

    if (this.props.businessModuleSettingsCollection) {
      businessModuleHierarchy = this.props.businessModuleSettingsCollection.getBusinessModulesHierarchy();
    }
    
    this.setState({
      businessModuleHierarchy
    });
  }
}