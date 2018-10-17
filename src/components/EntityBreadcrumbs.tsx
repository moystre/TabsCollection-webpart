import * as React from 'react';
import { storage } from 'sp-pnp-js/lib/pnp';
import { IBusinessModuleEntity } from '../workPointLibrary/BusinessModule';
import { getEntityHierarchy } from '../workPointLibrary/service';
import WorkPointStorage, { WorkPointStorageKey } from '../workPointLibrary/Storage';
import Breadcrumb from './Breadcrumb';
import breadcrumbStyles from './Breadcrumb.module.scss';
import styles from './EntityBreadcrumbs.module.scss';
import ToggleEntityPanelButton from './ToggleEntityPanelButton';
import { INavbarConfigProps } from './WorkPointNavBar';

const BreadcrumbDivider:React.SFC = (): JSX.Element => <div className={`${breadcrumbStyles.breadcrumbDivider} ms-Icon ms-Icon--ChevronRight`}></div>;

export interface IEntityBreadcrumbsProps extends INavbarConfigProps {
  toggleEntityPanel(event:any):void;
  entityPanelShown:boolean;
}

export interface IEntityBreadcrumbsState {
  entities: IBusinessModuleEntity[];
}

export default class EntityBreadcrumbs extends React.Component<IEntityBreadcrumbsProps, IEntityBreadcrumbsState> {

  constructor (props: IEntityBreadcrumbsProps) {
    super(props);
    
    this.state = {
      entities: []
    };
  }

  /**
   * Convenience function for wrapping arguments to 'getEntityHierarchy'
   */
  private _loadHierarchy = () => {
    return  getEntityHierarchy(
      this.props.currentEntity,
      this.props.context.solutionAbsoluteUrl,
      this.props.workPointSettingsCollection.businessModuleSettings
    );
  }

  public async componentDidMount(): Promise<void> {
    const { currentEntity } = this.props;

    // Only load hierarchy if were standing on a business module entity
    if (currentEntity) {
      const entityHierarchy = await storage.session.getOrPut(
        WorkPointStorage.getKey(
          WorkPointStorageKey.parents, 
          this.props.context.solutionAbsoluteUrl,
          currentEntity.ListId, 
          currentEntity.Id.toString()
        )
        , this._loadHierarchy
      );
      
      this.setState({
        entities: entityHierarchy
      });
    }
  }
  public render():JSX.Element {

    const { entities } = this.state;

    /**
     * TODO: Remove this <div> when React is finally updated to version 16 so we can use fragments.
     * 
     * @see https://reactjs.org/docs/fragments.html
     * @version React-v16
     */
    return (
      <div className={styles.entityBreadcrumbs}>
      {entities && entities.length > 0 && entities.map(
        (entity) => {
            return [
                <BreadcrumbDivider />,
                <Breadcrumb {...this.props} entity={entity} subModules={this.props.workPointSettingsCollection.businessModuleSettings.getSubModulesForModule(entity.Settings.Id)} title={entity.Title} text={entity.Title} iconUrl={entity.Settings.IconUrl} />
            ];
          }
      )}
      {entities && entities.length > 0 && <ToggleEntityPanelButton onClick={this.props.toggleEntityPanel} shown={this.props.entityPanelShown} />}
      </div>
    );
  }
}