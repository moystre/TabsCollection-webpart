import { BasePermissions } from 'sp-pnp-js';
import { IBusinessModuleHierarchyElement } from '../components/BusinessModuleHierarchy';
import { IFieldValue } from './EntityDetails';
import { IFieldMappingsSettings } from './FieldMappings';
import { ListItemObject, ListObject } from './List';
import { IStageSettings } from './Stage';

export interface IBusinessModule {
  id: string;
  listUrl: string;
  title: string;
  iconUrl: string;
  titleResources: object;
}

export class BusinessModuleSettingsCollection {

  public settings: IBusinessModuleSettings[];

  constructor(allSettings: IBusinessModuleSettings[]) {
    this.settings = allSettings;
  }

  /**
   * getSettingsForBusinessModule
   */
  public getSettingsForBusinessModule(businessModuleId: string): IBusinessModuleSettings {

    const businessModuleSettings: IBusinessModuleSettings = this.settings.filter(setting => {
      return businessModuleId === setting.Id;
    })[0];

    return businessModuleSettings;
  }

  /**
   * getSubModulesForModule
   */
  public getSubModulesForModule(businessModuleId: string): IBusinessModule[] {
    try {
      return this.settings
        .filter(businessModuleSetting => businessModuleSetting.Parent === businessModuleId)
        .map(setting => ({
          id: setting.Id,
          listUrl: setting.ListUrl,
          title: setting.Title,
          iconUrl: setting.IconUrl,
          titleResources: null
        }));
    } catch (excep) {
      return [];
    }
  }

  public getBusinessModulesHierarchy(): IBusinessModuleHierarchyElement[] {

    let businessModuleSettings: IBusinessModuleSettings[] = this.settings;

    try {

      const convertedBusinessModules = businessModuleSettings.map(bm => ({
        businessModuleSettings: bm,
        indentation: ""
      }));

      let hierarchyBusinessModules: IBusinessModuleHierarchyCompareElement[] = convertedBusinessModules.filter(bmCompareObject => {
        return bmCompareObject.businessModuleSettings.Parent === null || bmCompareObject.businessModuleSettings.Parent === undefined || bmCompareObject.businessModuleSettings.Parent === "";
      });

      let childBusinessModules: IBusinessModuleHierarchyCompareElement[] = convertedBusinessModules.filter(bmCompareObject => {
        return bmCompareObject.businessModuleSettings.Parent !== null && bmCompareObject.businessModuleSettings.Parent !== undefined && bmCompareObject.businessModuleSettings.Parent !== "";
      });

      if (childBusinessModules.length > 0) {

        while (childBusinessModules.length > 0) {

          var i = childBusinessModules.length;

          while (i--) {

            let iteration = childBusinessModules[i];
            let add = this.appendBusinessModuleToParent(iteration, hierarchyBusinessModules);

            if (add) {
              childBusinessModules.splice(i, 1);
              break;
            }
          }
        }
      }

      const businessModuleHierarchy: IBusinessModuleHierarchyElement[] = hierarchyBusinessModules.map(bmCompareObject => ({
        id: bmCompareObject.businessModuleSettings.Id,
        listUrl: bmCompareObject.businessModuleSettings.ListUrl,
        title: bmCompareObject.businessModuleSettings.Title,
        iconUrl: bmCompareObject.businessModuleSettings.IconUrl,
        titleResources: bmCompareObject.businessModuleSettings.TitleResources,
        indentation: bmCompareObject.indentation
      })
      );

      return businessModuleHierarchy;
    } catch (exception) {
      return null;
    }
  }

  private appendBusinessModuleToParent(childBusinessModule: IBusinessModuleHierarchyCompareElement, hierarchyBusinessModules: IBusinessModuleHierarchyCompareElement[]): Boolean {
    for (let i = 0; i < hierarchyBusinessModules.length; i++) {
      let bm = hierarchyBusinessModules[i];

      // Add indentation to business modules
      if (bm.businessModuleSettings.Id === childBusinessModule.businessModuleSettings.Parent) {
        childBusinessModule.indentation = bm.indentation + "\u00A0\u00A0";
        hierarchyBusinessModules.splice(i + 1, 0, childBusinessModule);
        return true;
      }
    }
  }
}

interface IBusinessModuleHierarchyCompareElement {
  businessModuleSettings: IBusinessModuleSettings;
  indentation: string;
}

export interface IBusinessModuleSettings {
  Id: string;
  IconUrl: string;
  Title: string;
  TitleResources: object;
  EntityName: string;
  EntityNameResource: object;
  Parent: string;
  ParentRelationName: string;
  ParentRelationNameResource: object;
  SitesEnabled: boolean;
  ListUrl: string;
  StageSettings: IStageSettings;
  FieldMappingsSettings: IFieldMappingsSettings;
  EnableEMMIntegration: boolean;

  // Lots of settingskeys are still not implemented / not needed
}

export interface IBusinessModuleEntity {
  Title: string;
  Id: number;
  ListId: string;
  ItemLocation?: string; // TODO: Can we force this somehow?
  EffectiveBasePermissions: BasePermissions;
  UniqueId: string;
  Settings: IBusinessModuleSettings;
  Lists: ListObject[];
  wpParentId?: number;
  wpSite?: string;
  ContentTypeId?:string;
  wp_stageHistory?: string;
  wp_stageDeadline?: string; // Date
  FieldValues?: IFieldValue[]; // Used for the processed field values
  //recycle():Promise<string>; // TODO: This is not used, as we could not get i working. See 'ToolMenu.tsx' under: 'case MyToolsCurrentItemActionType.Delete'
}

export class BusinessModuleEntity extends ListItemObject {
  public wpSite: string;
  public wpParentId: number;
  public EffectiveBasePermissions: BasePermissions;
  public Settings: IBusinessModuleSettings;
  public Lists: ListObject[];
  public ListId: string;
  public UniqueId: string;
}