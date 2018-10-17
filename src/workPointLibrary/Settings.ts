import { Item } from 'sp-pnp-js';
import { getNumberLocaleFromStringLocale } from '../workPointLibrary/Locales';
import { BusinessModuleSettingsCollection, IBusinessModuleSettings } from './BusinessModule';
import { EntityDetailsSettingsCollection } from './EntityDetails';
import { ListObject } from './List';
import { MyToolsSettingsCollection } from './MyTools';

// Keys related to WorkPointSettings
export const WORKPOINT_SETTINGS_KEY = "wpSettingsKey";
export const WORKPOINT_SETTINGS_VALUE = "wpSettingsValue";
export const WORKPOINT_SETTINGS_BUSINESS_MODULE_SETTINGS_KEY_PREFIX = "WP_BUSINESS_MODULE_SETTINGS_";
export const WORKPOINT_SETTINGS_MYTOOLS_SETTINGS_KEY = "WP_MY_TOOLS_SETTINGS";
export const WORKPOINT_SETTINGS_ENTITY_DETAIL_SETTINGS_KEY_PREFIX = "WP_ENTITY_DETAILS_SETTINGS_";
export const WORKPOINT_SETTINGS_TEMPLATE_LIBRARY_MAPPING = "WP_TEMPLATE_LIBRARY_MAPPING";
export const WORKPOINT_SETTINGS_APP_PRODUCT_ID = "WP_APP_PRODUCT_ID";

export interface IRootObject {
  workPointSettingsCollection: WorkPointSettingsCollection;
  rootLists: ListObject[];
}

export class WorkPointSettingsListItem extends Item {
  public wpSettingsKey: string;
  public wpSettingsValue: string;
}

export class WorkPointSettingsCollection {

  public businessModuleSettings: BusinessModuleSettingsCollection;
  public myToolsSettings: MyToolsSettingsCollection;
  public entityDetailsSettings: EntityDetailsSettingsCollection;

  /**
   * Builds the different parts of the WorkPoint settings collection.
   * 
   * Localizes Title resources if supplied with a culture name.
   * 
   * @param settings List item instances from the WorkPoint Settings list.
   * @param currentUICultureName Optional locale to use for overriding Title resources.
   */
  constructor (settings: WorkPointSettingsListItem[], currentUICultureName?:string) {

    if (!settings || !Array.isArray(settings) || settings.length < 1) {
      throw "No WorkPoint settings";
    }

    // Map business module settings
    const businessModuleSettingCandidates = settings
      .filter(setting => setting.wpSettingsKey.indexOf(WORKPOINT_SETTINGS_BUSINESS_MODULE_SETTINGS_KEY_PREFIX) === 0);

    // Handle that some solutions won't contain business module settings right away, or ever.
    if (businessModuleSettingCandidates.length < 1) {
      console.warn("No business module settings available");
      this.businessModuleSettings = null;
    } else {
      
      const businessModuleSettings: IBusinessModuleSettings[] = businessModuleSettingCandidates.map((setting) => {
        const bmSetting:IBusinessModuleSettings = JSON.parse(setting.wpSettingsValue);
        
        const desiredLocale: number = getNumberLocaleFromStringLocale(currentUICultureName);
  
        if (desiredLocale) {
          
          // Title translations
          if (bmSetting.Title && bmSetting.TitleResources) {
            const desiredTitleTranslation:string = bmSetting.TitleResources[desiredLocale];
            
            if (desiredTitleTranslation && desiredTitleTranslation !== "") {
              bmSetting.Title = desiredTitleTranslation;
            }
          }
  
          // Entity name translations
          if (bmSetting.EntityName && bmSetting.EntityNameResource) {
            const desiredEntityNameTranslation:string = bmSetting.EntityNameResource[desiredLocale];
  
            if (desiredEntityNameTranslation && desiredEntityNameTranslation !== "") {
              bmSetting.EntityName = desiredEntityNameTranslation;
            }
          }
  
          // Has Parent translations
          if (bmSetting.ParentRelationName && bmSetting.ParentRelationNameResource) {
            const desiredParentRelationNameTranslation:string = bmSetting.ParentRelationNameResource[desiredLocale];
    
            if (desiredParentRelationNameTranslation && desiredParentRelationNameTranslation !== "") {
              bmSetting.ParentRelationName = desiredParentRelationNameTranslation;
            }
          }
        }
  
        return bmSetting;
      });
  
      this.businessModuleSettings = new BusinessModuleSettingsCollection(businessModuleSettings);
    }

    // Map MyTools settings
    const myToolsSettingsCandidate = settings.filter((setting) => {
      return setting.wpSettingsKey === WORKPOINT_SETTINGS_MYTOOLS_SETTINGS_KEY;
    });

    if (myToolsSettingsCandidate.length < 1) {
      console.warn("No MyTools settings available");
      this.myToolsSettings = null;
    } else {

      this.myToolsSettings = new MyToolsSettingsCollection(myToolsSettingsCandidate[0].wpSettingsValue, currentUICultureName);
    }

    // Map Entity Details settings
    const entityDetailsSettingsCandidates = settings.filter(setting => setting.wpSettingsKey.indexOf(WORKPOINT_SETTINGS_ENTITY_DETAIL_SETTINGS_KEY_PREFIX) === 0);

    if (entityDetailsSettingsCandidates.length < 1) {
      console.warn("No entity details settings available");
      this.entityDetailsSettings = null;
    } else {

      this.entityDetailsSettings = new EntityDetailsSettingsCollection(entityDetailsSettingsCandidates);
    }

  }
}