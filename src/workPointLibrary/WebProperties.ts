

export interface IPropertiesResult<T> {
  AllProperties: T;
}

/**
 * Common WorkPoint web properties
 */
export enum WebProperty {
  rootSiteCollection = "WP_ROOT_SITECOLLECTION"
}

export interface IWebProperties {
  WP_x005f_ROOT_x005f_SITECOLLECTION: string;
}

export class WebProperties {
  public WP_ROOT_SITECOLLECTION: string;

  constructor (props:IPropertiesResult<IWebProperties>) {

    let webProperties = _removeDoubleQuotesFromWebProperties(props.AllProperties) as IWebProperties;
    
    let wpRootSitecollection:string = webProperties.WP_x005f_ROOT_x005f_SITECOLLECTION;
    this.WP_ROOT_SITECOLLECTION = wpRootSitecollection;
  }
}

function _removeDoubleQuotesFromWebProperties (sourceProperties:IEntityWebHierarchyProperties|IWebProperties):IEntityWebHierarchyProperties|IWebProperties {

  try {

    let tempProperties = sourceProperties;

    for (const key in tempProperties) {
      if (tempProperties.hasOwnProperty(key)) {
        let value = tempProperties[key];

        if (typeof value === "string" && value !== "") {
          tempProperties[key] = value.replace(/\"/g, "");
        }
      }
    }

    return tempProperties;

  } catch (exception) {
    console.warn(`EntityWebProperty quotes replacement faild with the following error: ${exception}`);

    return null;
  }
}

/**
 * Entity web specific properties
 */
export enum EntityWebProperty {
  listItemId = "WP_SITE_PARENT_LIST_ITEM",
  listId = "WP_SITE_PARENT_LIST",
  parentWeb = "WP_SITE_PARENT_WEB",
  parentSiteCollection = "WP_SITE_PARENT_SITECOLLECTION",
  itemLocation = "WP_ITEM_LOCATION"
}

export interface IEntityWebHierarchyProperties {
  /**
   * Business module list id <String Guid>
   */
  WP_x005f_SITE_x005f_PARENT_x005f_LIST: string;
  /**
   * Business module list item id <Number>
   */
  WP_x005f_SITE_x005f_PARENT_x005f_LIST_x005f_ITEM: string;
  /**
   * Hierarchy representation of this entity <String>
   */
  WP_x005f_ITEM_x005f_LOCATION: string;
  /**
   * Id of the Site collection Web <String Guid>
   */
  WP_x005f_SITE_x005f_PARENT_x005f_WEB: string;
  /**
   * Id of the Site collection Site. This is the site collection that the entity resides in, not neccessarily the solution root site collection. <String Guid>
   */
  WP_x005f_SITE_x005f_PARENT_x005f_SITECOLLECTION: string;
}

export class EntityWebProperties {
  public WP_SITE_PARENT_LIST: string;
  public WP_SITE_PARENT_LIST_ITEM: number;
  public WP_ITEM_LOCATION: string;
  public WP_SITE_PARENT_WEB: string;
  public WP_SITE_PARENT_SITECOLLECTION: string;

  constructor(props:IPropertiesResult<IEntityWebHierarchyProperties>) {

    // Remove all quotes in strings, so that legacy upgraded WorkPoint365 solutions will work.
    let entityWebHierarchyProperties = _removeDoubleQuotesFromWebProperties(props.AllProperties) as IEntityWebHierarchyProperties;

    let parentListItemId = parseInt(entityWebHierarchyProperties.WP_x005f_SITE_x005f_PARENT_x005f_LIST_x005f_ITEM);

    this.WP_SITE_PARENT_LIST = typeof entityWebHierarchyProperties.WP_x005f_SITE_x005f_PARENT_x005f_LIST === "string" ? entityWebHierarchyProperties.WP_x005f_SITE_x005f_PARENT_x005f_LIST : null;
    this.WP_SITE_PARENT_LIST_ITEM = parentListItemId ? parentListItemId : null;
    this.WP_ITEM_LOCATION = entityWebHierarchyProperties.WP_x005f_ITEM_x005f_LOCATION;
    this.WP_SITE_PARENT_WEB = entityWebHierarchyProperties.WP_x005f_SITE_x005f_PARENT_x005f_WEB;
    this.WP_SITE_PARENT_SITECOLLECTION = entityWebHierarchyProperties.WP_x005f_SITE_x005f_PARENT_x005f_SITECOLLECTION;
  }
}