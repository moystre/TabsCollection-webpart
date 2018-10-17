import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { storage, util } from 'sp-pnp-js/lib/pnp';
import { spODataEntity, spODataEntityArray } from 'sp-pnp-js/lib/sharepoint/odata';
import { Web } from 'sp-pnp-js/lib/sharepoint/webs';
import { IActivity, IActivityResponse } from '../components/Activity';
import { BusinessModuleEntity, BusinessModuleSettingsCollection, IBusinessModuleEntity, IBusinessModuleSettings } from './BusinessModule';
import { BusinessModuleFields, BusinessModuleSelectFields, DefaultListFields } from './Fields';
import { getSolutionRelativeUrl } from './Helper';
import { BasicFieldObject, BasicViewObject, ListObject, SitePageListObject, WorkPointContentType } from './List';
import { IRootObject, WorkPointSettingsCollection, WorkPointSettingsListItem, WORKPOINT_SETTINGS_APP_PRODUCT_ID, WORKPOINT_SETTINGS_BUSINESS_MODULE_SETTINGS_KEY_PREFIX, WORKPOINT_SETTINGS_ENTITY_DETAIL_SETTINGS_KEY_PREFIX, WORKPOINT_SETTINGS_KEY, WORKPOINT_SETTINGS_MYTOOLS_SETTINGS_KEY, WORKPOINT_SETTINGS_TEMPLATE_LIBRARY_MAPPING, WORKPOINT_SETTINGS_VALUE } from './Settings';
import WorkPointStorage, { WorkPointStorageKey } from './Storage';
import { EntityWebProperties, WebProperties, WebProperty } from './WebProperties';

export function getSitePagesByIdAndUrl(listId: string, webUrl: string): Promise<SitePageListObject[]> {
  const web = new Web(webUrl);
  return web.lists.getById(listId).items.select("FieldValuesAsText/FileLeafRef,Title").expand("FieldValuesAsText").filter("startswith(ContentTypeId,'0x0101009D1CB255DA76424F860D91F20E6C4118')").getAs(spODataEntityArray(SitePageListObject));
}

export function getViewsByIdAndUrl(listId: string, webUrl: string): Promise<BasicViewObject[]> {
  const web = new Web(webUrl);
  return web.lists.getById(listId).views.select("Title", "ServerRelativeUrl").filter('Hidden eq false').getAs(spODataEntityArray(BasicViewObject));
}

/**
 * Gets list fields for a specific business module.
 * 
 * Cached.
 * 
 * @param listId Id of the business module to load fields for.
 * @param solutionUrl The WP365 Solution URL.
 */
export function getBusinessModuleListFields(listId:string, solutionUrl:string): Promise<BasicFieldObject[]> {
  
  const web = new Web(solutionUrl);
  const cacheKey:string = WorkPointStorage.getKey(WorkPointStorageKey.businessModuleFields, solutionUrl, listId);
      
  return web.lists.getById(listId).fields
  .select("Title, InternalName, TypeAsString, FieldTypeKind, EntityPropertyName")
  .filter("Hidden eq false")
  .usingCaching({
    expiration: util.dateAdd(new Date(), "day", 1),
    key: cacheKey,
    storeName: "local"
  }).getAs(spODataEntityArray(BasicFieldObject));
}

/**
 * Gets lists for a given web. Filters out unsupported libraries (Wiki Pages) and hidden elements.
 * 
 * @param webAbsoluteUrl Absolute URL of the web to get lists for.
 * @param selectFields Fields to select from lists.
 */
export function getWebLists(webAbsoluteUrl:string, selectFields:string[] = DefaultListFields): Promise<ListObject[]> {
  const web = new Web(webAbsoluteUrl);

  const listFilterString:string = `(Hidden eq false and IsCatalog eq false and IsApplicationList eq false and BaseTemplate ne 119 or (BaseTemplate eq 119 and IsApplicationList eq true and EntityTypeName ne 'Wiki_x0020_Pages'))`;

  // Filter out system lists, but include default site pages gallery
  return web.lists
    .select(...selectFields)
    .filter(listFilterString)
    .getAs(spODataEntityArray(ListObject));
}

/**
 * A combined method fetching both the WorkPoint business module relevant settings and the Root Lists of the solution (Business modules and so fourth).
 * 
 * @description Specialized settings needs to be fetched elsewhere.
 * 
 * @param solutionAbsoluteUrl The solution absolute URL.
 * @param solutionRelativeUrl Relative URL for the solution.
 * @param currentUICultureName String representation of SharePoint locales
 */
export async function getWorkPointSettingsAndRootLists(solutionAbsoluteUrl:string, solutionRelativeUrl:string, currentUICultureName:string): Promise<IRootObject> {
  
  const web = new Web(solutionAbsoluteUrl);

  let batch = web.createBatch();

  let internalPromises:any[] = [];

  const rootListsCacheKey:string = WorkPointStorage.getKey(WorkPointStorageKey.rootLists, solutionAbsoluteUrl);

  const listFilterString:string = `(Hidden eq false and IsCatalog eq false and IsApplicationList eq false and BaseTemplate ne 119 or (BaseTemplate eq 119 and IsApplicationList eq true and EntityTypeName ne 'Wiki_x0020_Pages'))`;

  internalPromises.push(web.lists
    .select(...BusinessModuleSelectFields)
    .expand("ContentTypes")
    // Filter out system lists, but include default site pages gallery
    .filter(listFilterString)
    .inBatch(batch)
    .usingCaching({
      expiration: util.dateAdd(new Date(), "day", 1),
      key: rootListsCacheKey,
      storeName: "local"
    })
    .getAs(spODataEntityArray(ListObject)));

  /* Keys we want:
    + WP_BUSINESS_MODULE_SETTINGS_<bm guid>
    + WP_ENTITY_DETAIL_SETTINGS_<bm guid>
    + WP_MY_TOOLS_SETTINGS
    */
  const settingsListRelativeUrl:string = `${solutionRelativeUrl}/lists/workpointsettings`;

  const settingsCacheKey:string = WorkPointStorage.getKey(WorkPointStorageKey.workPointSettings, solutionAbsoluteUrl);

  internalPromises.push(web.getList(settingsListRelativeUrl).items
    .select(WORKPOINT_SETTINGS_KEY, WORKPOINT_SETTINGS_VALUE)
    .filter(
      `${WORKPOINT_SETTINGS_KEY} eq '${WORKPOINT_SETTINGS_MYTOOLS_SETTINGS_KEY}' or startswith(${WORKPOINT_SETTINGS_KEY}, '${WORKPOINT_SETTINGS_BUSINESS_MODULE_SETTINGS_KEY_PREFIX}') or startswith(${WORKPOINT_SETTINGS_KEY}, '${WORKPOINT_SETTINGS_ENTITY_DETAIL_SETTINGS_KEY_PREFIX}')`
    )
    .inBatch(batch)
    .usingCaching({
      expiration: util.dateAdd(new Date(), "day", 1),
      key: settingsCacheKey,
      storeName: "local"
    })
    .getAs(spODataEntityArray(WorkPointSettingsListItem)));
  
  await batch.execute();
  
  const [rootLists, workPointSettings] = await Promise.all(internalPromises);

  return Promise.resolve({
    rootLists,
    workPointSettingsCollection: new WorkPointSettingsCollection(workPointSettings, currentUICultureName)
  });
}

export function loadEntityParent (parentListId: string, parentId: number, rootSiteCollectionUrl: string, businessModuleSettingsCollection: BusinessModuleSettingsCollection): Promise<IBusinessModuleEntity> {

  const web = new Web(rootSiteCollectionUrl);
  const parentBusinessModuleSettings: IBusinessModuleSettings = businessModuleSettingsCollection.getSettingsForBusinessModule(parentListId);

  let selectProperties: string[] = ["Id", "Title"];

  // Only load wpSite property if parent business module has sites enabled
  if (parentBusinessModuleSettings.SitesEnabled) {
    selectProperties.push(BusinessModuleFields.Site);
  }

  // Only load wpParent property if parent business module has a parent
  if (parentBusinessModuleSettings.Parent) {
    selectProperties.push(BusinessModuleFields.ParentId);
  }

  return web.lists.getById(parentListId).items.select(selectProperties.join(",")).filter(`Id eq ${parentId}`).getAs(spODataEntityArray(BusinessModuleEntity)).then(items => {

    const foundEntity: IBusinessModuleEntity = items[0];
    foundEntity.ListId = parentListId;
    foundEntity.Settings = parentBusinessModuleSettings;

    return foundEntity;
  });
}

/**
 * Load full hierarchy in a promise, containing the child promise calls within
 * 
 * @param currentEntity The current business module entity.
 * @param solutionUrl Absolute URL of the WorkPoint 365 solution.
 * @param businessModuleSettingsCollection Business module settings collection for all business modules.
 * 
 * @returns Promise<IBusinessModuleEntity[]>
 */
export function getEntityHierarchy(currentEntity: IBusinessModuleEntity, solutionUrl:string, businessModuleSettingsCollection: BusinessModuleSettingsCollection): Promise<IBusinessModuleEntity[]> {

  const checkParentSuccess = (entity: IBusinessModuleEntity) => {
    return !entity.wpParentId === true;
  };

  const promise = new Promise<IBusinessModuleEntity[]>((resolve, reject) => {

    let hierarchy: IBusinessModuleEntity[] = [];

    // Recursive parent poller
    const parentPoll = (entity: IBusinessModuleEntity, isDone: boolean = false) => {

      hierarchy.push(entity);
      if (isDone) return Promise.resolve();

      const parentListId: string = entity.Settings.Parent;
      const parentId: number = entity.wpParentId;

      const nextPromise = loadEntityParent(parentListId, parentId, solutionUrl, businessModuleSettingsCollection);
      return nextPromise.then(parentEntity => parentPoll(parentEntity, checkParentSuccess(parentEntity)));
    };

    const result = parentPoll(currentEntity, checkParentSuccess(currentEntity)).then(r => {
      resolve(hierarchy.reverse());
    });
  });

  return promise;
}

export function recycleEntity(webUrl:string, listId:string, itemId:number): Promise<string> {
  const web = new Web(webUrl);

  return web.lists.getById(listId).items.getById(itemId).recycle();
}

/**
 * Fetch web page properties for a given web. 
 * 
 * TODO: Extend this function to use an indexedDB query, so we spare SharePoint request. The structure could look like this:
 * { webUrl: {EntityWebProperties} }
 * 
 * @param webUrl The url of the current web to fetch page properties from
 * @param propertyKeys The requested keys to fetch from the EntityWebProperty enum
 * @returns Promise<EntityWebProperties>
 */
export function getEntityPropertyBagResource(webUrl:string, propertyKeys:string[]): Promise<EntityWebProperties> {
  const web = new Web(webUrl);

  return web.select("AllProperties").select(...propertyKeys).expand("AllProperties").get().then(properties => {
    return new EntityWebProperties(properties);
  });
}

/**
 * Fetcher function for caching solution URL
 * 
 * @private
 * @param siteUrl URL of the site we try to fetch cached Solution URL from
 */
const _getWorkPointRootSiteCollectionUrl = async (siteUrl:string):Promise<string> => {
  const web = new Web(siteUrl);

  let returnSiteUrl:string = null;

  const properties:any = await web.select("AllProperties").select(WebProperty.rootSiteCollection).expand("AllProperties").get();
  
  const webProps = new WebProperties(properties);

  if (typeof webProps.WP_ROOT_SITECOLLECTION === "string" && webProps.WP_ROOT_SITECOLLECTION !== "") {
    returnSiteUrl = webProps.WP_ROOT_SITECOLLECTION;
  } else {
    returnSiteUrl = siteUrl;
  }

  // Remove trailing slashes
  return returnSiteUrl.replace(/\/+$/, "");
};

/**
 * Same as 'getEntityPropertyBagResource', though limited to only fetching the WP_ROOT_SITECOLLECTION url of the web and included caching.
 * 
 * In case of One-site and general multi-site collection, a new request will be made the first time the user visits these. A cache entry will exist for each site-collection as well in the localStorage cache.
 * 
 * 
 * Its worth noting that in case of not finding the property, we assume the user is on the root site colletion page.
 * 
 * Warning - The solution url is heavily cached for a year at a time!
 * 
 * @see getEntityPropertyBagResource
 * 
 * @param siteUrl The url of the current web to fetch page properties from.
 * 
 * @returns Promise<string>
 */
export async function getRootSiteCollectionUrl(siteUrl:string): Promise<string> {

  const storageKey:string = WorkPointStorage.getKey(WorkPointStorageKey.solutionUrl, siteUrl);

  return await storage.local.getOrPut(storageKey, () => _getWorkPointRootSiteCollectionUrl(siteUrl), util.dateAdd(new Date(), "year", 1)).then(url => {
    return url;
  });
}

/**
 * Used to lookup basic information about WorkPoint business module entities.
 * 
 * @param solutionUrl The WorkPoint 365 solution URL.
 * @param itemId Business module list item id.
 * @param listId The id of the business module list (guid).
 * @param selectFields An optional collection of known business module list fields (string[]). If empty or falsy, will return all default fields.
 * 
 * @returns Promise<IBusinessModuleEntity>
 */
export function loadEntityInformation(solutionUrl: string, itemId: number, listId: string, selectFields?: string[]): Promise<IBusinessModuleEntity> {
  const web = new Web(solutionUrl);
  if (!selectFields || selectFields.length === 0) {
    return web.lists.getById(listId).items.getById(itemId).getAs(spODataEntity(BusinessModuleEntity));
  } else {
    return web.lists.getById(listId).items.getById(itemId).select(selectFields.join(",")).getAs(spODataEntity(BusinessModuleEntity));
  }
}

interface IAppResult {
  value:IAppInstance[];
}

interface IAppInstance {
  StartPage:string;
  AppWebFullUrl: string;
}

/**
 * Gets the WorkPoint Add-in app launch parameters, needed for Wizard Dialogs and the Express Panel.
 * 
 * @param solutionUrl URL of the root site collection
 * @param pageContext The pageContext from some instance of ApplicationCustomizer, or FieldCustomizer context.
 * @param spHttpClient SharePoint REST client.
 * 
 * @deprecated Use WPRootSiteCollectionURL instead of context!
 * 
 * @returns Promise<IWorkPointAppLaunchParameters>
 */
export async function getWorkPointAppLaunchParameters(solutionUrl:string, pageContext:PageContext, spHttpClient:SPHttpClient): Promise<IWorkPointAppLaunchParameters> {

  const storageKey:string = WorkPointStorage.getKey(WorkPointStorageKey.appLaunchParameters, pageContext.site.absoluteUrl);

  let launchParameters: IWorkPointAppLaunchParameters = null;

  try {
    launchParameters = await storage.local.getOrPut(storageKey, () => _loadWorkPointAppLaunchParameters(solutionUrl, spHttpClient), util.dateAdd(new Date(), "day", 1));
  } catch (exception) {
    console.warn(`Could not fetch the WorkPoint App Launch parameters. It threw the following error: ${exception}`);
  }

  return launchParameters;
}

export interface IWorkPointAppLaunchParameters {
  appLaunchUrl: string;
  appWebFullUrl: string;
}

/**
 * Loads the WorkPoint Add-in app instance in preferred app order
 * 
 * @description We need this for launching Wizards and the Express panel.
 * 
 * @param solutionAbsoluteUrl URL of the root site collection.
 * @param spHttpClient SharePoint REST client.
 * 
 * @returns Promise<IWorkPointAppLaunchParameters>
 */
async function _loadWorkPointAppLaunchParameters(solutionAbsoluteUrl:string, spHttpClient:SPHttpClient):Promise<IWorkPointAppLaunchParameters> {

  // Try loading the WorkPoint 365 apps in this prioritized order.
  let appProductIds:string[] = [
    "9E966F12-C674-4F72-AA29-F631E3412049", // newAppProductId
    "AD21865C-F9D3-4327-A04F-2599282B7919", // app2ProductId
    "C297DF77-E970-4220-996E-D17FBBDF3C03", // officeAppProductId
    "3DE7157A-F9D5-4788-8C3A-07CBEC497FAC", // rcAppProductId
    "D103F9A9-41CC-4013-B4DC-22311DE9D655", // firstAppProductId
    "CF616A48-FCBA-4478-A7BF-006623240574", // feature
    "5F51073C-8A82-4EE2-B441-6C2B05275E96" // oldAppProductId
  ];

  // If we have App Product Id information saved in WorkPoint Settings, we use that as first priority.
  try {
    let appProductIdInWorkPointSettings:string = await _loadAppProductIdFromSettings(solutionAbsoluteUrl);

    // Remove excessive double quotes.
    appProductIdInWorkPointSettings = appProductIdInWorkPointSettings.replace(/\"/g, "");

    appProductIds.unshift(appProductIdInWorkPointSettings);
  } catch (exception) {}

  /**
   * Recursively iterates over the defined App Product Id's and resolves when we get a match in the defined order.
   * 
   */
  async function _getAppInstance():Promise<IAppInstance> {

    // We have no more product id's to check
    if (appProductIds.length === 0) {
      return null;
    }
      
    try {

      const appId = appProductIds.shift();
    
      const addinUrl:string = `${solutionAbsoluteUrl}/_api/web/getappinstancesbyproductid('${appId}')?$Select=StartPage,AppWebFullUrl`;
  
      const appCandidate:IAppResult = await spHttpClient.get(addinUrl, SPHttpClient.configurations.v1).then((result:SPHttpClientResponse) => {
        return result.json();
      }).catch((exception) => {
        throw `The call to fetch the WorkPoint 365 Add-in with the product id:'${appId}' failed.`;
      });
  
      if (appCandidate.value.length < 1) {
        throw "No app instance found.";
      }

      return appCandidate.value[0];

    } catch (exception) {
      return _getAppInstance();
    }
  }

  const appInstance = await _getAppInstance();

  return _convertAppInstanceToWorkPointAppLaunchParameters(appInstance);
}

function _convertAppInstanceToWorkPointAppLaunchParameters(appInstance:IAppInstance):IWorkPointAppLaunchParameters {

  try {
    let appLaunchUrl:string = "";
    const lastIndexOfSlash = appInstance.StartPage.lastIndexOf('/');
    appLaunchUrl = appInstance.StartPage.substring(0, lastIndexOfSlash);
  
    return { appLaunchUrl, appWebFullUrl: appInstance.AppWebFullUrl};
  } catch (exception) {
    throw "No WorkPoint 365 Add-in could be found.";
  }
}

/**
 * Fetch the App Product Id for the configured WorkPoint 365 app.
 * This setting can be configured in the WorkPoint Settings list, it is however, not always present.
 * 
 * @param solutionAbsoluteUrl URL of the root site collection.
 */
async function _loadAppProductIdFromSettings(solutionAbsoluteUrl:string):Promise<string> {

  try {
  
    const web = new Web(solutionAbsoluteUrl);
  
    const solutionRelativeUrl = getSolutionRelativeUrl(solutionAbsoluteUrl);
    const settingsListRelativeUrl:string = `${solutionRelativeUrl}/lists/workpointsettings`;
  
    const settings:WorkPointSettingsListItem[] = await web.getList(settingsListRelativeUrl).items
      .select(WORKPOINT_SETTINGS_KEY, WORKPOINT_SETTINGS_VALUE)
      .filter(
        `${WORKPOINT_SETTINGS_KEY} eq '${WORKPOINT_SETTINGS_APP_PRODUCT_ID}'`
      )
      .getAs(spODataEntityArray(WorkPointSettingsListItem));
  
    if (!settings || !Array.isArray(settings) || settings.length < 1) {
      throw "No App Product Id stored in WorkPoint settings";
    }

    const appProductIdCandidate = settings[0];
    const appProductIdValue = appProductIdCandidate.wpSettingsValue;

    if (typeof appProductIdValue !== "string" || appProductIdValue === "") {
      throw "App Product Id value is not in correct format.";
    }

    return appProductIdValue;

  } catch (exception) {
    throw exception;
  }
}

export interface ITokenResult {
  Token: IToken;
}

export interface IToken {
  access_token: string;
  expires_on: string;
  resource: string;
}

const _fetchToken = async (tenantUrl:string, sPHttpClient:SPHttpClient):Promise<IToken> => {

  let tokenResult: ITokenResult = null;
  let token: IToken = null;

  try {

    const contextUrl:string = `${tenantUrl}/_api/sphomeservice/context?$expand=Token`;
  
    const result:any = await sPHttpClient.get(contextUrl, SPHttpClient.configurations.v1);
    if (result.ok) {
      tokenResult = await result.json();
      token = tokenResult.Token;
    } else {
      throw "Could not fetch SharePoint context token.";
    }

  } catch (exception) {
      // TODO: Handle?
  }

  return token;
};

export const getActivityForWeb = async (tenantUrl:string, applicationCustomizerContext:ApplicationCustomizerContext):Promise<IActivity[]> => {
  let activities:IActivity[] = [];

  const { spHttpClient } = applicationCustomizerContext;

  const token = await _fetchToken(tenantUrl, spHttpClient);

  if (token) {

    try {

      const siteId = applicationCustomizerContext.pageContext.site.id.toString();
      const webId = applicationCustomizerContext.pageContext.web.id.toString();
      const siteUrl = applicationCustomizerContext.pageContext.site.absoluteUrl;
      const count = 15;

      let requestHeaders: HeadersInit = new Headers();
      requestHeaders.append("authorization", `Bearer ${token.access_token}`);
    
      const fetchUrl:string = `${token.resource}/api/v1/site/activities?SiteId=${siteId}&WebId=${webId}&count=${count}&url=${siteUrl}`;
    
      const activityResponse:any = await spHttpClient.get(fetchUrl, SPHttpClient.configurations.v1, { headers: requestHeaders });

      if (activityResponse.ok) {

        const successFullActivityResponse:IActivityResponse = await activityResponse.json();

        activities = successFullActivityResponse.Activities;

      } else {
        throw "Could not fetch activities for the web.";
      }

    } catch (exception) {
      // TODO: Handle?
    }
  } else {

    // TODO: No token acquired, how do we display this?

  }

  return activities;
};

export interface ITemplateLibrarySettings {
  [key: string]: string; // GUID => GUID
}

/**
 * Gets the Template library mappings from the WorkPoint settings list.
 * 
 * Cached in session storage.
 * 
 * @param solutionAbsoluteUrl Absolute URL of the WorkPoint 365 solution.
 * @param solutionRelativeUrl Relative URL of the WorkPoint 365 solution.
 * 
 * @returns A dictionary of Guids => Guids in the form of 'ITemplateLibrarySettings'.
 */
export const getTemplateLibraryMappings = async (solutionAbsoluteUrl:string, solutionRelativeUrl:string):Promise<ITemplateLibrarySettings> => {

  try {

    const settingsListRelativeUrl: string = `${solutionRelativeUrl}/lists/workpointsettings`;
  
    const settingsCacheKey: string = WorkPointStorage.getKey(WorkPointStorageKey.templateLibraryMappings, solutionAbsoluteUrl);
  
    const web = new Web(solutionAbsoluteUrl);
  
    const settings:WorkPointSettingsListItem[] = await web
      .getList(settingsListRelativeUrl).items
      .select(WORKPOINT_SETTINGS_KEY, WORKPOINT_SETTINGS_VALUE)
      .filter(
        `${WORKPOINT_SETTINGS_KEY} eq '${WORKPOINT_SETTINGS_TEMPLATE_LIBRARY_MAPPING}'`
      )
      .usingCaching({
        expiration: util.dateAdd(new Date(), "day", 1),
        key: settingsCacheKey,
        storeName: "session"
      })
      .getAs(spODataEntityArray(WorkPointSettingsListItem));
  
    if (!settings || !Array.isArray(settings) || settings.length < 1) {
      throw "No Template library settings";
    }

    const templateLibrarySettings:ITemplateLibrarySettings = JSON.parse(settings[0].wpSettingsValue);

    return templateLibrarySettings;

  } catch (exception) {
    return null;
  }
};

/**
 * Fetches only WorkPoint settings.
 * If root lists are wanted, use separate method.
 * 
 * @see getWorkPointSettingsAndRootLists
 * 
 * @param solutionAbsoluteUrl The solution absolute URL.
 * @param solutionRelativeUrl Relative URL for the solution.
 * @param currentUICultureName String representation of SharePoint locales
 */
export async function getWorkPointSettings(solutionAbsoluteUrl:string, solutionRelativeUrl:string, currentUICultureName:string): Promise<WorkPointSettingsCollection> {

  const web = new Web(solutionAbsoluteUrl);
  
  const settingsListRelativeUrl:string = `${solutionRelativeUrl}/lists/workpointsettings`;

  const settingsCacheKey:string = WorkPointStorage.getKey(WorkPointStorageKey.workPointSettings, solutionAbsoluteUrl);
  
  const settings:WorkPointSettingsListItem[] = await web.getList(settingsListRelativeUrl).items
    .select(WORKPOINT_SETTINGS_KEY, WORKPOINT_SETTINGS_VALUE)
    .filter(
      `${WORKPOINT_SETTINGS_KEY} eq '${WORKPOINT_SETTINGS_MYTOOLS_SETTINGS_KEY}' or startswith(${WORKPOINT_SETTINGS_KEY}, '${WORKPOINT_SETTINGS_BUSINESS_MODULE_SETTINGS_KEY_PREFIX}') or startswith(${WORKPOINT_SETTINGS_KEY}, '${WORKPOINT_SETTINGS_ENTITY_DETAIL_SETTINGS_KEY_PREFIX}')`
    )
    .usingCaching({
      expiration: util.dateAdd(new Date(), "day", 1),
      key: settingsCacheKey,
      storeName: "local"
    })
    .getAs(spODataEntityArray(WorkPointSettingsListItem));

  const workPointSettingsCollection = new WorkPointSettingsCollection(settings, currentUICultureName);

  return workPointSettingsCollection;
}

/**
 * TODO: Document
 * 
 * @param solutionAbsoluteUrl 
 * @param templateLibraryListId 
 */
export async function getTemplateLibraryContentTypes(solutionAbsoluteUrl:string, templateLibraryListId:string):Promise<WorkPointContentType[]> {
  
  const web = new Web(solutionAbsoluteUrl);

  const settingsCacheKey:string = WorkPointStorage.getKey(WorkPointStorageKey.templateLibraryContentTypes, solutionAbsoluteUrl);

  const contentTypes:WorkPointContentType[] = await web.lists.getById(templateLibraryListId).contentTypes.select("DocumentTemplateUrl", "Name", "StringId")
    .usingCaching({
      expiration: util.dateAdd(new Date(), "day", 1),
      key: settingsCacheKey,
      storeName: "session"
    }).get();

  return contentTypes;
}