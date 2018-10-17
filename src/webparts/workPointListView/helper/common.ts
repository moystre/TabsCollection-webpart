import * as pnp from 'sp-pnp-js';
import { OpenWebByIdResult, Web } from 'sp-pnp-js';

export async function getWebSettings(): Promise<any> {
  const webSettingsResult = await pnp.sp.web.select("AllProperties").expand("AllProperties").get();
  let allProperties = webSettingsResult.AllProperties;
  let webSettings = {
    ListID: allProperties.WP_x005f_SITE_x005f_PARENT_x005f_LIST,
    ListItemID: allProperties.WP_x005f_SITE_x005f_PARENT_x005f_LIST_x005f_ITEM,
    SiteID: allProperties.WP_x005f_SITE_x005f_PARENT_x005f_SITECOLLECTION, // The site collection id
    WebID: allProperties.WP_x005f_SITE_x005f_PARENT_x005f_WEB // The sites Id
  };

  return webSettings;
}

export async function getEntityListItem(webSettings: any): Promise<any> {
  return new Promise<any>(resolve => {
    pnp.sp.site.openWebById(webSettings.WebID).then(targetWeb => {
      pnp.sp.site.rootWeb.lists.getById(webSettings.ListID).items.getById(webSettings.ListItemID).get().then((item) => {
        resolve(item);
        //resolve(item.wpParentId);
      });
    });
  });
}

/**
 * TODO: This is never going to work. Make it business module name independent
 * 
 * @param wpParentId 
 */
export async function getCompanyListItem(wpParentId: any): Promise<any> {
  return new Promise<any>(resolve => {
    pnp.sp.site.rootWeb.lists.getByTitle('Companies').items.getById(wpParentId).get().then((item) => {
      resolve(item.wpWebId); // wpSite
    });
  });
}

export async function getParentEntityListItem(webSettings: any, wpParentId: any): Promise<any> {
  return new Promise<any>(resolve => {
    pnp.sp.site.rootWeb.lists.getById(webSettings.ListID).items.getById(wpParentId).get().then(item => {
      resolve(item.wpWebId);
    });
  });
}

export async function getWebById(webId: string): Promise<Web> {
  const webResult:OpenWebByIdResult = await pnp.sp.site.openWebById(webId);
  return Promise.resolve(webResult.web);
}
