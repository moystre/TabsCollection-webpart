import * as pnp from 'sp-pnp-js';

export default class ListService {
  constructor(context: any) {
    pnp.setup({
      spfxContext: context
    });     
  }

  public getLists(isParentSite?: boolean): Promise<any[]> {
    let web = pnp.sp.web;
    if(isParentSite) {
      web = pnp.sp.site.rootWeb;
    }
    return new Promise<any>((resolve: (results: any) => void, reject: (error: any) => void): void => {
        web.lists.select("Title,Id").get().then(lists => {
          resolve(lists);
        }).catch((error) => {
          reject(error);
        });
      });
  }

  public getFields(listId: string, isParentSite?: boolean): Promise<any> {  
    let web = pnp.sp.web;
    if(isParentSite) {
      web = pnp.sp.site.rootWeb;
    }  
    return new Promise<any[]>((resolve: (results: any[]) => void, reject: (error: any) => void): void => {
      web.lists.getById(listId).fields.filter('Hidden eq false').get().then(fields => {        
        resolve(fields);            
      }).catch((error) => {
        console.log(error);
        reject(error);
      });
    });
  }

  public getChoiceValues(listId: string, field: string,isParentSite?: boolean): Promise<any> {
    let web = pnp.sp.web;
    if(isParentSite) {
      web = pnp.sp.site.rootWeb;
    }
    return new Promise<any[]>(resolve => {
      web.lists.getById(listId).fields.getByInternalNameOrTitle(field).get().then(result => {
        resolve(result.Choices);
      });
    });
  }

  public getViews(listId: string, isParentSite?: boolean): Promise<any> {    
    let web = pnp.sp.web;
    if(isParentSite) {
      web = pnp.sp.site.rootWeb;
    }
    return new Promise<any[]>((resolve: (results: any[]) => void, reject: (error: any) => void): void => {
      web.lists.getById(listId).views.select("Id", "Title").get().then(views => {
        resolve(views);            
      }).catch((error) => {
        console.log(error);
        reject(error);
      });
    });
  }
}