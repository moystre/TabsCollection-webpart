import { ListItem } from "@microsoft/microsoft-graph-types";
import { ServiceScope } from "@microsoft/sp-core-library";
import { AadHttpClient } from "@microsoft/sp-http";
import { Web } from "sp-pnp-js/lib/pnp";
import * as webApis from "../../config/webapi-config.json";

export const FavoriteFields = {
  Id: "Id",
  FavoriteListId: "wpFavFavoriteListId",
  FavoriteListName: "wpFavFavoriteListName",
  FavoriteType: "wpFavFavoriteType",
  IsMailList: "wpFavIsMailList",
  EntityId: "wpFavItemId",
  EntityListId: "wpFavListId",
  UserEmail: "wpFavUserEmail",
  ListName: "wpFavListName",
  RelativeSitePath: "wpFavRelativeSitePath",
  RelativeSiteUrl: "wpFavRelativeSiteUrl",
  SortOrder: "wpFavSortOrder",
  Title: "wpFavTitle",
  JournalId: "wpFavJournalId",
  RelativeListPath: "wpFavRelativeListPath",
  RelativeFolderPath: "wpFavRelativeFolderPath",
};

export enum FavoriteTypes
{
    List = 10,
    Folder = 20,
    DocumentSet = 25,
    Email = 30,
    Entity = 40
}

export interface IFavorite {
  Id:number;
  Title:string;
  ListId:string;
  ListName:string;
  ItemId:number;
  UserEmail:string;
  FavoriteType:FavoriteTypes;
  RelativeSitePath:string;
  RelativeSiteUrl:string;
  IsMailList?:boolean;
}

export interface IFavoriteFull extends IFavorite {
  SortOrder:number;
  JournalId?:string;
  RelativeListPath:string;
  RelativeFolderPath:string;
  FavoriteListID?:string;
  FavoriteListName:string;
  Breadcrumb:string;
  Name:string;
  RelativeUrl:string;
  IconUrl:string;
  FavoriteTypeEnum:object;
}

/**
 * Everything above 0 is considered a success.
 * The number above 0 is the id of the newly created favorite.
 */
export enum FavoriteAddResponses {
  FavoriteAlreadyExists = -10,
  CannotFindBusinessModuleEntity = -20
}

/**
 * Responses when deleting a favorite.
 */
export enum FavoriteDeleteResponses {
  FavoriteDoesNotExist = 10,
  FavoriteDoesNotBelongToCurrentUser = 20
}

/**
 * Adds a favorite
 * 
 * @param favorite Favorite object to add.
 * @param serviceScope Service scope from the extension/webpart.
 * @param solutionUrl Absolute URL of the WorkPoint 365 solution.
 */
export const addFavorite = async (favorite:IFavorite, serviceScope:ServiceScope, solutionUrl:string):Promise<FavoriteAddResponses> => {

  try {
    let requestHeaders: HeadersInit = new Headers();
    requestHeaders.append("WorkPoint365Url", solutionUrl);
    requestHeaders.append('Content-Type', 'application/json');

    const body = {
      favorite
    };

    const aadHttpClient: AadHttpClient = new AadHttpClient(serviceScope, webApis[0].id);

    const addFavoriteResponse = await aadHttpClient.post(`${webApis[0].url}/api/Favorites/Add`, AadHttpClient.configurations.v1, { headers: requestHeaders, body: JSON.stringify(body) });

    if (addFavoriteResponse.ok) {
      let addFavoriteResult:FavoriteAddResponses = await addFavoriteResponse.json();

      switch (addFavoriteResult) {
        case FavoriteAddResponses.CannotFindBusinessModuleEntity:
          throw "Could not find business module entity.";
        case FavoriteAddResponses.FavoriteAlreadyExists:
          throw "This item is already favorited.";
      }
      return Promise.resolve(addFavoriteResult);

    } else {
      throw addFavoriteResponse.statusText;
    }

  } catch (exception) {
    console.warn(`Could not add favorite: ${exception}`);
    return Promise.resolve(null);
  }
};

/**
 * Delete a favorite.
 * 
 * @param favoriteId Id of the favorite (SharePoint list item id from the favorite list).
 * @param serviceScope Service scope from the extension/webpart.
 * @param solutionUrl Absolute URL of the WorkPoint 365 solution.
 */
export const deleteFavorite = async (favoriteId:number, serviceScope:ServiceScope, solutionUrl:string):Promise<FavoriteDeleteResponses> => {

  try {
    let requestHeaders: HeadersInit = new Headers();
    requestHeaders.append("WorkPoint365Url", solutionUrl);
    requestHeaders.append('Content-Type', 'application/json');
    requestHeaders.append("favoriteId", favoriteId.toString());

    const aadHttpClient: AadHttpClient = new AadHttpClient(serviceScope, webApis[0].id);

    const deleteFavoriteResponse = await aadHttpClient.fetch(`${webApis[0].url}/api/Favorites/Delete`, AadHttpClient.configurations.v1, { headers: requestHeaders });

    if (deleteFavoriteResponse.ok) {
      let deleteFavoriteResult:FavoriteDeleteResponses = await deleteFavoriteResponse.json();

      switch (deleteFavoriteResult) {
        case FavoriteDeleteResponses.FavoriteDoesNotBelongToCurrentUser:
          throw "Favorite does not belong to this user.";
        case FavoriteDeleteResponses.FavoriteDoesNotExist:
          throw "Favorite does not exist.";
      }
      return Promise.resolve(deleteFavoriteResult);

    } else {
      throw deleteFavoriteResponse.statusText;
    }

  } catch (exception) {
    console.warn(`Could not delete favorite: ${exception}`);
    return Promise.resolve(null);
  }
};

/**
 * Fetches a singular favorite list item id.
 * Only supports entity favorites at the moment.
 * 
 * @param listId Business module list id.
 * @param itemId Business module list item id.
 * @param solutionAbsoluteUrl Absolute URL for the WorkPoint 365 solution.
 * @param solutionRelativeURL Relative URL for the WorkPoint 365 solution.
 * @param userEmail User email.
 * 
 * @returns Promise containing item id from the matching favorite in the favorites list or 0 if not found.
 */
export const getEntityFavoriteStatus = async (listId:string, itemId:number, solutionAbsoluteUrl:string, solutionRelativeURL:string, userEmail:string):Promise<number> => {
  
  try {
    const favoritesListRelativeUrl: string = `${solutionRelativeURL}/lists/Favorites`;
    const web = new Web(solutionAbsoluteUrl);
  
    const matchingFavorites:ListItem[] = await web.getList(favoritesListRelativeUrl)
    .items
    .select(FavoriteFields.Id)
    .filter(`${FavoriteFields.FavoriteType} eq ${FavoriteTypes.Entity} and ${FavoriteFields.UserEmail} eq '${userEmail}' and ${FavoriteFields.EntityListId} eq '${listId}' and ${FavoriteFields.EntityId} eq '${itemId}'`)
    .get();

    if (!matchingFavorites) {
      throw "Could not make Favorites call";
    }
    
    if (matchingFavorites.length === 0) {
      return Promise.resolve(0);
    } else {
      return Promise.resolve(matchingFavorites[0][FavoriteFields.Id]);
    }

  } catch (exception) {
    console.warn(`Favorites could not be fetched. The following error was thrown: ${exception}`);
    return Promise.reject(null);
  }
};