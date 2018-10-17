import { ServiceScope } from "@microsoft/sp-core-library";
import { AadHttpClient } from "@microsoft/sp-http";
import * as webApis from "../../config/webapi-config.json";

export enum SiteStatus {
  Provisioning = "*Provisioning*",
  IgnoreCreateSiteOnItemAddedEvent = "*IgnoreCreateSiteOnItemAddedEvent*",
  Failed = "*Failed*",
  FrontEndQueryFailure = "*FrontEndQueryFailure*"
}

/**
 * Queues site creation for provided entities.
 * 
 * @param entityIds List of entity integer ids, as strings, to create sites for.
 * @param businessModuleId The list GUID of the business module.
 * @param serviceScope Service scope from the extension/webpart.
 * @param solutionUrl Absolute URL of the WorkPoint 365 solution.
 */
export const createEntitySites = async (entityIds:string[], businessModuleId:string, serviceScope:ServiceScope, solutionUrl:string):Promise<boolean> => {

  try {
    let requestHeaders: HeadersInit = new Headers();
    requestHeaders.append("WorkPoint365Url", solutionUrl);
    requestHeaders.append('Content-Type', 'application/json');

    const body = {
      ItemIds: entityIds,
      BusinessModuleId: businessModuleId
    };

    const aadHttpClient: AadHttpClient = new AadHttpClient(serviceScope, webApis[0].id);

    const createSitesResponse = await aadHttpClient.post(`${webApis[0].url}/api/Command/CreateSites`, AadHttpClient.configurations.v1, { headers: requestHeaders, body: JSON.stringify(body) });

    if (createSitesResponse.ok) {
      let createSitesResult:boolean = await createSitesResponse.json();

      if (createSitesResult) {
        return Promise.resolve(true);
      }
    }

    throw "Could not start site creation job.";

  } catch (exception) {
    console.warn(`createEntitySites: ${exception}`);
    return Promise.resolve(false);
  }
};

/**
 * Queues site deletion for provided entities.
 * 
 * @param entityIds List of entity integer ids, as strings, to delete sites for.
 * @param businessModuleId The list GUID of the business module.
 * @param serviceScope Service scope from the extension/webpart.
 * @param solutionUrl Absolute URL of the WorkPoint 365 solution.
 */
export const deleteEntitySites = async (entityIds:string[], businessModuleId:string, serviceScope:ServiceScope, solutionUrl:string):Promise<boolean> => {

  try {
    let requestHeaders: HeadersInit = new Headers();
    requestHeaders.append("WorkPoint365Url", solutionUrl);
    requestHeaders.append('Content-Type', 'application/json');

    const body = {
      ItemIds: entityIds,
      BusinessModuleId: businessModuleId
    };

    const aadHttpClient: AadHttpClient = new AadHttpClient(serviceScope, webApis[0].id);

    const createSitesResponse = await aadHttpClient.post(`${webApis[0].url}/api/Command/DeleteSites`, AadHttpClient.configurations.v1, { headers: requestHeaders, body: JSON.stringify(body) });

    if (createSitesResponse.ok) {
      let createSitesResult:boolean = await createSitesResponse.json();

      if (createSitesResult) {
        return Promise.resolve(true);
      }
    }

    throw "Could not start site deletion job.";
  } catch (exception) {
    console.warn(`deleteEntitySites: ${exception}`);
    return Promise.resolve(false);
  }
};