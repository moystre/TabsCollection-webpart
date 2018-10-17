import ServiceScope from "@microsoft/sp-core-library/lib/serviceScope/ServiceScope";
import { AadHttpClient } from "@microsoft/sp-http";
import { storage, util } from "sp-pnp-js/lib/pnp";
import * as webApis from '../../config/webapi-config.json';
import WorkPointStorage, { WorkPointStorageKey } from "./Storage";

export enum UserLicenseStatus {
  None,
  Full,
  Limited,
  External
}

export interface IUserLicense {
  Status: UserLicenseStatus;
  LoginName: string;
  Version: string;
  SolutionUrl: string;
}

/**
 * Fetches user license status as convenience method for the userLicense caching
 * 
 * If something goes wrong with the license check, we allow the user to have a license until next valid license check.
 * 
 * @private
 * @param solutionUrl WorkPoint 365 solution URL.
 * @param aadHttpClient AadHttpClient responsible for making the request.
 */
const _fetchUserLicenseStatus = async (solutionUrl:string, aadHttpClient:AadHttpClient):Promise<IUserLicense> => {

  let licenseStatus:IUserLicense = null;

  try {
    let requestHeaders: HeadersInit = new Headers();
    requestHeaders.append("WorkPoint365Url", solutionUrl);
    
    const licenseResult = await aadHttpClient.get(`${webApis[0].url}/api/License/UserLicense`, AadHttpClient.configurations.v1, { headers: requestHeaders});

    if (licenseResult.ok) {
      licenseStatus = await licenseResult.json();
    } else {
      throw "Could not fetch user license.";
    }
    
  } catch (exceptionMessage) {
    licenseStatus = { Status: UserLicenseStatus.Full, LoginName: null, Version: null, SolutionUrl: solutionUrl };
  }
  
  return licenseStatus;
};

/**
 * Gets the user license status, with fallback to giving user full license upon failing to make the request.
 * Cached for 1 hour at a time.
 * 
 * @param solutionUrl WorkPoint 365 solution URL.
 * @param serviceScope Service scope derived from application customizer context (this.context.serviceScope)
 * 
 * @returns Promise<IUserLicense>
 */
export const getUserLicenseStatus = async (solutionUrl:string, serviceScope:ServiceScope):Promise<IUserLicense> => {
  
  const storageKey:string = WorkPointStorage.getKey(WorkPointStorageKey.userLicense, solutionUrl);
    
  const aadHttpClient: AadHttpClient = new AadHttpClient(serviceScope, webApis[0].id);
  
  return await storage.session.getOrPut(storageKey, () => _fetchUserLicenseStatus(solutionUrl, aadHttpClient), util.dateAdd(new Date(), "hour", 1));
};