/**
 * Storage keys to be used in either localStorage or sessionStorage
 */
export enum WorkPointStorageKey {
  businessModuleFields = "WPBMF",
  rootLists = "WPRL",
  parents = "WPP",
  entityPanelHeight = "WPEPH",
  workPointSettings = "WPS",
  templateLibraryMappings = "WPTLM",
  templateLibraryContentTypes = "WPTLCT",
  appLaunchParameters = "WPALP",
  solutionUrl = "WPSU",
  expressPanelWidth = "WPEPW",
  userLicense = "WPUL",
  stageFilter = "WPSF",
  entityPanelVisibility = "WPEPV",
}

/**
 * Clears a particular provided storage.
 * 
 * @param storage The storage instance to clear, either localStorage or sessionStorage.
 * @param solutionAbsoluteUrl Absolute URL of the WorkPoint 365 solution.
 */
const _clearStorage = (storage:Storage, solutionAbsoluteUrl:string):void => {
  const storageItems:string[] = Object.keys(storage);
  const ourStorageItems:string[] = storageItems.filter(item => item.indexOf(solutionAbsoluteUrl) !== -1);

  ourStorageItems.forEach(item => storage.removeItem(item));
};

/**
 * Clears users solution independent preferences.
 */
const _clearUserPreferences = ():void => {

  /**
   * Know storage (session and local) keys that are used for preference storing.
   */
  const localStoragePreferences:string[] = [WorkPointStorageKey.expressPanelWidth, WorkPointStorageKey.entityPanelHeight];
  const sessionPreferences:string[] = [WorkPointStorageKey.entityPanelVisibility];
  
  localStoragePreferences.forEach(item => localStorage.removeItem(item));
  sessionPreferences.forEach(item => sessionStorage.removeItem(item));
};

/**
 * Used to interpret raw PnP storage keys.
 */
export interface IRawStorageObject {
  expiration: string;
}

export default class WorkPointStorage {

  /**
   * Get the solution unique storage key to represent a WorkPoint cache element
   * 
   * @param storageKey Given 'WorkPointStorageKey' to associate cache with
   * @param solutionAbsoluteUrl The solution url string (root site collection absolute url)
   * @param additionalProperties Additional string arguments that will be added to storage key string as ".<argument>"
   */
  public static getKey(storageKey:WorkPointStorageKey, solutionAbsoluteUrl:string, ...additionalProperties:string[]):string {
    const remainingProperties:string = (additionalProperties && additionalProperties.length > 0) ? `.${additionalProperties.join(".")}` : "";
    return `${storageKey}.${solutionAbsoluteUrl.toLowerCase()}${remainingProperties}`;
  }

  /**
   * Removes all WorkPoint 365 web client cache, session- and localStorage.
   * 
   * @param solutionAbsoluteUrl Absolute URL of the WorkPoint 365 solution.
   */
  public static recycleAll(solutionAbsoluteUrl:string):void {
    try {
      solutionAbsoluteUrl = solutionAbsoluteUrl.toLowerCase();

      try {
        _clearStorage(localStorage, solutionAbsoluteUrl);
      } catch (exception) {}

      try {
        _clearStorage(sessionStorage, solutionAbsoluteUrl);
      } catch (exception) {}

      try {
        _clearUserPreferences();
      } catch (exception) {}

    } catch (exception) { }
  }
}