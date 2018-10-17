import { BaseListViewCommandSet, Command, IListViewCommandSetExecuteEventParameters, IListViewCommandSetListViewUpdatedParameters } from '@microsoft/sp-listview-extensibility';
import { ReallyDeleteThisSite } from 'WorkPointStrings';
import { IBusinessModuleSettings } from '../../../workPointLibrary/BusinessModule';
import { getUserLicenseStatus, IUserLicense, UserLicenseStatus } from '../../../workPointLibrary/License';
import { getWorkPointSettings } from '../../../workPointLibrary/service';
import { createEntitySites, deleteEntitySites, SiteStatus } from '../../../workPointLibrary/Site';


export interface ISiteManagementCommandSetProperties {
  disabledCommandIds: string[];
}

export default class SiteManagementCommandSet extends BaseListViewCommandSet<ISiteManagementCommandSetProperties> {

  protected _businessModuleSettings: IBusinessModuleSettings = null;

  public async onInit(): Promise<void> {

    const userLicenseStatus: IUserLicense = await getUserLicenseStatus(this.context.pageContext.site.absoluteUrl, this.context.serviceScope);

    if (userLicenseStatus && userLicenseStatus.Status === UserLicenseStatus.None) {
      return Promise.reject("User is not licensed");
    }

    const workPointSettingsCollection = await getWorkPointSettings(this.context.pageContext.site.absoluteUrl, this.context.pageContext.site.serverRelativeUrl, this.context.pageContext.cultureInfo.currentUICultureName);

    const businessModuleSettings = workPointSettingsCollection.businessModuleSettings.getSettingsForBusinessModule(this.context.pageContext.list.id.toString());


    if (businessModuleSettings && businessModuleSettings.SitesEnabled) {
      this._businessModuleSettings = this._businessModuleSettings;
      return Promise.resolve(null);
    } else {
      return Promise.reject("Sites is not enabled for this business module");
    }
  }

  /**
   * TODO: These update handlers only take care of single selections for the moment.
   * 
   * @param event
   */
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {

    // These commands should be hidden unless exactly one row is selected.
    const addSiteCommand: Command = this.tryGetCommand('ADD_SITE_TO_ENTITY');
    const removeSiteCommand: Command = this.tryGetCommand('REMOVE_SITE_FROM_ENTITY');
    
    if (event.selectedRows && event.selectedRows.length === 1) {

      const entityCandidates:string[] = event.selectedRows.map(item => item.getValueByName("wpSite"));
      const wpSiteValue:string = entityCandidates[0];

      let sitePresent:boolean = false;

      if (!wpSiteValue || wpSiteValue === "" || wpSiteValue === SiteStatus.Failed || wpSiteValue === SiteStatus.IgnoreCreateSiteOnItemAddedEvent) {
        sitePresent = false;
      } else {
        sitePresent = true;
      }

      if (addSiteCommand) {
        addSiteCommand.visible = !sitePresent;
      }

      if (removeSiteCommand) {
        removeSiteCommand.visible = sitePresent;
      }
    } else {

      if (addSiteCommand) {
        addSiteCommand.visible = false;
      }

      if (removeSiteCommand) {
        removeSiteCommand.visible = false;
      }
    }
  }

  public async onExecute(event: IListViewCommandSetExecuteEventParameters):Promise<void> {

    try {

      /**
       * TODO: For now we only handle creating one site at a time, because of scaling issues.
       */
      if (event.selectedRows && event.selectedRows.length === 1) {
  
        const selectedEntities:string[] = event.selectedRows.map(item => {
          const itemId: string = item.getValueByName("ID");
          return itemId;
        });
  
        switch (event.itemId) {
          case 'ADD_SITE_TO_ENTITY':
            await createEntitySites(selectedEntities, this.context.pageContext.list.id.toString(), this.context.serviceScope, this.context.pageContext.site.absoluteUrl);
            window.location.href = window.location.href;
            break;
          case 'REMOVE_SITE_FROM_ENTITY':
            if (confirm(ReallyDeleteThisSite)) {
              await deleteEntitySites(selectedEntities, this.context.pageContext.list.id.toString(), this.context.serviceScope, this.context.pageContext.site.absoluteUrl);
              window.location.href = window.location.href;
            }
            break;
          default:
            throw new Error('Unknown command');
        }
      }

    } catch (exception) {
      console.warn(`SiteManagement command not recognized: ${event.itemId}`);
    }
  }
}
