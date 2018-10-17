import { override } from '@microsoft/decorators';
import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
import * as moment from 'moment';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import pnp, { Web } from 'sp-pnp-js/lib/pnp';
import WorkPointNavBar, { INavbarConfigProps } from '../../components/WorkPointNavBar';
import { IWorkPointContext } from '../../components/WorkPointNavBarInterfaces';
import { IBusinessModuleEntity, IBusinessModuleSettings } from '../../workPointLibrary/BusinessModule';
import { IEntityDetailsSettings, IFieldValue } from '../../workPointLibrary/EntityDetails';
import { getFieldMappingType, getFieldMaps, IFieldMap, IFieldMappingType } from '../../workPointLibrary/FieldMappings';
import { BusinessModuleFields } from '../../workPointLibrary/Fields';
import { getSolutionRelativeUrl } from '../../workPointLibrary/Helper';
import { getUserLicenseStatus, IUserLicense } from '../../workPointLibrary/License';
import { BasicFieldObject, ListObject } from '../../workPointLibrary/List';
import * as DataService from '../../workPointLibrary/service';
import { WorkPointSettingsCollection } from '../../workPointLibrary/Settings';
import { EntityWebProperty } from '../../workPointLibrary/WebProperties';


const LOG_SOURCE: string = 'WorkPointApplicationCustomizer';

/* A Custom Action which can be run during execution of a Client Side Application */
export default class WorkPointApplicationCustomizer extends BaseApplicationCustomizer<null> {

  private _topPlaceholder: PlaceholderContent | undefined;
  private _navbarProps: INavbarConfigProps = null;

  @override
  public async onInit(): Promise<void> {

    /**
     * Style to remove SharePoint site header.
     */
    try {
      const headerElementCandidates = document.getElementsByClassName("ms-compositeHeader");
      if (headerElementCandidates.length > 0) {
        const headerElement:HTMLElement = headerElementCandidates[0] as HTMLElement;
        headerElement.style.display = "none";
      }
    } catch (exception) {}

    /**
     * Remove ugly top border for command bar on list pages.
     */
    try {
      const topCommandBarCandidates = document.getElementsByClassName("od-TopBar-commandBar");
      if (topCommandBarCandidates.length > 0) {
        const topCommandBarElement:HTMLElement = topCommandBarCandidates[0] as HTMLElement;
        topCommandBarElement.style.borderTop = "none";
      }
    } catch (exception) {}

    const placeholderAvailable: boolean = this._tryCreatePlaceholder();
    if (!placeholderAvailable) {
      return Promise.resolve<void>();
    }

    // Setup the moment locale, based on the currentUICultureName
    moment.locale(this.context.pageContext.cultureInfo.currentUICultureName);

    // Setup base PnP environment variables
    pnp.setup({
      spfxContext: this.context
    });

    const navbarLoadingProps:INavbarConfigProps = {
      context: { sharePointContext: this.context, solutionRelativeUrl: null, solutionAbsoluteUrl: null, appWebFullUrl: null, appLaunchUrl: null, userLicense: { Status: null, LoginName: null, SolutionUrl: null, Version: null } },
      ready: false,
      workPointSettingsCollection: null,
      rootLists: null,
      currentEntity: null,
    };

    this._renderLoadingPlaceHolder(navbarLoadingProps);
    
    const siteAbsoluteUrl:string = this.context.pageContext.site.absoluteUrl;
    const webAbsoluteUrl:string = this.context.pageContext.web.absoluteUrl;

    const solutionAbsoluteUrl:string = await DataService.getRootSiteCollectionUrl(siteAbsoluteUrl);

    const solutionRelativeUrl = getSolutionRelativeUrl(solutionAbsoluteUrl);

    const appLaunchParameters:DataService.IWorkPointAppLaunchParameters = await DataService.getWorkPointAppLaunchParameters(solutionAbsoluteUrl, this.context.pageContext, this.context.spHttpClient);
    
    const userLicense:IUserLicense = await getUserLicenseStatus(solutionAbsoluteUrl, this.context.serviceScope);

    const workPointContext: IWorkPointContext = { solutionAbsoluteUrl, solutionRelativeUrl, ...appLaunchParameters, sharePointContext: this.context, userLicense };

    const entityWeb:boolean = workPointContext.solutionAbsoluteUrl !== webAbsoluteUrl;

    if (entityWeb) {
      const propertyBagSelectProperties: string[] = [EntityWebProperty.listId, EntityWebProperty.listItemId, EntityWebProperty.itemLocation].map(key => `AllProperties/${key}`);

      const [workPointSettingsAndRootLists, entityWebProperties, webLists] = await Promise.all(
        [
          DataService.getWorkPointSettingsAndRootLists(workPointContext.solutionAbsoluteUrl, workPointContext.solutionRelativeUrl, workPointContext.sharePointContext.pageContext.cultureInfo.currentUICultureName),
          DataService.getEntityPropertyBagResource(webAbsoluteUrl, propertyBagSelectProperties),
          DataService.getWebLists(this.context.pageContext.web.absoluteUrl)
        ]
      );

      // Settings and root lists
      const workPointSettingsCollection: WorkPointSettingsCollection = workPointSettingsAndRootLists.workPointSettingsCollection;
      const rootLists: ListObject[] = workPointSettingsAndRootLists.rootLists;

      // Entity web
      const entityId = entityWebProperties.WP_SITE_PARENT_LIST_ITEM;
      const entityListId = entityWebProperties.WP_SITE_PARENT_LIST;
      const entityItemLocation = entityWebProperties.WP_ITEM_LOCATION;
      const entityBusinessModuleSettings:IBusinessModuleSettings = workPointSettingsCollection.businessModuleSettings.getSettingsForBusinessModule(entityListId);

      // Set current entity settings, now that we know the listId of the entity
      const entityModuleHasParentModule:boolean = entityBusinessModuleSettings.Parent !== null && entityBusinessModuleSettings.Parent !== undefined && entityBusinessModuleSettings.Parent !== "";

      // If staging is enabled, add custom stage field to the query
      const businessModuleHasStagingEnabled:boolean = Boolean(entityBusinessModuleSettings.StageSettings.Enabled && entityBusinessModuleSettings.StageSettings.Stages.length > 0);

      let entitySelectProperties: string[] = [
        BusinessModuleFields.Id,
        BusinessModuleFields.Title,
        BusinessModuleFields.EffectiveBasePermissions,
        BusinessModuleFields.UniqueId,
        // We assume they have a site column
        BusinessModuleFields.Site
      ];
      
      // Get the parent information if available
      if (entityModuleHasParentModule) {
        entitySelectProperties.push(BusinessModuleFields.ParentId);
      }

      // If staging is enabled, we will need these fields loaded
      if (businessModuleHasStagingEnabled) {
        entitySelectProperties.push(BusinessModuleFields.ContentTypeId);
        entitySelectProperties.push(BusinessModuleFields.StageDeadline);
        entitySelectProperties.push(BusinessModuleFields.StageHistory);
      }

      let entity: IBusinessModuleEntity = null;

      // Get fields for Entity Details webpart
      const entityDetailSettings: IEntityDetailsSettings = workPointSettingsCollection.entityDetailsSettings ? workPointSettingsCollection.entityDetailsSettings.getSettingsForBusinessModule(entityListId) : null;

      if (entityDetailSettings && entityDetailSettings.fields && entityDetailSettings.fields.length > 0) {
      
        const availableFields = await DataService.getBusinessModuleListFields(entityListId, workPointContext.solutionAbsoluteUrl);

        const desiredEntityInternalFieldNames: string[] = entityDetailSettings.fields.map(field => field.internalFieldName);

        let entityDetailsAvailableFields: string[] = [];
        
        availableFields.forEach(field => {

          const indexOfField: number = desiredEntityInternalFieldNames.indexOf(field.InternalName);

          if (indexOfField !== -1) {
            desiredEntityInternalFieldNames.splice(indexOfField, 1);

            entityDetailsAvailableFields.push(field.EntityPropertyName);
          }
        });

        const [entityResult, fieldValues] = await Promise.all([
          // Load entity information from the root web
          DataService.loadEntityInformation(
            workPointContext.solutionAbsoluteUrl,
            entityId,
            entityListId,
            entitySelectProperties
          ),
          new Web(workPointContext.solutionAbsoluteUrl).lists.getById(entityListId).items.getById(entityId).fieldValuesAsHTML.select(entityDetailsAvailableFields.join(",")).get()
        ]);

        entity = entityResult;

        const fieldMap: IFieldMap[] = getFieldMaps(entityBusinessModuleSettings.FieldMappingsSettings);

        // Map entity field value objects
        const fieldValuesProcessed: IFieldValue[] = entityDetailsAvailableFields.map(internalName => {

          let displayName: string = null;
          let displayValue: string = null;
          let typeAsString: string = null;

          /**
           * If displayValue is undefined, it might mean that the internalFieldName is incorrectly fetched for list columns. Example: 
           * 
           * Correct internalFieldName:
           * Hello_x0020_Space
           * 
           * Fetched name from FieldValues call:
           * Hello_x005f_x0020_x005f_Space
           * 
           * Try to fetch the value again with manipulated name
           */
          if (fieldValues[internalName] === undefined) {
            const needle:RegExp = /\_x0020\_/g;
            const usableInternalFieldName:string = internalName.replace(needle, "_x005f_x0020_x005f_");
            displayValue = fieldValues[usableInternalFieldName];

            // Still no match? Maybe theres an special characters (underscores) in our internalFieldName
            if (displayValue === undefined) {
              const underScoreNeedle:RegExp = /\_/g;
              const underScoreReplacedInternalFieldName = internalName.replace(underScoreNeedle, "_x005f_");
              displayValue = fieldValues[underScoreReplacedInternalFieldName];
            }
          } else {
            // Get the display value for this field by its internalFieldName
            displayValue = fieldValues[internalName];
          }

          // No value was found, so set an empty string, as this will print (None) later on
          if (typeof displayValue !== "string" || displayValue === "") {
            displayValue = "";
          }

          let fieldMappingType: IFieldMappingType = getFieldMappingType(internalName, fieldMap);
          
          for (let i = 0; i < availableFields.length; i++) {
            const iteration:BasicFieldObject = availableFields[i];

            if (iteration.InternalName === internalName) {
              displayName = iteration.Title;
              typeAsString = iteration.TypeAsString;
              break;
            }
          }

          return {
            internalFieldName: internalName,
            displayName: displayName || internalName,
            value: displayValue,
            type: typeAsString,
            fieldMappingType: fieldMappingType
          };
        });

        entity.FieldValues = fieldValuesProcessed;

      } else {

        // Load entity information from the root web
        // TODO: Use IndexedDB to store/cache this data
        entity = await DataService.loadEntityInformation(
          workPointContext.solutionAbsoluteUrl,
          entityId,
          entityListId,
          entitySelectProperties
        );
      }

      entity.Settings = entityBusinessModuleSettings;
      entity.Lists = webLists;
      entity.ListId = entityListId;
      entity.ItemLocation = entityItemLocation;

      this._navbarProps = {
        currentEntity: entity,
        workPointSettingsCollection,
        rootLists,
        context: workPointContext,
        ready: true
      };

    } else {

      const settingsAndRootListsResponse = await DataService.getWorkPointSettingsAndRootLists(workPointContext.solutionAbsoluteUrl, workPointContext.solutionRelativeUrl, workPointContext.sharePointContext.pageContext.cultureInfo.currentUICultureName);

      // Settings and root lists
      const workPointSettingsCollection: WorkPointSettingsCollection = settingsAndRootListsResponse.workPointSettingsCollection;
      const rootLists: ListObject[] = settingsAndRootListsResponse.rootLists;

      this._navbarProps = {
        workPointSettingsCollection,
        rootLists,
        currentEntity: null,
        context: workPointContext,
        ready: true
      };
    }

    // Added to handle possible changes on the existence of placeholders
    //this.context.placeholderProvider.changedEvent.add(this, this._placeholderChangedEventHandler);


    /**
     * Calling _renderPlaceholders will always be called by SharePoint, as this event is fired immediately on page load.
     */
    this.context.application.navigatedEvent.add(this, (eventArgs) => {
      console.log(`App navigated to: ${JSON.stringify(eventArgs)}`);
      this._renderPlaceHolder();
    });

    return Promise.resolve<void>();
  }

  private _placeholderChangedEventHandler():void {
    console.log("The placeholder changed!");
    // Is this neccessary? We can detect it being called first time on the page, when it is already rendered. Error?
    //this._renderPlaceHolders();
  }

  private _tryCreatePlaceholder ():boolean {

    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, {
        onDispose: this._onDispose
      });
    }

    // The extension should not assume that the expected placeholder is available.
    if (!this._topPlaceholder) {
      return false;
    }

    if (this._topPlaceholder.domElement) {
      return true;
    } else {
      return false;
    }
  }

  private _renderLoadingPlaceHolder(props: INavbarConfigProps): void {
    ReactDom.render(
      React.createElement(
        WorkPointNavBar, { ...props }
      ),
      this._topPlaceholder.domElement as HTMLElement
    );
  }

  private _renderPlaceHolder(): void {

    ReactDom.render(
      React.createElement(
        WorkPointNavBar, { ...this._navbarProps }
      ),
      this._topPlaceholder.domElement as HTMLElement
    );
  }

  private _onDispose(): void {
    console.log('[WorkPointApplicationCustomizer._onDispose] Disposed custom top placeholders.');
    ReactDom.unmountComponentAtNode(this._topPlaceholder.domElement as HTMLElement);
  }
}
