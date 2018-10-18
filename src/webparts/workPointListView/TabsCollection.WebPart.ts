import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneButton, PropertyPaneButtonType, PropertyPaneDropdown, PropertyPaneTextField } from '@microsoft/sp-webpart-base';
import * as moment from 'moment';
import { IDropdownOption } from 'office-ui-fabric-react';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Web } from 'sp-pnp-js';
import * as strings from 'WorkPointStrings';
import { BusinessModuleFields } from '../../workPointLibrary/Fields';
import { getWebAbsoluteUrl } from '../../workPointLibrary/Helper';
import { getUserLicenseStatus, IUserLicense, UserLicenseStatus } from '../../workPointLibrary/License';
import { BasicListObject, BasicViewObject, ListViewFieldObject } from '../../workPointLibrary/List';
import * as DataService from '../../workPointLibrary/service';
import { EntityWebProperties, EntityWebProperty } from '../../workPointLibrary/WebProperties';
import { ITabOptions, ITabsCollectionProps, ITabSettings } from './components/ITabsCollection';
import { TabsCollection } from './components/TabsCollection';

export default class TabsCollectionWebart extends BaseClientSideWebPart<any> {
  protected _userLicenseStatus: IUserLicense = null;

  private lists: IDropdownOption[];
  private fieldOptions: IDropdownOption[];
  private entityFieldOptions: IDropdownOption[];
  private viewOptions: IDropdownOption[];

  private solutionAbsoluteURL: string;
  private isRootWeb: boolean;
  private targetWeb: Web;
  private targetWebUrl: string;
  private entityWebProperties: EntityWebProperties;

  public tabProperties: ITabSettings[] = [];
  public tabViewOptions: ITabOptions[] = [];
  public scopeOptions = [{
    key: 'rootSite',
    text: 'Root site'
  }];

  public propertyPaneGroups;

  public finder = '##### ';

  public async onInit(): Promise<void> {
    console.log(this.finder + 'onInit');
    const siteAbsoluteUrl: string = this.context.pageContext.site.absoluteUrl;
    const webAbsoluteUrl: string = this.context.pageContext.web.absoluteUrl;
    this.solutionAbsoluteURL = await DataService.getRootSiteCollectionUrl(siteAbsoluteUrl);
    const propertyBagSelectProperties: string[] = [EntityWebProperty.listId, EntityWebProperty.listItemId, EntityWebProperty.itemLocation].map(key => `AllProperties/${key}`);
    const entityWebProperties = await DataService.getEntityPropertyBagResource(webAbsoluteUrl, propertyBagSelectProperties);
    this.entityWebProperties = entityWebProperties;

    try {
      for (let i = 0; i < 2; i++) {
        this.tabProperties[i] = {
          scope: 'rootSite',
          list: '',
          title: '',
          view: '',
          listName: ''
        };
        this.tabViewOptions[i] = {
          items: []
        };
      }
    } catch (exception) {
      console.log(exception);
    }

    this.targetWebUrl = await this.getTargetWebURL(1);
    this.targetWeb = new Web(this.targetWebUrl);

    this._userLicenseStatus = await getUserLicenseStatus(this.solutionAbsoluteURL, this.context.serviceScope);

    moment.locale(this.context.pageContext.cultureInfo.currentUICultureName);

    return Promise.resolve<void>();
  }

  public render(): void {
    console.log(this.finder + 'render');

    if (this._userLicenseStatus && this._userLicenseStatus.Status === UserLicenseStatus.None) {
      this.domElement.innerHTML = strings.YouHaveNoWorkPoint365License;
      return null;
    }
    const element: React.ReactElement<ITabsCollectionProps> = React.createElement(TabsCollection, {
      tabsArray: this.tabProperties,
      webpartSettings: this.tabProperties[this.tabProperties.length - 1],
      needsConfiguration: this.needsConfiguration(this.tabProperties.length),
      context: this.context,
      solutionAbsoluteURL: this.solutionAbsoluteURL,
      entityListId: this.entityWebProperties.WP_SITE_PARENT_LIST,
      entityListItemId: this.entityWebProperties.WP_SITE_PARENT_LIST_ITEM,
      targetWebUrl: this.targetWebUrl,
      title: this.properties.title
    });

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    console.log(this.finder + 'onPropertyPaneConfigurationStart');
    var tabIndex = this.tabProperties.length - 1;

    if (this.solutionAbsoluteURL === this.context.pageContext.web.absoluteUrl) {
      this.isRootWeb = true;
      this.tabProperties[tabIndex].scope = 'rootSite';
    } else {
      this.isRootWeb = false;
    }
    if (this.lists) {
      return;
    }

    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');

    try {
      const listOptions: IDropdownOption[] = await this.loadListOptions();
      this.lists = listOptions;

      if (this.tabProperties[tabIndex].list) {
        const fieldOptions: IDropdownOption[] = await this.loadTargetListFilteringFieldOptions(this.tabProperties.length);
        this.fieldOptions = fieldOptions;
        const viewOptions: IDropdownOption[] = await this.loadViewOptions(this.tabProperties.length); // check dem osv andre
        this.tabViewOptions[tabIndex].items = viewOptions;
      }
      if (!this.isRootWeb) {
        const entityFieldOptions: IDropdownOption[] = await this.loadEntityFieldOptions();
        this.entityFieldOptions = entityFieldOptions;
      }
    } catch (exception) { 
      console.log(exception);
    }
    this.context.propertyPane.refresh();
    this.context.statusRenderer.clearLoadingIndicator(this.domElement);
    this.render();
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    console.log(this.finder + 'onPropertyPaneFieldChanged');
    var tabIndex: number = 0;

    if (propertyPath.indexOf('scope') !== -1) {
      tabIndex = Number(propertyPath.substring(5));

      this.getTargetWebURL(tabIndex).then((targetWebUrl: string) => {
        this.targetWebUrl = targetWebUrl;
        this.targetWeb = new Web(targetWebUrl);

        super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

        const previousList: string = this.tabProperties[tabIndex].list;

        this.lists = [];
        this.tabViewOptions[tabIndex].items = [];

        this.tabProperties[tabIndex].list = undefined;
        this.tabProperties[tabIndex].view = undefined;

        this.onPropertyPaneFieldChanged('list' + tabIndex, previousList, this.tabProperties[tabIndex].list);

        this.context.propertyPane.refresh();

        this.loadListOptions().then((listOptions: IDropdownOption[]) => {
          this.lists = listOptions; // her
          const prevList: string = this.tabProperties[tabIndex].list;
          this.tabProperties[tabIndex].list = listOptions[4].key as string;

          console.log('ØØØØØØØØØØ');
          console.log(listOptions[tabIndex].text);
          console.log(listOptions[tabIndex].key);
          console.log(this.tabProperties[tabIndex].list);

          this.tabProperties[tabIndex].listName = listOptions[tabIndex].text;
          this.onPropertyPaneFieldChanged('list' + tabIndex, prevList, this.tabProperties[tabIndex].list);

          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          this.render();
          this.context.propertyPane.refresh();
        });
      });
      return;
    }
    if (propertyPath.indexOf('list') !== -1 && newValue) {
      tabIndex = Number(propertyPath.substring(4));

      if (this.propertyPaneGroups[tabIndex].groupName) {
        if (this.propertyPaneGroups[tabIndex].groupName.toString().indexOf('Tab') !== -1) {
          this.tabProperties[tabIndex].list = newValue;
          this.tabProperties[tabIndex].title = this.propertyPaneGroups[tabIndex].groupName;
          this.tabProperties[tabIndex].listName = newValue;
          console.log(newValue);
        }
      }
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      const previousView: string = this.tabProperties[tabIndex].view;
      this.tabProperties[tabIndex].view = undefined;
      this.tabViewOptions[tabIndex].items = [];

      this.onPropertyPaneFieldChanged('view' + tabIndex, previousView, this.tabProperties[tabIndex].view);

      this.loadViewOptions(tabIndex).then((viewOptions: IDropdownOption[]) => {
        try {
          this.tabViewOptions[tabIndex].items = viewOptions;
          this.tabProperties[tabIndex].view = viewOptions[0].key as string;
          this.onPropertyPaneFieldChanged('view' + tabIndex, undefined, this.tabProperties[tabIndex].view);
          this.context.propertyPane.refresh();
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          this.render();
          this.context.propertyPane.refresh();
        } catch (exception) {
          console.log(exception);
        }
      });
      return;
    }
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    console.log(this.finder + 'getPropertyPaneConfiguration');

    let titleGroup = [];
    let buttonsGroup = [];

    titleGroup.push(PropertyPaneTextField('title', {
      label: strings.Title,
    }));

    if (!this.isRootWeb) {
      this.scopeOptions.push({
        key: 'currentSite',
        text: 'Current site'
      });
      try {
        const entityItemLocation = this.entityWebProperties.WP_ITEM_LOCATION;
        let wpItemLocationStringParts: number = entityItemLocation.split(";").filter(part => part !== "").length;
        let numberOfEntitiesInCurrentHierarchy: number = (wpItemLocationStringParts - 1) / 2;
        if (numberOfEntitiesInCurrentHierarchy > 1) {
          this.scopeOptions.push({
            key: 'parentSite',
            text: 'Parent site'
          });
        }
      } catch (exception) {
        console.warn(`List view web part configruation pane could not determine if this entity has a parent.`);
      }
    }

    this.propertyPaneGroups = [];

    this.propertyPaneGroups.push({
      groupFields: titleGroup
    });

    var tabControls;
    for (var i = 1; i <= this.tabProperties.length - 1; i++) {
      let tabNumber = i;
      tabControls = [];
      if (this.tabViewOptions[tabNumber]) {
        tabControls.push(PropertyPaneDropdown('scope' + tabNumber, {
          label: strings.Scope,
          options: this.scopeOptions
        }));
        tabControls.push(PropertyPaneDropdown('list' + tabNumber, {
          label: strings.SelectList,
          options: this.lists
        }));
        tabControls.push(PropertyPaneDropdown('view' + tabNumber, {
          label: strings.SelectView,
          options: this.tabViewOptions[tabNumber].items
        }));

        this.propertyPaneGroups.push({
          groupName: 'Tab ' + tabNumber,
          groupFields: tabControls
        });
      }
    }

    buttonsGroup.push(PropertyPaneButton('deleteTabButton', {
      text: 'Delete Tab ',
      buttonType: PropertyPaneButtonType.Normal,
      onClick: this.buttonDeleteTab.bind(this),
    }));
    buttonsGroup.push(PropertyPaneButton('addTabButton', {
      text: 'Add new Tab',
      buttonType: PropertyPaneButtonType.Primary,
      onClick: this.buttonAddTab.bind(this),
    }));

    this.propertyPaneGroups.push({
      groupFields: buttonsGroup
    });

    return {
      pages: [
        {
          header: {
            description: strings.ListViewPropertyPaneDescription
          },
          groups: this.propertyPaneGroups,
          displayGroupsAsAccordion: false
        }
      ]
    };
  }

  private buttonAddTab(): any {
    if (this.tabProperties.length != 7) {
      let tabIndex = this.tabProperties.length;
      this.tabProperties.push({
        scope: 'rootSite',
        list: '',
        title: 'Tab ' + tabIndex,
        view: '',
        listName: ''
      });
      this.tabViewOptions.push({
        items: []
      });

      this.loadListOptions().then((listOptions: IDropdownOption[]) => {
        this.lists = listOptions;
        this.tabProperties[tabIndex].list = listOptions[0].key as string;
        this.tabProperties[tabIndex].listName = listOptions[0].text;
      });
    }
  }

  private buttonDeleteTab(): any {
    if (this.tabProperties.length == 1 || this.tabProperties.length == 2) {
      return;
    } else {
      this.tabProperties.pop();
      this.tabViewOptions.pop();
    }
  }

  private loadListOptions = async (): Promise<IDropdownOption[]> => {
    let dropDownOptions: IDropdownOption[] = [];
    try {
      const web = this.targetWeb;
      if (!web) {
        throw "No web available";
      }
      const lists: BasicListObject[] = await this.getLists(web);
      dropDownOptions = lists.map(list => ({ key: list.Id, text: list.Title }));
    } catch (exception) {
      console.warn(`WorkPoint list view webpart loading lists exception: ${exception}`);
    }
    return Promise.resolve(dropDownOptions);
  }

  private loadTargetListFilteringFieldOptions = async (tabIndex: number): Promise<IDropdownOption[]> => {
    let fieldOptions: IDropdownOption[] = null;
    try {
      if (!this.tabProperties[tabIndex].list) {
        throw "No list has been selected";
      }
      const web = this.targetWeb;
      const fields: ListViewFieldObject[] = await this.getFields(web, tabIndex);
      fieldOptions = fields.map(field => {
        if (!field.ReadOnlyField && field.Filterable && field.Sortable) {
          return { key: field.StaticName, text: field.Title };
        } else {
          return null;
        }
      }).filter(field => field !== null);
    } catch (exception) {
      console.warn(`WorkPoint list view webpart loading fields exception: ${exception}`);
      fieldOptions = [];
    }
    return Promise.resolve(fieldOptions);
  }

  private loadEntityFieldOptions = async (): Promise<IDropdownOption[]> => {
    let fieldOptions: IDropdownOption[] = null;
    try {
      const solutionWeb: Web = new Web(this.solutionAbsoluteURL);
      const entityFields: ListViewFieldObject[] = await this.getEntityFields(solutionWeb, this.entityWebProperties.WP_SITE_PARENT_LIST);
      fieldOptions = entityFields.map(entityField => ({ key: entityField.StaticName, text: entityField.Title }));
    } catch (exception) {
      console.warn(`WorkPoint list view webpart loading entity fields exception: ${exception}`);
      fieldOptions = [];
    }
    return fieldOptions;
  }

  private loadViewOptions = async (tabIndex: number): Promise<IDropdownOption[]> => {
    let viewOptions: IDropdownOption[] = null;
    try {
      if (!this.tabProperties[tabIndex].list) {
        throw "No list has been selected";
      }
      const web = this.targetWeb;
      const views: BasicViewObject[] = await this.getViews(web, tabIndex);
      viewOptions = views.map(view => ({ key: view.Id, text: view.Title }));
    } catch (exception) {
      console.warn(`WorkPoint list view webpart loading views exception: ${exception}`);
      viewOptions = [];
    }
    return Promise.resolve(viewOptions);
  }

  private getLists = async (web: Web): Promise<BasicListObject[]> => {
    return web.lists.select("Title,Id").get();
  }

  private getFields = async (web: Web, tabIndex: number): Promise<ListViewFieldObject[]> => {
    return web.lists.getById(this.tabProperties[tabIndex].list).fields.select(`Title, InternalName, EntityPropertyName, TypeAsString, FieldTypeKind, ReadOnlyField, Filterable, Sortable, StaticName`).filter('Hidden eq false').get();
  }

  private getViews = async (web: Web, tabIndex: number): Promise<BasicViewObject[]> => {
    return web.lists.getById(this.tabProperties[tabIndex].list).views.select("Id, Title, ServerRelativeUrl, Hidden").filter(`Hidden eq false`).get();
  }

  private getEntityFields = async (web: Web, listId: string): Promise<ListViewFieldObject[]> => {
    return web.lists.getById(listId).fields.select(`Title, InternalName, EntityPropertyName, TypeAsString, FieldTypeKind, ReadOnlyField, Filterable, Sortable, StaticName`).filter('Hidden eq false').get();
  }

  private getTargetWebURL = async (tabIndex: number): Promise<string> => {
    switch (this.tabProperties[tabIndex].scope) {
      case "currentSite":
        return this.context.pageContext.web.absoluteUrl;
      case "parentSite":
        const parentRelativeWebUrl: string = await this.fetchEntityParentWebRelativeURL();
        return getWebAbsoluteUrl(this.solutionAbsoluteURL, parentRelativeWebUrl);
      case "rootSite":
      default:
        return this.solutionAbsoluteURL;
    }
  }

  private fetchEntityParentWebRelativeURL = async (): Promise<string> => {
    let webUrl = this.context.pageContext.web.absoluteUrl;
    const entityItemLocation = this.entityWebProperties.WP_ITEM_LOCATION;
    let wpItemLocationStringParts: string[] = entityItemLocation.split(";").filter(part => part !== "");
    try {
      wpItemLocationStringParts.pop();
      wpItemLocationStringParts.pop();
      const parentEntityListItemId: number = parseInt(wpItemLocationStringParts.pop());
      const parentEntityListId: string = wpItemLocationStringParts.pop();
      const parentListItem = await DataService.loadEntityInformation(this.solutionAbsoluteURL, parentEntityListItemId, parentEntityListId, [BusinessModuleFields.Site]);
      webUrl = parentListItem.wpSite;
    } catch (exception) {
      console.warn("List view webpart could not load the entity parent web, due to invalid 'wpItemLocation' value. Defaulting to current entity value.");
    }
    return webUrl;
  }

  private needsConfiguration = (tabIndex: number): boolean => {
    console.log(this.finder + 'needsConfiguration');
    var tabIndex = tabIndex - 1;
    return this.tabProperties[tabIndex].list === null ||
      this.tabProperties[tabIndex].list === undefined ||
      this.tabProperties[tabIndex].list.trim().length === 0 ||
      this.tabProperties[tabIndex].view === null ||
      this.tabProperties[tabIndex].view === undefined ||
      this.tabProperties[tabIndex].view.trim().length === 0;
  }
}
