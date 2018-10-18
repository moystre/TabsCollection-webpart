import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as CamlBuilder from 'camljs';
import { ActionButton, CheckboxVisibility, ConstrainMode, DetailsList, DetailsListLayoutMode, DetailsRow, IColumn, Icon, IDetailsRowProps, Link, SelectionMode, Spinner, SpinnerSize } from 'office-ui-fabric-react';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import * as React from 'react';
import { storage, Web } from 'sp-pnp-js';
import * as strings from 'WorkPointStrings';
import SiteField from '../../../components/SiteField';
import { addExpression, addFields, buildQueryElement, replaceViewQueryElement } from '../../../workPointLibrary/CAMLHelper';
import { IBasicStageFilterMessage } from '../../../workPointLibrary/Event';
import { caseInsensitiveStringCompare, getFileIconFromExtension, getOfficeDocumentIconFromExtension, isOfficeDocument } from '../../../workPointLibrary/Helper';
import { BasicViewObject } from '../../../workPointLibrary/List';
import * as DataService from '../../../workPointLibrary/service';
import { WorkPointStorageKey } from '../../../workPointLibrary/Storage';
import { ILookupShallowObject, IManagedMetadataShallowObject, ITabsCollectionProps, ITabState } from './ITabsCollection';
import styles from './TabsCollection.module.scss';

const WP_NONE_GROUP_HEADER: string = `<${strings.None}>`;

export class Tab extends React.Component<ITabsCollectionProps, ITabState> {
  constructor(props: ITabsCollectionProps) {
    super(props);
    let entityStageFilterKey: string = null;
    // Stage filter is only available for entity lists of the current entity. Otherwise, totally irrellevant.
    if (props.webpartSettings.scope === "currentSite") {
      try {
        let solutionAbsoluteURL = props.solutionAbsoluteURL.toLowerCase();
        let key = `${WorkPointStorageKey.stageFilter}.${solutionAbsoluteURL}.${props.entityListId}.${props.entityListItemId}`;
        entityStageFilterKey = key;
      } catch (exception) { }
    }

    this.state = {
      listBaseTemplate: null,
      listRelativeUrl: null,
      viewRelativeUrl: null,
      currentStageKey: entityStageFilterKey,
      items: [],
      currentFolderPath: null,
      parentFolders: [],
      columns: [],
      paginationString: null,
      paginateBackwardString: null,
      paginateForwardString: null,
      webpartIsLoading: true,
      fetchingListItems: true,
      rowLimit: 30,
      viewFieldOrders: [],
      viewFields: null,
      viewXml: null,
    };
  }

  public componentDidMount(): void {
    if (!this.props.needsConfiguration) {
      this.configureViewFields();
    }
    window.addEventListener("message", this.handleStageFilterChange);
  }

  /**
   * Just rerun the search, as it always reads the newest stage filter value from the sessionStorage.
   * It will also work when the stage filter is removed, as the stage filter value in the sessionStorage is not present anymore.
   * The 'currentStageKey' should exist already, so no need to reaply.
   */
  private handleStageFilterChange = (event: MessageEvent): void => {
    // This is not dangerous, and therefore we do not whitelist domains posting.
    let postMessage = event.data as IBasicStageFilterMessage;
    // Stage filter is only available for entity lists of the current entity. Otherwise, totally irrellevant.
    if (postMessage.type === "stage" && this.props.webpartSettings.scope === "currentSite") {
      this.configureViewFields();
    }
  }

  public componentWillUnmount(): void {
    window.removeEventListener("message", this.handleStageFilterChange);
  }

  public componentWillReceiveProps(nextProps: ITabsCollectionProps): void {
    if (!nextProps.needsConfiguration) {
      this.setState({
        currentFolderPath: null,
        parentFolders: [],
      }, () => this.configureViewFields());
    }
  }

  private _onItemInvoked = (item: any): void => {
    if (item.FSObjType === "1") {
      this._goDownFolder(item);
    }
  }

  private _goUpFolder = (): void => {
    let parentFolders = this.state.parentFolders;
    let currentFolder = parentFolders.length > 0 ? parentFolders[parentFolders.length - 1] : null;
    if (parentFolders.length > 0) {
      parentFolders = parentFolders.slice(0, -1);
    }
    this.setState({
      currentFolderPath: currentFolder,
      parentFolders: parentFolders
    }, () => {
      this._getListItems(currentFolder);
    });
  }

  private _goDownFolder = (item: any): void => {
    let parentFolders = this.state.parentFolders;
    parentFolders.push(item.FileDirRef);
    this.setState({
      parentFolders: parentFolders,
      currentFolderPath: item.FileDirRef,
    });
    this._getListItems(item.FileRef);
  }

  public render(): React.ReactElement<ITabsCollectionProps> {
    const { needsConfiguration, webpartSettings } = this.props;
    const { webpartIsLoading, items, columns, parentFolders, fetchingListItems, viewRelativeUrl } = this.state;
    return (
      <div className={styles.tabsCollection}>
        <div className={styles.container}>
          {(needsConfiguration) ? (
            <div className={styles.noSettings}>{strings.NoSettings}</div>
          ) : (
              <div className={styles.whiteBackground}>
                {(webpartIsLoading) ? (<Spinner size={SpinnerSize.large} label={`${strings.Loading}...`} />) :
                  (<div className={styles.row}>
                    <div>
                      {(parentFolders.length > 0) &&
                        <ActionButton
                          iconProps={{ iconName: 'NavigateBack' }}
                          onClick={this._goUpFolder}
                          text={strings.GoUp} />}
                      {fetchingListItems ? <Spinner size={SpinnerSize.large} label={`${strings.Loading}...`} /> :
                        <DetailsList
                          items={items}
                          columns={columns}
                          selectionMode={SelectionMode.single}
                          checkboxVisibility={CheckboxVisibility.hidden}
                          setKey='flatView'
                          layoutMode={DetailsListLayoutMode.justified}
                          constrainMode={ConstrainMode.unconstrained}
                          isHeaderVisible={true}
                          onItemInvoked={this._onItemInvoked}
                          onRenderRow={this._onRenderRow} />}
                    </div>
                    {(items.length === 0 && !fetchingListItems) ? (
                      <div>{strings.NoItemsToShowHere}</div>)
                      : (<div>
                        {(this.state.paginateForwardString || this.state.paginateBackwardString) && (
                          <div className={styles.msPaging}>
                            {this.state.paginateBackwardString && <ActionButton iconProps={{ iconName: 'PageLeft' }} text={strings.GoToPreviousPage} onClick={() => this._onChangePageClick(false)} />}
                            {this.state.paginateForwardString && <ActionButton iconProps={{ iconName: 'PageRight' }} text={strings.GoToNextPage} onClick={() => this._onChangePageClick(true)} />}
                          </div>
                        )}
                      </div>
                      )}
                  </div>
                  )}
              </div>
            )}
        </div>
      </div>
    );
  }

  private _onChangePageClick = (forward: boolean): void => {
    const paginationString = forward ? this.state.paginateForwardString : this.state.paginateBackwardString;
    this.setState({
      paginateForwardString: null,
      paginateBackwardString: null,
      paginationString: paginationString,
      fetchingListItems: true
    }, () => {
      this._getListItems();
    });
  }

  private overrideViewXmlWithWorkPointFiltering = (viewXmlString: string): string => {
    const { webpartSettings } = this.props;
    const { currentStageKey, listBaseTemplate } = this.state;
    let outputViewXmlString: string = viewXmlString;
    let expressions: CamlBuilder.IExpression[] = [];

    /**
     * Stage filter is only available for entity lists of the current entity. Otherwise, totally irrellevant.
     * @description Maybe we could improve the list view web part, so that it shows in the GUI, that the page is being filtered by a stage filter value, and also let them remove it.
     */
    if (webpartSettings.scope === "currentSite") {
      let currentStage: string = storage.session.get(currentStageKey);
      // let currentStage: string = sessionStorage.getItem(currentStageKey);
      if (currentStage) {
        expressions = addExpression("SP.FieldText", "wp_tag", currentStage, expressions);
      }
    }
    let fieldsToAdd: string[] = [
      'FileDirRef',
      'FileLeafRef',
      'FileRef'
    ];
    if (listBaseTemplate === 101) {
      fieldsToAdd.push('FSObjType');
    }
    outputViewXmlString = addFields(fieldsToAdd, outputViewXmlString);
    // Query is required, so ensure that Query and Where elements exists
    let queryElementXmlString: string = buildQueryElement(outputViewXmlString, expressions);
    outputViewXmlString = replaceViewQueryElement(outputViewXmlString, queryElementXmlString);
    return outputViewXmlString;
  }

  private _fetchResults = async (url: string, requestBody: any): Promise<any> => {
    const response: SPHttpClientResponse = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, { body: JSON.stringify(requestBody) });
    const result = await response.json();
    return result;
  }

  // Fetches the list items using a CAML query.
  private _getListItems = async (folderRelativeUrl?: string): Promise<void> => {
    const { webpartSettings, targetWebUrl } = this.props;
    const { viewXml, viewFields, paginationString } = this.state;
    if (!viewXml) {
      console.warn("No valid viewXML for this list view. Cannot render list.");
    }
    let overrideViewXml: string = this.overrideViewXmlWithWorkPointFiltering(viewXml);
    let fetchListItemsURL: string = `${targetWebUrl}/_api/web/Lists(guid'${webpartSettings.list}')/RenderListDataAsStream${paginationString ? paginationString : ""}`;

    const requestBody = {
      parameters: {
        RenderOptions: 4103,
        ViewXml: overrideViewXml,
        AddRequiredFields: true,
        ...folderRelativeUrl && { FolderServerRelativeUrl: folderRelativeUrl }
      }
    };
    const result = await this._fetchResults(fetchListItemsURL, requestBody);
    let tempItems = this.mapItemValues(result.ListData.Row);

    // If grouping is not enabled, we do not need to call the buildGrouping method, and can set 'webpartIsLoading' to false instantly.
    // Preserve groups is called set so to not rebuild groups from a filtered result set. This would diminish the amount of groups when users sorts or filters columns or selects a new group.
    // let isFinishedLoading:boolean = !webpartSettings.grouping;
    //check
    this.setState({
      items: tempItems,
      paginateForwardString: result.ListData.NextHref,
      paginateBackwardString: result.ListData.PrevHref,
      paginationString: null,
      webpartIsLoading: false,
      ... { fetchingListItems: false },
      currentFolderPath: folderRelativeUrl,
      columns: this.buildResultColumns(viewFields),
    });
  }
  // Maps special fields, eg: SP.FieldLookup and SP.FieldUser.
  private mapItemValues = (items: any): any => {
    const { viewFields } = this.state;
    let tempItems = [];
    items.forEach(item => {
      let tempItem: any = item;
      viewFields.forEach(field => {
        if (field.Type === 'SP.FieldLookup' || field.Type === 'SP.FieldUser') {
          try {
            const lookupProspect: any = item[field.InternalName];
            let lookupValues: ILookupShallowObject[] = null;
            if (!lookupProspect) {
              throw `LookupValue not found for field: ${field.InternalName}`;
            }
            if (Array.isArray(lookupProspect) && lookupProspect.length > 0) {
              lookupValues = lookupProspect;
            } else {
              lookupValues = [lookupProspect];
            }
            const simpleTextRepresentation: string[] = lookupValues.map(value => {
              switch (field.Type) {
                case "SP.FieldUser":
                  return value.title;
                case "SP.FieldLookup":
                default:
                  return value.lookupValue;
              }
            });
            tempItem[field.InternalName] = simpleTextRepresentation.join(", ");
          } catch (exception) {
            tempItem[field.InternalName] = "";
          }
        } else if (field.Type === "SP.Taxonomy.TaxonomyField") {
          // Grouping field name is not needed when we add a Taxonomy field
          try {
            const metaDataProspect: any = item[field.InternalName];
            let managedMetaDataValues: IManagedMetadataShallowObject[] = null;
            if (!metaDataProspect) {
              throw `Managed meta data value not found for field: ${field.InternalName}`;
            }
            if (Array.isArray(metaDataProspect) && metaDataProspect.length > 0) {
              managedMetaDataValues = metaDataProspect;
            } else {
              managedMetaDataValues = [metaDataProspect];
            }
            const textValues: string[] = managedMetaDataValues.map(value => value.Label);
            tempItem[field.InternalName] = textValues.join("; ");
          } catch (exception) {
            tempItem[field.InternalName] = WP_NONE_GROUP_HEADER;
          }
        } else if (field.TypeAsString === "Boolean") {
          const visualValue: string = item[field.InternalName];
          tempItem[field.InternalName] = visualValue;
        } else {
          let visualValue: string = item[field.InternalName];
          tempItem[field.InternalName] = visualValue;
        }
      });
      tempItems.push(tempItem);
    });
    return tempItems;
  }

  private configureViewFields = async (): Promise<void> => {
    const { webpartSettings, targetWebUrl } = this.props;
    let tempFields = [];
    const web: Web = new Web(targetWebUrl);
    let entityItemForFiltering: any = null;
    if (!caseInsensitiveStringCompare(this.props.solutionAbsoluteURL, this.props.context.pageContext.web.absoluteUrl)) {
      entityItemForFiltering = await DataService.loadEntityInformation(this.props.solutionAbsoluteURL, this.props.entityListItemId, this.props.entityListId);
    }
    if (webpartSettings.list) {
    var list = await web.lists.getById(webpartSettings.list).select('BaseTemplate,RootFolder/ServerRelativeUrl').expand('RootFolder').get();
    }
    let view: BasicViewObject = await web.lists.getById(webpartSettings.list).getView(webpartSettings.view).get();
    let viewFields = await web.lists.getById(webpartSettings.list).getView(webpartSettings.view).fields.get();
    let fields = await web.lists.getById(webpartSettings.list).fields.get();

    viewFields.Items.forEach(internalField => {
      let field = fields.filter(f => {
        if (internalField === 'LinkTitle')
          return f.Title === 'Title';
        else
          return f.Title === internalField || f.InternalName === internalField;
      })[0];
      // If field was found
      if (field) {
        let tempField: any = {
          Title: field.Title,
          InternalName: field.InternalName,
          Sortable: field.Sortable,
          Filterable: field.Filterable,
          TypeAsString: field.TypeAsString,
          Type: field['odata.type'],
          IsLinkTitle: internalField === 'LinkTitle' || internalField === 'LinkFilename' ? true : false
        };
        if (field['odata.type'] === 'SP.FieldLookup') {
          tempField.LookupList = field.LookupList;
          tempField.LookupWebId = field.LookupWebId;
        }
        tempFields.push(tempField);
      }
    });

    this.setState({
      listBaseTemplate: list.BaseTemplate,
      listRelativeUrl: list.RootFolder.ServerRelativeUrl,
      viewRelativeUrl: view.ServerRelativeUrl,
      viewFieldOrders: viewFields.Items,
      viewFields: tempFields,
      viewXml: view.ListViewXml,
      rowLimit: view.RowLimit
    }, () => {
      this._getListItems(this.state.currentFolderPath);
    });
  }

  private buildResultColumns = (viewFields: any): IColumn[] => {
    const { webpartSettings } = this.props;
    const { viewRelativeUrl, listRelativeUrl } = this.state;
    const newColumns: IColumn[] = [];
    for (let i = 0; i < viewFields.length; i++) {
      let viewField = viewFields[i];
      let col: IColumn = {
        key: viewField.InternalName,
        fieldName: viewField.InternalName,
        name: viewField.Title,
        isResizable: true,
        minWidth: 90,
        maxWidth: 100,
      };
      if (viewField.InternalName === "wpSite") {
        col.onRender = (item: any) => {
          const wpSiteValue = item[viewField.InternalName];
          if (wpSiteValue) {
            const title: string = item["Title"];
            const id: number = item["ID"];
            const listId: string = webpartSettings.list;
            const siteField = <SiteField title={title} urlValue={wpSiteValue} listId={listId} itemId={id} solutionAbsoluteUrl={this.props.solutionAbsoluteURL} />;
            return (<span>{siteField}</span>);
          } else {
            return <span />;
          }
        };
      } else if (viewField.Type === 'SP.FieldDateTime') {
        col.onRender = (item: any) => {
          if (item[viewField.InternalName])
            return (<span>{item[viewField.InternalName]}</span>);
          else
            return <span />;
        };
      } else if (viewField.Type === 'SP.FieldLookup') {
        col.onRender = (item: any) => {
          return <span>{item[viewField.InternalName]}</span>;
        };
      } else if (viewField.Type === 'SP.FieldUser') {
        let userLink = this.props.context.pageContext.site.absoluteUrl + '/_layouts/15/userdisp.aspx?ID=';
        col.onRender = (item: any) => {
          return <span>{item[viewField.InternalName]}</span>;
        };
      } else if (viewField.Type === 'SP.FieldUrl') {
        col.onRender = (item: any) => {
          if (item[viewField.InternalName])
            return (<Link target='_blank' href={item[viewField.InternalName].Url}>{item[viewField.InternalName].Description}</Link>);
          else
            return <span />;
        };
      } else if (viewField.Type === 'SP.FieldMultiLineText') {
        col.isMultiline = true;
        col.onRender = (item: any) => {
          if (item[viewField.InternalName]) {
            let multiField = document.createElement('div');
            multiField.innerHTML = item[viewField.InternalName];
            return <span>{multiField.innerText}</span>;
          } else {
            return <span />;
          }
        };
      }

      if (viewField.IsLinkTitle) {
        col.minWidth = 200;
        col.maxWidth = 250;
        col.onRender = (item: any) => {
          let columnValue: any = item[viewField.InternalName];
          // Add backup to "FileLeafRef", when we cant fetch the requested field.
          if (!columnValue) {
            columnValue = item["FileLeafRef"];
          }
          if (columnValue) {
            let linkToItem = listRelativeUrl + '/DispForm.aspx?ID=' + item.ID;
            if (this.state.listBaseTemplate === 101) {
              if (item.FSObjType === "1") {
                return (<span>
                  <Link onClick={() => {
                    this._goDownFolder(item);
                  }}>{columnValue}</Link>
                </span>);
              }
              if (item.ServerRedirectedEmbedUrl)
                linkToItem = item.ServerRedirectedEmbedUrl;
              else {
                linkToItem = decodeURI(viewRelativeUrl + '?id=' + item.FileRef + '&parent=' + item.FileDirRef);
              }
            }
            return (<span><Link href={linkToItem} className='linkTitleColumn' target='_blank'>{columnValue}</Link>
              <span id={'wp-targetCallout-' + item.ID} />
            </span>);
          } else {
            return <span />;
          }
        };
      }

      if (viewField.InternalName === 'DocIcon') {
        col.isIconOnly = true;
        col.minWidth = 17,
          col.maxWidth = 17,
          col.onRender = (item: any) => {
            if (item.FSObjType === "1") {
              return <Icon iconName='FolderHorizontal' className='ms-font-l' />;
            } else {
              const isOfficeIcon: boolean = isOfficeDocument(item.File_x0020_Type);
              const icon: string = isOfficeIcon ? getOfficeDocumentIconFromExtension(item.File_x0020_Type) : getFileIconFromExtension(item.File_x0020_Type);
              if (isOfficeIcon) {
                return <div className={`ms-BrandIcon--icon16 ms-BrandIcon--${icon}`}></div>;
              } else {
                return <i className={`ms-Icon ms-Icon--${icon}`} aria-hidden="true"></i>;
              }
            }
          };
      }
      newColumns.push(col);
    }
    return newColumns;
  }

  private _onRenderRow = (props: IDetailsRowProps): JSX.Element => {
    return (
      <div className='wp-hoveringItem'>
        <DetailsRow
          {...props} />
      </div>
    );
  }
}

export class TabsCollection extends React.Component<ITabsCollectionProps, any> {
  constructor(props: ITabsCollectionProps) {
    super(props);
    this.state = {
      activeTab: 1
    };
  }

  public render(): React.ReactElement<ITabsCollectionProps> {
    const { tabsArray,targetWebUrl, solutionAbsoluteURL, needsConfiguration, entityListItemId, entityListId, context, title } = this.props;
    return (
      <div className={styles.tabsCollection}>
        <div className={styles.title}>
          <a className={styles.title}>{title}</a>
        </div>
        <br></br>
        <div className={styles.tabsRow}>
          {tabsArray.map(tab =>
            <div>
              {tab.title ?
                this.state.activeTab == this.getTabIndex(tab.title) ?
                <PrimaryButton
                className={styles.listTab}
                text={tab.listName? tab.listName : tab.title}
                onClick={(event: React.MouseEvent<HTMLButtonElement>) => {
                  this.selectTab(tab.title);
                }} />
                :
                <DefaultButton
                  className={styles.listTabSelected}
                  text={tab.listName? tab.listName : tab.title}
                  onClick={(event: React.MouseEvent<HTMLButtonElement>) => {
                    this.selectTab(tab.title);
                  }} /> 
                : null}
              &nbsp;
              </div>
          )}
        </div>
        <hr></hr>
        <Tab
          tabsArray={tabsArray}
          webpartSettings={tabsArray[this.state.activeTab]}
          solutionAbsoluteURL={solutionAbsoluteURL}
          needsConfiguration={needsConfiguration}
          context={context}
          entityListId={entityListId}
          entityListItemId={entityListItemId}
          targetWebUrl={targetWebUrl}
          title={title}>
        </Tab>

      </div>
    );
  }
  private async selectTab(tabTitle: string): Promise<void> {
    var tabIndex = Number(tabTitle.substring(4));
    this.setState({ activeTab: tabIndex });
  }

  private getTabIndex(tabTitle: string): number {
    return Number(tabTitle.substring(4));
  }
}

