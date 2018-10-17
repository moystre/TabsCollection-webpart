import { IDropdownOption } from "office-ui-fabric-react";

export interface ITabSettings {
    title: string;
    scope: "rootSite" | "parentSite" | "currentSite";
    list: string;
    view: string;
    listName: string;
}

export interface ITabsCollectionProps {
    tabsArray: ITabSettings[];
    webpartSettings: ITabSettings;
    solutionAbsoluteURL: string;
    needsConfiguration: boolean;
    context: any;
    entityListId: string;
    entityListItemId: number;
    targetWebUrl: string;
    title: string
}

export interface ITabsCollectionState {
    selectedTab: number;
}

export interface ITabState {
    listBaseTemplate: any;
    listRelativeUrl: string;
    viewRelativeUrl: string;
    currentStageKey: string;
    items: any[];
    currentFolderPath: string;
    parentFolders: any[];
    columns: any[];
    paginationString: string;
    paginateBackwardString: string;
    paginateForwardString: string;
    webpartIsLoading: boolean;
    fetchingListItems: boolean;
    rowLimit: number;
    viewFieldOrders: any[];
    viewFields: any[];
    viewXml: string;
}

export interface ITabOptions {
    items: IDropdownOption[];
}

export interface IManagedMetadataShallowObject {
    Label: string;
    TermGuid: string;
}

export interface ILookupShallowObject {
    email?: string;
    picture?: string;
    id: number;
    title: string;
    isSecretFieldValue?: boolean;
    lookupId?: number;
    lookupValue?: string;
}

export interface ITabsTestProps {
    index: number;
}
