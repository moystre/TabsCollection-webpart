import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { IBaseProps } from 'office-ui-fabric-react/lib/Utilities';
import { IBusinessModuleEntity } from '../workPointLibrary/BusinessModule';
import { IUserLicense } from '../workPointLibrary/License';

export interface IWorkPointContext {
  solutionAbsoluteUrl: string;
  solutionRelativeUrl: string;
  appLaunchUrl: string;
  appWebFullUrl: string;
  sharePointContext: ApplicationCustomizerContext;
  userLicense: IUserLicense;
}

export interface IWorkPointBaseProps extends IBaseProps {
  context: IWorkPointContext;
}

export enum FoldoutType {
  listViewCollection,
  listItemCollection,
  listView,
  listItem,
  businessModuleViewCollection,
  businessModuleItemCollection,
  businessModuleView
}

export interface IFoldoutItemData {
  title: string;
  text: string;
  iconClass?: string;
  iconUrl?: string;
  menuItems?: IFoldoutItemData[];
  type: FoldoutType;
  context: IWorkPointContext;
  entity: IBusinessModuleEntity;
  list?: IListData;
  url?: string;
}

export interface IListData {
  id: string;
  parentWebUrl: string;
  defaultViewUrl: string;
  baseTemplate: number;
}

export interface IFoldoutMenuData {
  opened: boolean;
  menuItems: IFoldoutItemData[];
  context: ApplicationCustomizerContext;
}

export interface IFieldAndValueObject {
  field: string;
  value: string;
}