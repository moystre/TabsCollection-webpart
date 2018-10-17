export interface IOneSidedRelationItem {
  Id: number;
  Title: string;
  EntityListItemId: number;
  EntityListId: string;
  EntityTitle: string;
  Description?: string;
  Responsible: string;
  Start?: string;
  End?: string;
  IsActive: boolean;
  Type: number;
}

export interface IRelationItem {
  Id: number;

  wpRelationAType: ILookupValue;
  wpRelationAListId: string;
  wpRelationAItemId: number;
  wpRelationATitle: string;

  wpRelationBType: ILookupValue;
  wpRelationBListId: string;
  wpRelationBItemId: number;
  wpRelationBTitle: string;

  wpRelationDescription: string;
  wpRelationResponsible: ILookupValue;
  wpRelationStart: string;
  wpRelationEnd: string;
}

export interface ILookupValue {
  ID: number;
  Title: string;
}

export interface IRelationTypeItem {
  ID: number;
  Title: string;
}