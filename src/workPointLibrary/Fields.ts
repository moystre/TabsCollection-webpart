// Field keys
export const BusinessModuleFields = {
  Id: "Id",
  Title: "Title",
  ParentId: "wpParentId",
  Site: "wpSite",
  EffectiveBasePermissions: "EffectiveBasePermissions",
  StageDeadline: "wp_stageDeadline",
  StageHistory: "wp_stageHistory",
  ContentTypeId: "ContentTypeId",
  UniqueId: "UniqueId",
  ItemLocation: "wpItemLocation"
};

/**
 * Default list properties
 */
export const DefaultListFields = [
  "Title",
  "BaseTemplate",
  "Id",
  "ParentWebUrl",
  "DefaultNewFormUrl",
  "DefaultViewUrl",
  "EffectiveBasePermissions"
];

/** 
 * Default business module list select fields
 */
export const BusinessModuleSelectFields = [
  "Title",
  "BaseTemplate",
  "Id",
  "ParentWebUrl",
  "DefaultNewFormUrl",
  "DefaultViewUrl",
  "DefaultEditFormUrl",
  "EffectiveBasePermissions",
  "ContentTypes/Name",
  "ContentTypes/StringId"
];

/**
 * Fields on the Relation list
 */
export const RelationListFields = [
  "Id",
  "wpRelationATitle",
  "wpRelationATypeId",
  "wpRelationAItemId",
  "wpRelationAListId",
  "wpRelationAType/Title",
  "wpRelationAType/ID",
  "wpRelationBTitle",
  "wpRelationBTypeId",
  "wpRelationBItemId",
  "wpRelationBListId",
  "wpRelationBType/Title",
  "wpRelationBType/ID",
  "wpRelationDescription",
  "wpRelationResponsible/ID",
  "wpRelationResponsible/Title",
  "wpRelationStart",
  "wpRelationEnd"
];