import { BasePermissions, Field, Item, List } from 'sp-pnp-js';
import { ContentType } from 'sp-pnp-js/lib/sharepoint/contenttypes';
import { View } from 'sp-pnp-js/lib/sharepoint/views';

export class BasicListObject extends List {
  public Id: string;
  public Title: string;
}

export class ListObject extends BasicListObject {
  public EffectiveBasePermissions: BasePermissions;
  public ContentTypes: WorkPointContentType[];
  public BaseTemplate: number;
  public DefaultViewUrl: string;
  public DefaultNewFormUrl: string;
  public DefaultEditFormUrl: string;

  public ParentWebUrl: string;
}

export class ListItemObject extends Item {
  public Id: number;
  public Title: string;
}

export class SitePageListObject extends Item {
  public Title:string;
  public FieldValuesAsText: SitePageListObjectFieldValues;
}

export class BasicViewObject extends View {
  public Title: string;
  public Id: string;
  public ServerRelativeUrl: string;
  public Hidden: boolean;
  public RowLimit: number;
  public ListViewXml: string;
}

export interface SitePageListObjectFieldValues {
  FileLeafRef: string;
}

export class WorkPointContentType extends ContentType {
  public StringId: string;
  public Name: string;
  public DocumentTemplateUrl: string;
}

export class BasicFieldObject extends Field {
  public Title: string;
  public InternalName: string;
  public EntityPropertyName: string;
  public TypeAsString: string;
  public FieldTypeKind: number;
}

export class ListViewFieldObject extends BasicFieldObject {
  public ReadOnlyField: boolean;
  public Filterable: boolean;
  public Sortable: boolean;
  public CanBeDeleted: boolean;
  public StaticName: string;
}