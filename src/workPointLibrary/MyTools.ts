import { getNumberLocaleFromStringLocale } from "./Locales";

export interface IMyToolsSettings {
  BusinessModuleId: string;
  Groups: IMyToolsGroupSetting[];
}

export interface IMyToolsGroupSetting {
  GroupTitle: string;
  TitleResource: object;
  Buttons: IMyToolsButtonSetting[];
}

export interface IMyToolsButtonSetting {
  Title: string; // "Email actionflow"
  TitleResource: object; // { "1030": "E-mail actionflow", "1033": "Email actionflow" }
  Icon: string; // "https://workpoint.azureedge.net/images/icons/color/Doc_32.png"
  Id?: string; // "0886eefc-6f0c-4ee0-b134-61eaf1794f47"
  Type: MyToolsButtonType; // WizardType eg: NewSiteListItem
  Wizard: string; // "DocumentProvisioning" in old Wizards and "1d6ad226-ac83-473b-be01-cb8978608236" in new Wizards 
}

export interface IWizardArgument<ValueType> {
  Value: ValueType;
  IsReadOnly: boolean;
}

export interface IMyToolsAdvancedWizardButtonSetting extends IMyToolsButtonSetting {
  BMAId?: string; // Optional business module id (A for for future support of 2 pre-selected business modules, eg. "BMBId")
  ItemAId?: string;
  FilterFieldA?: string;
  FilterValueA?: string;
  
  /* These values exist, but are not used in front-end
  TemplateLanguage?: IWizardArgument<string>;
  TemplateSet?: IWizardArgument<string>;
  DocumentLibrary?: IWizardArgument<string>;
  Folder?: IWizardArgument<string>;
  SaveType?: IWizardArgument<string>;
  DocumentSetContentType?: IWizardArgument<string>;
  FileName?: IWizardArgument<string>;
  */
}

export interface IMyToolsCustomScriptButtonSetting extends IMyToolsButtonSetting {
  Code: string; // Custom javascript
}

export interface IMyToolsItemCreationButtonSetting extends IMyToolsButtonSetting {
  TriggerId: string; // string GUID
}

export interface IMyToolsNewListItemButtonSetting extends IMyToolsButtonSetting {
  ListUrl: string;
  ContentType?: string; // "0x0100D7ACB9D6CFA0E94B952EA746ED6BF621"
  Upload?: boolean; // Used?
}

export interface IMyToolsNewParentListItemButtonSetting extends IMyToolsNewListItemButtonSetting {
  ParentBmId: string; // "0886eefc-6f0c-4ee0-b134-61eaf1794f47"
}

export interface IMyToolsRelationWizardButtonSetting extends IMyToolsButtonSetting {
  RelationTypeA?: number; // 3
  RelationTypeB?: number; // 25
  BusinessModule?: string; // "0886eefc-6f0c-4ee0-b134-61eaf1794f47"
  FilterField?: string;
  FilterValue?: string;
}

export interface IMyToolsLinkButtonSetting extends IMyToolsButtonSetting {
  Url: string;
  Target: "_blank"|"_self"|"dialog";
}

export interface IMyToolsCurrentItemActionButtonSetting extends IMyToolsButtonSetting {
  ActionType: MyToolsCurrentItemActionType;
}

export enum MyToolsCurrentItemActionType {
  Edit = "Edit",
  Delete = "Delete",
  ChangeStage = "ChangeStage",
  View = "View"
}

export enum MyToolsButtonType {
  OpenWizard = "OpenWizard",
  NewBusinessModuleEntity = "NewBusinessModuleEntity",
  Link = "Link",
  CurrentItemAction = "CurrentItemAction",
  NewSiteListItem = "NewSiteListItem",
  NewRootListItem = "NewRootListItem",
  NewParentListItem = "NewParentListItem",
  AddRelation = "AddRelation",
  CustomScript = "CustomScript",
  ItemCreationTrigger = "ItemCreationTrigger",
  Favorite = "Favorite",
}

export class MyToolsSettingsCollection {

  public settings:IMyToolsSettings[] = null;

  /**
   * Builds the MyTools part of the WorkPoint settings.
   * 
   * Localizes Title resources if supplied with a culture name.
   * 
   * @param myToolsSettingsJSONString JSON string representing the MyTools settings.
   * @param currentUICultureName Optional locale to use for overriding Title resources.
   */
  constructor(myToolsSettingsJSONString:string, currentUICultureName?:string) {

    const myToolsSettings:IMyToolsSettings[] = JSON.parse(myToolsSettingsJSONString);
      
    const desiredLocale: number = getNumberLocaleFromStringLocale(currentUICultureName);

    if (desiredLocale) {
      myToolsSettings.forEach(moduleSetting => {

        // Iterate Groups
        moduleSetting.Groups.forEach(group => {

          if (group.TitleResource) {
            const desiredGroupTitleTranslation:string = group.TitleResource[desiredLocale];
  
            if (desiredGroupTitleTranslation && desiredGroupTitleTranslation !== "") {
              group.GroupTitle = desiredGroupTitleTranslation;
            }
          }

          // Iterate Buttons
          group.Buttons.forEach(button => {

            if (button.TitleResource) {
              const desiredButtonTitleTranslation:string = button.TitleResource[desiredLocale];
  
              if (desiredButtonTitleTranslation && desiredButtonTitleTranslation !== "") {
                button.Title = desiredButtonTitleTranslation;
              }
            }
          });

        });
      });
    }

    this.settings = myToolsSettings;
  }

  public getSettingsForSolution(): IMyToolsSettings {
    const myToolsSettings:IMyToolsSettings = this.settings.filter(setting => {
      return setting.BusinessModuleId === null;
    })[0];

    return myToolsSettings;
  }

  public getSettingsForBusinessModule(businessModuleId:string): IMyToolsSettings {
    const myToolsSettings:IMyToolsSettings = this.settings.filter(setting => {
      return businessModuleId === setting.BusinessModuleId;
    })[0];

    return myToolsSettings;
  }
}