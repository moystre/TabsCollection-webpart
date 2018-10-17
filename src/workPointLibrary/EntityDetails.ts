import { IFieldMappingType } from "./FieldMappings";
import { ensureGuidString } from "./Helper";
import { WorkPointSettingsListItem } from "./Settings";

export interface IField {
  internalFieldName: string;
}

export interface IFieldValue extends IField {
  value: string;
  displayName: string;
  type: string;
  fieldMappingType: IFieldMappingType;
}

export interface IEntityDetailsSettings {
  id: string;
  fields: IField[];
}

export class EntityDetailsSettingsCollection {
  public settings: IEntityDetailsSettings[];

  constructor (allSettings: WorkPointSettingsListItem[]) {
    
    // Map settings values
    const allEntityDetailSettings = allSettings.map(settingsListItem => {

      const wpSettingsKeyParts: string[] = settingsListItem.wpSettingsKey.split("_");

      const id: string = ensureGuidString(wpSettingsKeyParts[wpSettingsKeyParts.length-1]);

      // Escape first " of wpSettingsValue
      const fields: IField[] = settingsListItem.wpSettingsValue
        .replace(/\"/g, "")
        .split(";").map(internalFieldName => ({
          internalFieldName
        })
      );

      return {
        id,
        fields
      };
    });

    this.settings = allEntityDetailSettings;
  }

  /**
   * Given a business module id, returns the matching entity detail settings, if any.
   * 
   * @param businessModuleId String Guid, representing the business module list id.
   */
  public getSettingsForBusinessModule(businessModuleId: string): IEntityDetailsSettings {

    try {
    
      const entityDetailSettingsCandidate: IEntityDetailsSettings[] = this.settings.filter(setting => {
        return businessModuleId === setting.id;
      });
    
      if (entityDetailSettingsCandidate.length < 1) {
        throw `No entity detail settings for business module with id: ${businessModuleId}.`;
      }
      
      return entityDetailSettingsCandidate[0];

    } catch (exception) {
      console.warn(exception);
      return null;
    }
  }
}