export interface IFieldMap {
  internalName: string;
  type: IFieldMappingType;
}

export enum IFieldMappingType {
  "Address",
  "City",
  "ZipCode",
  "Country",
  "Phone",
  "Email",
  "Url"
}

export interface IFieldMappingsSettings {

  /**
   * Not implemented:
   * ActiveSettings
   * CommonFieldMappingsSettings
   * MyEntitiesFieldNames
   * CurrentUserMappingField
   */
  
  AddressField: string;
  CityField: string;
  ZipcodeField: string;
  CountryField: string;
  PhoneFieldNames: string[];
  EmailFieldNames: string[];
  UrlFieldNames: string[];
}

export function getFieldMappingType(fieldName: string, fieldMap:IFieldMap[]):IFieldMappingType {

  for (let i = 0; i < fieldMap.length; i++) {
    const iteration = fieldMap[i];

    if (fieldName === iteration.internalName) {
      return iteration.type;
    }
  }

  return null;
}

export function getFieldMaps(fieldMappingsSettings:IFieldMappingsSettings):IFieldMap[] {
  let fieldMaps: IFieldMap[] = [];

  // Singular fields
  if (typeof fieldMappingsSettings.AddressField === "string" && fieldMappingsSettings.AddressField !== "") {
    fieldMaps.push({
      internalName: fieldMappingsSettings.AddressField,
      type: IFieldMappingType.Address
    });
  }

  if (typeof fieldMappingsSettings.CityField === "string" && fieldMappingsSettings.CityField !== "") {
    fieldMaps.push({
      internalName: fieldMappingsSettings.CityField,
      type: IFieldMappingType.City
    });
  }

  if (typeof fieldMappingsSettings.ZipcodeField === "string" && fieldMappingsSettings.ZipcodeField !== "") {
    fieldMaps.push({
      internalName: fieldMappingsSettings.ZipcodeField,
      type: IFieldMappingType.ZipCode
    });
  }

  if (typeof fieldMappingsSettings.CountryField === "string" && fieldMappingsSettings.CountryField !== "") {
    fieldMaps.push({
      internalName: fieldMappingsSettings.CountryField,
      type: IFieldMappingType.Country
    });
  }

  // Multiple fields
  if (Array.isArray(fieldMappingsSettings.PhoneFieldNames) && fieldMappingsSettings.PhoneFieldNames.length > 0) {
    fieldMappingsSettings.PhoneFieldNames.forEach(internalFieldName => {

      if (typeof internalFieldName === "string" && internalFieldName !== "") {
        fieldMaps.push({
          internalName: internalFieldName,
          type: IFieldMappingType.Phone
        });
      }
    });
  }

  if (Array.isArray(fieldMappingsSettings.EmailFieldNames) && fieldMappingsSettings.EmailFieldNames.length > 0) {
    fieldMappingsSettings.EmailFieldNames.forEach(internalFieldName => {

      if (typeof internalFieldName === "string" && internalFieldName !== "") {
        fieldMaps.push({
          internalName: internalFieldName,
          type: IFieldMappingType.Email
        });
      }
    });
  }
  
  if (Array.isArray(fieldMappingsSettings.UrlFieldNames) && fieldMappingsSettings.UrlFieldNames.length > 0) {
    fieldMappingsSettings.UrlFieldNames.forEach(internalFieldName => {

      if (typeof internalFieldName === "string" && internalFieldName !== "") {
        fieldMaps.push({
          internalName: internalFieldName,
          type: IFieldMappingType.Url
        });
      }
    });
  }

  return fieldMaps;
}