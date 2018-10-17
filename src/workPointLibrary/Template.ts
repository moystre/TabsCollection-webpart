/**
 * Gets part of the wpword URL scheme.
 * Convenience method for common wpword URL parameters.
 * 
 * @private
 * 
 * @param rowLimit Unknown what this does.
 * @param businessModuleId Id of the business module to create a template for.
 * @param solutionAbsoluteUrl Absolute URL of the WorkPoint 365 solution. 
 * @param templateAbsoluteUrl Abosolute URL for the content type template to use for a new document template.
 */
const getCommonUrlParts = (rowLimit:number, businessModuleId:string, solutionAbsoluteUrl:string, templateAbsoluteUrl:string):string => {
  return `wpword: -RowLimit=${rowLimit} -ListID=${businessModuleId} -Url="${solutionAbsoluteUrl}" -SaveLocation="${solutionAbsoluteUrl}/Template%20Library/" -Template="${encodeURI(templateAbsoluteUrl)}" -IsWorkPoint365=1`;
};

/**
 * Gets a WorkPoint Express 'wpword' template creation URL, using provided arguments.
 * 
 * @param rowLimit Unknown what this does.
 * @param businessModuleId Id of the business module to create a template for.
 * @param solutionAbsoluteUrl Absolute URL of the WorkPoint 365 solution. 
 * @param templateAbsoluteUrl Abosolute URL for the content type template to use for a new document template.
 */
export const getCreateWordUrl = (rowLimit:number = 10, businessModuleId:string, solutionAbsoluteUrl:string, templateAbsoluteUrl:string):string => {
  return `${getCommonUrlParts(rowLimit, businessModuleId, solutionAbsoluteUrl, templateAbsoluteUrl)} -Type=create`;
};

/**
 * Gets a WorkPoint Express 'wpword' template edit URL, using provided arguments.
 * 
 * @param rowLimit Unknown what this does.
 * @param businessModuleId Id of the business module to create a template for.
 * @param solutionAbsoluteUrl Absolute URL of the WorkPoint 365 solution. 
 * @param templateAbsoluteUrl Abosolute URL for the content type template to use for a new document template.
 */
export const getEditWordUrl = (rowLimit:number = 10, businessModuleId:string, solutionAbsoluteUrl:string, templateAbsoluteUrl:string):string => {
  return `${getCommonUrlParts(rowLimit, businessModuleId, solutionAbsoluteUrl, templateAbsoluteUrl)} -Type=edit`;
};