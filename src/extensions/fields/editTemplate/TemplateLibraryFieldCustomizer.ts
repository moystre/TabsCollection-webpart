import { BaseFieldCustomizer, IFieldCustomizerCellEventParameters } from '@microsoft/sp-listview-extensibility';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import TemplateEditField, { ITemplateEditFieldProps } from '../../../components/TemplateEditField';
import { getClientProgramByExtension } from '../../../workPointLibrary/Helper';
import { getUserLicenseStatus, IUserLicense, UserLicenseStatus } from '../../../workPointLibrary/License';
import { getTemplateLibraryMappings, ITemplateLibrarySettings } from '../../../workPointLibrary/service';

export default class TemplateLibraryFieldCustomizersFieldCustomizer
  extends BaseFieldCustomizer<null> {

  protected TemplateLibraryMappings:ITemplateLibrarySettings = null;

  public async onInit(): Promise<void> {
    const userLicenseStatus:IUserLicense = await getUserLicenseStatus(this.context.pageContext.site.absoluteUrl, this.context.serviceScope);

    if (userLicenseStatus && userLicenseStatus.Status === UserLicenseStatus.None) {
      return Promise.reject("User is not licensed");
    }

    this.TemplateLibraryMappings = await getTemplateLibraryMappings(this.context.pageContext.site.absoluteUrl, this.context.pageContext.web.serverRelativeUrl);

    return Promise.resolve(null);
  }

  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {

    let name:string = event.fieldValue;

    // Handle empty title value
    if (!name || name === "" || name === "0") {
      name = event.listItem.getValueByName("FileName");
    }

    const fileName:string = event.listItem.getValueByName("FileLeafRef");
    const extensionStart:number = fileName.lastIndexOf(".");
    const extension:string = fileName.slice(extensionStart + 1);
    const documentType = getClientProgramByExtension(extension);
    const fileUrl:string = event.listItem.getValueByName("FileRef");
    const id:string = event.listItem.getValueByName("ID");
    const contentTypeId:string = event.listItem.getValueByName("ContentTypeId");

    const templateEditField: React.ReactElement<{}> = React.createElement(TemplateEditField, {
      name,
      templateLibrarySettings: this.TemplateLibraryMappings,
      fileName,
      fileType: documentType,
      contentTypeId,
      id,
      solutionAbsoluteUrl: this.context.pageContext.site.absoluteUrl,
      fileUrl
    } as ITemplateEditFieldProps);

    ReactDOM.render(templateEditField, event.domElement);
  }

  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    ReactDOM.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }
}
