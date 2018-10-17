import { BaseFieldCustomizer, FieldCustomizerContext, IFieldCustomizerCellEventParameters } from '@microsoft/sp-listview-extensibility';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { YouHaveNoWorkPoint365License } from 'WorkPointStrings';
import SiteField, { ISiteFieldProps } from '../../../components/SiteField';
import { getUserLicenseStatus, IUserLicense, UserLicenseStatus } from '../../../workPointLibrary/License';

export default class SiteFieldCustomizer
  extends BaseFieldCustomizer<null> {

  protected _userLicenseStatus:IUserLicense = null;

  public async onInit(): Promise<void> {
    this._userLicenseStatus = await getUserLicenseStatus(this.context.pageContext.site.absoluteUrl, this.context.serviceScope);
    return Promise.resolve(null);
  }
  
  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {

    if (this._userLicenseStatus && this._userLicenseStatus.Status === UserLicenseStatus.None) {
      event.domElement.innerHTML = YouHaveNoWorkPoint365License;
      return null;
    }

    // Use this method to perform your custom cell rendering.
    const urlValue: string = event.fieldValue;
    const title:string = event.listItem.getValueByName("Title");
    const itemId:number = event.listItem.getValueByName("ID");
    const context:FieldCustomizerContext = this.context;
    const listId:string = context.pageContext.list.id.toString();
    const solutionAbsoluteUrl:string = context.pageContext.site.absoluteUrl;

    const siteField: React.ReactElement<{}> = React.createElement(SiteField, { urlValue, title, listId, itemId, solutionAbsoluteUrl } as ISiteFieldProps);

    ReactDOM.render(siteField, event.domElement);
  }

  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    ReactDOM.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }
}
