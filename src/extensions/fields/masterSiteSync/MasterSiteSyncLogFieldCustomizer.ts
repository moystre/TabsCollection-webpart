import { BaseFieldCustomizer, IFieldCustomizerCellEventParameters } from '@microsoft/sp-listview-extensibility';
import * as moment from 'moment';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import LogLink from '../../../components/LogLink';

export default class MasterSiteSyncLogFieldCustomizer
  extends BaseFieldCustomizer<null> {

  public onInit(): Promise<void> {
    
    // Setup the moment locale, based on the currentUICultureName
    moment.locale(this.context.pageContext.cultureInfo.currentUICultureName);
    return Promise.resolve();
  }

  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    // Use this method to perform your custom cell rendering.
    const loggingScopeId: string = event.fieldValue;
    const dateValue:string = event.listItem.getValueByName("wpMasterSiteSyncDate");

    // No loggingScopeId? Bail out
    if (typeof loggingScopeId !== "string" || loggingScopeId === "") {
      return null;
    }

    let logEndTime:moment.Moment = null;

    if (typeof dateValue === "string" && dateValue !== "") {
      logEndTime = moment(dateValue).add(30, "minutes");
    }
  
    const logLinkField: React.ReactElement<{}> = React.createElement(LogLink, { loggingScopeId, logEndTime, pageContext: this.context.pageContext, spHttpClient: this.context.spHttpClient });

    ReactDOM.render(logLinkField, event.domElement);
  }

  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    ReactDOM.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }
}