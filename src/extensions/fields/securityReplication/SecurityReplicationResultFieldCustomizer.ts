import { BaseFieldCustomizer, IFieldCustomizerCellEventParameters } from '@microsoft/sp-listview-extensibility';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import JobResult from '../../../components/JobResult';

export default class SecurityReplicationResultFieldCustomizer
  extends BaseFieldCustomizer<null> {

  public onInit(): Promise<void> {
    return Promise.resolve();
  }

  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    const resultField: React.ReactElement<{}> = React.createElement(JobResult, { result: event.fieldValue });
    ReactDOM.render(resultField, event.domElement);
  }

  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    ReactDOM.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }
}