import { Link } from 'office-ui-fabric-react';
import * as React from 'react';
import * as strings from 'WorkPointStrings';
import { IWizardStartMessage } from '../workPointLibrary/Event';
import { ITemplateLibrarySettings } from '../workPointLibrary/service';
import { getEditWordUrl } from '../workPointLibrary/Template';

export interface ITemplateEditFieldProps {
  name: string;
  fileName: string;
  fileType: "Word" | "Excel";
  contentTypeId: string;
  templateLibrarySettings: ITemplateLibrarySettings;
  id: string;
  solutionAbsoluteUrl: string;
  fileUrl: string;
}

export interface ITemplateEditFieldState {
  businessModuleId: string;
  templateAbsoluteUrl: string;
}

export default class TemplateEditField extends React.Component<ITemplateEditFieldProps, ITemplateEditFieldState> {

  constructor (props:ITemplateEditFieldProps) {
    super(props);
    
    const businessModuleId:string = this.props.templateLibrarySettings[this.props.contentTypeId] || "00000000-0000-0000-0000-000000000000";

    const templateAbsoluteUrl:string = `${location.protocol}//${location.host}${this.props.fileUrl}`;

    this.state = {
      businessModuleId,
      templateAbsoluteUrl
    };
  }

  /**
   * Opens Excel template edit form based on a business module id and item id.
   */
  protected openExcelTemplateEditForm = ():void => {
    const wizardUrl:string = `MailMergeTemplate/Edit?businessModuleListID=${this.state.businessModuleId}&itemID=${this.props.id}`;
    const wizardMessage: IWizardStartMessage = {
      type: "workpointwizard",
      url: wizardUrl
    };
    window.postMessage(wizardMessage, this.props.solutionAbsoluteUrl);
  }

  public render ():JSX.Element {

    const { businessModuleId, templateAbsoluteUrl } = this.state;
    
    switch (this.props.fileType) {
      case "Excel": {

        if (businessModuleId && businessModuleId !== "00000000-0000-0000-0000-000000000000") {
          return <Link title={strings.EditTemplate} onClick={this.openExcelTemplateEditForm}>{this.props.name}</Link>;
        } else {
          return <Link title={strings.OpenTemplate} href={`ms-excel:ofe|u|${templateAbsoluteUrl}`}>{this.props.name}</Link>;
        }
      }
      case "Word": {

        if (businessModuleId && businessModuleId !== "00000000-0000-0000-0000-000000000000") {

          const editTemplateUrl:string = getEditWordUrl(10, businessModuleId, this.props.solutionAbsoluteUrl, templateAbsoluteUrl);

          return <Link title={strings.EditTemplate} href={editTemplateUrl}>{this.props.name}</Link>;
        } else {
          return <Link title={strings.OpenTemplate} href={`ms-word:ofe|u|${templateAbsoluteUrl}`}>{this.props.name}</Link>;
        }
      }
      default:
        return <span>{this.props.name}</span>;

    }
  }
}