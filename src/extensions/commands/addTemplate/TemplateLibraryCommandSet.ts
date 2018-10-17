import { BaseListViewCommandSet, IListViewCommandSetExecuteEventParameters, IListViewCommandSetListViewUpdatedParameters } from '@microsoft/sp-listview-extensibility';
import WordTemplatePickerDialog from '../../../components/TemplatePicker';
import { IWizardStartMessage } from '../../../workPointLibrary/Event';
import { getUserLicenseStatus, IUserLicense, UserLicenseStatus } from '../../../workPointLibrary/License';

export default class TemplateLibraryCommandsCommandSet extends BaseListViewCommandSet<null> {

  public async onInit(): Promise<void> {
    
    const userLicenseStatus:IUserLicense = await getUserLicenseStatus(this.context.pageContext.site.absoluteUrl, this.context.serviceScope);

    if (userLicenseStatus && userLicenseStatus.Status === UserLicenseStatus.None) {
      return Promise.reject("User is not licensed");
    }
    
    return Promise.resolve(null);
  }

  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    /**
     * Not used at the moment. We do not have any buttons that depend on anything selected.
     * 
     * TODO: Maybe we could show an "Edit WorkPoint 365 template" when selecting something? 
     *
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
    */
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {

    switch (event.itemId) {
      case 'ADD_WORD_TEMPLATE': {
        const dialog: WordTemplatePickerDialog = new WordTemplatePickerDialog();
        dialog.context = this.context;
        dialog.show();
        break;
      }

      case 'ADD_EXCEL_TEMPLATE': {
        const wizardUrl:string = `MailMergeTemplate`;
        const wizardMessage: IWizardStartMessage = {
          type: "workpointwizard",
          url: wizardUrl
        };
        window.postMessage(wizardMessage, this.context.pageContext.site.absoluteUrl);
        break;
      }
      default:
        throw new Error('Unknown command');
    }
  }
}
