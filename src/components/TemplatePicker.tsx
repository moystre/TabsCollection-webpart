import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';
import { Button, DialogContent, DialogFooter } from 'office-ui-fabric-react';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import * as strings from 'WorkPointStrings';
import { IBusinessModuleSettings } from '../workPointLibrary/BusinessModule';
import { getTemplateLibraryContentTypes, getTemplateLibraryMappings, getWorkPointSettings } from '../workPointLibrary/service';
import { getCreateWordUrl } from '../workPointLibrary/Template';

export interface IWordTemplate {
  title: string;
  href: string;
}

interface INewTemplatePickerDialogContentProps {
  templates: IWordTemplate[];
  message: string;
  close: () => void;
  submit: (color: string) => void;
}

class TemplatePickerDialogContent extends React.Component<INewTemplatePickerDialogContentProps, null> {

  constructor(props:INewTemplatePickerDialogContentProps) {
    super(props);
  }

  public render(): JSX.Element {

    return (
      <DialogContent
        title={strings.NewWordTemplate}
        subText={this.props.message}
        onDismiss={this.props.close}
        showCloseButton={true}
      >
        {this.props.templates.length > 0 && this.props.templates.map(template => {
          return <Button text={template.title} href={template.href} />;
        })}
        <DialogFooter>
          <Button text={strings.Cancel} title={strings.Cancel} onClick={this.props.close} />
        </DialogFooter>
      </DialogContent>
    );
  }
}

export default class WordTemplatePickerDialog extends BaseDialog {
  public context: ListViewCommandSetContext;
  protected templates: IWordTemplate[];

  private createWordTemplateUrl = (businessModuleId:string, templateAbsoluteUrl:string):string => {

    const newWordTemplateUrl:string = getCreateWordUrl(10, businessModuleId, this.context.pageContext.site.absoluteUrl, templateAbsoluteUrl);
    return newWordTemplateUrl;
  }

  public async onBeforeOpen():Promise<void> {

    let templates:IWordTemplate[] = [];

    try {

      const [templateLibraryMappings, workPointSettingsCollection, templateLibraryContentTypes] = await Promise.all([
        getTemplateLibraryMappings(this.context.pageContext.site.absoluteUrl, this.context.pageContext.site.serverRelativeUrl),
        getWorkPointSettings(this.context.pageContext.site.absoluteUrl, this.context.pageContext.site.serverRelativeUrl, this.context.pageContext.cultureInfo.currentUICultureName),
        getTemplateLibraryContentTypes(this.context.pageContext.site.absoluteUrl, this.context.pageContext.list.id.toString())
      ]);
  
      const businessModuleSettings:IBusinessModuleSettings[] = workPointSettingsCollection.businessModuleSettings.settings;
  
      templates = businessModuleSettings.map(bmSetting => {
  
        try {
  
          const title:string = bmSetting.Title;
          const businessModuleId:string = bmSetting.Id;
          let contentTypeId: string = null;
          let templateUrl: string = null;
          
          const loweredBusinessModuleId = bmSetting.Id.toLowerCase();
    
          // Check if current business module has mapped a content type in the template library
          for (let mappingContentTypeId in templateLibraryMappings) {
            const loweredTemplateMappingBusinessModuleId = templateLibraryMappings[mappingContentTypeId].toLowerCase();
    
            if (loweredTemplateMappingBusinessModuleId === loweredBusinessModuleId) {
              contentTypeId = mappingContentTypeId;
              break;
            }
          }
    
          if (!contentTypeId) {
            throw "No mapping was found";
          }
    
          // Find associated DocumentTemplateUrl for the given content type.
          for (let templateLibraryContentType of templateLibraryContentTypes) {
            const loweredTemplateLibraryContentTypeId = templateLibraryContentType.StringId.toLowerCase();
            const loweredBusinessModuleContentTypeId = contentTypeId.toLowerCase();
    
            if (loweredTemplateLibraryContentTypeId === loweredBusinessModuleContentTypeId) {
              templateUrl = templateLibraryContentType.DocumentTemplateUrl;
            }
          }
  
          const templateAbsoluteUrl = templateUrl.replace(this.context.pageContext.site.serverRelativeUrl, this.context.pageContext.site.absoluteUrl);
          const href:string = this.createWordTemplateUrl(businessModuleId, templateAbsoluteUrl);
    
          let template:IWordTemplate = {
            title,
            href
          };
  
          return template;
  
        } catch (exception) {
          return null;
        }
      }).filter(template => template !== null);
  
      try {
  
        // Find template for "No MailMerge"
        for (let mappingContentTypeId in templateLibraryMappings) {
          if (templateLibraryMappings[mappingContentTypeId] === '00000000-0000-0000-0000-000000000000') {
            for (let templateLibraryContentType of templateLibraryContentTypes) {
  
              const loweredContentTypeId = templateLibraryContentType.StringId.toLowerCase();
              const loweredMappingContentTypeId = mappingContentTypeId.toLowerCase();
  
              if (loweredContentTypeId === loweredMappingContentTypeId) {
    
                const templateAbsoluteUrl:string = templateLibraryContentType.DocumentTemplateUrl.replace(this.context.pageContext.site.serverRelativeUrl, this.context.pageContext.site.absoluteUrl);
                templates.push({
                  title: "Template with no mail-merge",
                  href: `ms-word:nft|u|${templateAbsoluteUrl}|s|${this.context.pageContext.site.absoluteUrl}/Template Library`
                });
                break;
              }
            }
          }
        }
  
      } catch (exception) {
        console.warn("New Word Template item could not be created for the No mail merge content type");
      }

    } catch (exception) {
      console.warn(`Word templates could not be loaded. It failed with the following error: ${exception}`);
    }

    this.templates = templates;
  }

  public render(): void {

    let message: string;
    message = `${strings.ChooseABusinessModuleToCreateWordTemplateFor}. ${strings.YouCanAlsoCreateATemplateWithoutMailMerge}.`;

    ReactDOM.render(<TemplatePickerDialogContent
      templates={this.templates}
      close={ this.close }
      message={ message }
      submit={ this._submit }
    />, this.domElement);
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false
    };
  }

  private _submit = ():void => {
    this.close();
  }
}