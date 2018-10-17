import { Callout, CommandBar, DirectionalHint } from 'office-ui-fabric-react';
import * as React from 'react';
import * as strings from 'WorkPointStrings';

export default class DocumentCallout extends React.Component<any, any> {
  constructor(props: any) {
    super(props);
  }

  public render(): React.ReactElement<any> {
    const { data, viewRelativeUrl } = this.props;
    const webUrl = this.props.context.pageContext.web.absoluteUrl;
    return (
      <Callout
        className='ms-DocumentCallout-callout'
        role={'alertdialog'}
        gapSpace={0}
        target={this.props.targetElement}
        directionalHint={DirectionalHint.rightCenter}
        onDismiss={this.props.onCalloutDismiss}
        setInitialFocus={true}
      >
        <div className='ms-DocumentCallout-header'>
          <p className='ms-DocumentCallout-title' id={'callout-label-1'}>
            {data.Title}
          </p>
        </div>
        <div className='ms-DocumentCallout-inner'>
          <div className='ms-DocumentCallout-content'>
            <img src={webUrl + '/_layouts/15/getpreview.ashx?path=' + data.FileRef} />
          </div>
        </div>
        <CommandBar items={[
          {
            key: 'openLink',
            name: strings.LinkToItem,
            className: 'ms-CommandBarItem',
            icon: 'Link',
            href: decodeURI(viewRelativeUrl + '?id=' + data.FileRef + '&parent=' + data.FileDirRef),
            target: '_blank'
          }
        ]} />
      </Callout>
    );
  }
}