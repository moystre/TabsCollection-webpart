import * as React from 'react';

import breadcrumbStyles from './Breadcrumb.module.scss';
import { Icon } from './Icon';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { IWorkPointBaseProps } from './WorkPointNavBarInterfaces';

export interface IHomeBreadcrumbData extends IWorkPointBaseProps {
  text?: string;
  title: string;
  onMouseUp: React.EventHandler<React.MouseEvent<HTMLDivElement>>;
  iconClass: string;
  iconUrl?: string;
  style: object;
}

export default class HomeBreadcrumb extends React.Component<IHomeBreadcrumbData> {

  constructor(props: IHomeBreadcrumbData) {
    super(props);
  }

  protected navigateToRootSite = (event:any):void => {
    const mouseEvent = event as MouseEvent;
    window.location.href = this.props.context.solutionAbsoluteUrl;
  }

  public render(): JSX.Element {
    const { text, onMouseUp, iconClass, iconUrl, title, style } = this.props;

    return (
      <div 
        className={breadcrumbStyles.breadcrumb}
        style={style}
        title={title}
        onMouseUp={onMouseUp}
        onDoubleClick={this.navigateToRootSite}
      >
        <div className={breadcrumbStyles.iconAndText}>
          <Icon iconClass={iconClass} iconUrl={iconUrl} />
          {text &&
            <span className={breadcrumbStyles.text}>{text}</span>
          }
        </div>
      </div>
    );
  }
}