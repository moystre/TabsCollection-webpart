import { Callout, DirectionalHint } from 'office-ui-fabric-react/lib/Callout';
import * as React from 'react';
import * as strings from 'WorkPointStrings';
import styles from './ContextMenu.module.scss';
import { Icon } from './Icon';


//import { CalloutContent } from 'office-ui-fabric-react/lib/components/Callout/CalloutContent';
//import { IStyleFunction } from 'office-ui-fabric-react/lib/Utilities';
export interface IOnTopElement {
  title: string;
  iconUrl?: string;
  iconClass?: string;
  url?: string;
}

const OnTopElement:React.SFC<IOnTopElement> = (props):JSX.Element => {
  
  const { title, iconUrl, iconClass, url } = props;

  const iconSpace: boolean = (iconUrl === undefined && iconClass === undefined) ? false : true;

  const elementContent:JSX.Element = (
    <div className={styles.iconAndText}>
      <Icon iconClass={iconClass} iconUrl={iconUrl} />
      {title &&
        <span style={iconSpace ? {marginLeft: 8} : null} className={styles.text}>{title}</span>
      }
    </div>
  );

  if (url) {

    return (
      <a href={url} title={`${strings.GoTo} ${title}`} className={styles.clickableOnTopElement}>
        {elementContent}
      </a>
    );
  } else {

    return (
      <div className={styles.onTopElement}>
        {elementContent}
      </div>
    );
  }
};

export interface IContextMenuProps {
  target: HTMLElement;
  close(): void;
  onTopElement?: IOnTopElement | null;
}

export default class ContextMenu extends React.Component<IContextMenuProps> {

  public render ():JSX.Element | null {

    const { onTopElement, target, close, children } = this.props;

    let contextMenuStyle:React.CSSProperties;
    
    let coverTarget: boolean = false;
    let directionalHint:DirectionalHint = DirectionalHint.rightCenter;

    if (onTopElement) {
      coverTarget = true;
      directionalHint = DirectionalHint.topAutoEdge;
    }

    const targetBoundingRect = target.getBoundingClientRect();
    const targetWidth = targetBoundingRect.width - 2 /* Accounts for 1px border */;

    contextMenuStyle = {
      minWidth: targetWidth
    };

    // TODO: When office-ui-fabric-react is updated we will be able to style components individually
    //const styleFunction:IStyleFunction<ICalloutContentStyleProps, ICalloutContentStyles> = (props) => ({calloutMain: styles.callout, root: styles.root});

    return (
      <Callout
        target={target}
        isBeakVisible={false}
        directionalHint={directionalHint}
        coverTarget={coverTarget}
        className={styles.root}
        //getStyles={styleFunction}
        onDismiss={close}
      >
          <div className={styles.layerContent} style={coverTarget ? contextMenuStyle : null}>
            {coverTarget &&
              <OnTopElement {...onTopElement}/>
            }
            {children}
          </div>

      </Callout>
    );
  }
}