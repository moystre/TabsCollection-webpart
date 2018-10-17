import * as React from 'react';

export enum IconSize {
  microscopic = 10,
  tiny = 16,
  small = 24,
  medium = 32,
  large = 48,
  huge = 64
}

export interface IIconData {
  iconUrl?: string;
  iconClass?: string;
  className?: string;
  size?: IconSize;
  style?: object;
}

export const Icon = ({ size = IconSize.tiny, iconUrl, iconClass, className = '', style }: IIconData): JSX.Element => {

  let iconElement: JSX.Element;

  const sizeStyle: object = {
    width: `${size}px`,
    height: `${size}px`
  };

  /**
   * This size styling is only relevant for image icons.
   */
  const newStyle = {...style, ...sizeStyle };

  if (iconUrl) {
    iconElement = <img style={newStyle} aria-hidden="true" src={iconUrl} className={className} />;
  } else if (iconClass) {
    iconElement = <span aria-hidden="true" style={style} className={`ms-Icon ms-Icon--${iconClass} ${className}`} />;
  } else {
    iconElement = null;
  }

  return iconElement;
};