import * as React from 'react';
import Transition from 'react-transition-group/Transition';

export interface IRollInProps {
  inProp: boolean;
  duration: number;
  top: number;
  mountOnEnter: boolean;
  unmountOnExit: boolean;
  origin: "left"|"right";
}

export const RollIn:React.SFC<IRollInProps> = ({ inProp, duration, top, mountOnEnter, unmountOnExit, origin, children }):JSX.Element => {

  let defaultStyle: React.CSSProperties = null;
  let transitionStyles: object = null;

  switch (origin) {
    case "left": {

      defaultStyle = {
        transition: `transform ${duration}ms ease-out`,
        transform: `translateX(-100%)`,
        position: "absolute",
        left: 0,
        top,
        bottom: 0,
        height: `calc(100% - ${top}px)`
      };
    
      transitionStyles = {
        entering: { transform: `translateX(-100%)` },
        entered: { transform: `translateX(0)` },
        exiting: { transform: `translateX(-100%)` }
      };

      break;
    }
    case "right": {

      defaultStyle = {
        transition: `transform ${duration}ms ease-out`,
        transform: `translateX(100%)`,
        position: "absolute",
        right: 0,
        top,
        bottom: 0,
        height: `calc(100% - ${top}px)`
      };
    
      transitionStyles = {
        entering: { transform: `translateX(100%)` },
        entered: { transform: `translateX(0)` },
        exiting: { transform: `translateX(100%)` }
      };

      break;
    }
  }

  return (
    <Transition
      in={inProp}
      timeout={duration}
      mountOnEnter={mountOnEnter}
      unmountOnExit={unmountOnExit}>
      {(state) => (
        <div style={{
          ...defaultStyle,
          ...transitionStyles[state]
        }}>
          {children}
        </div>
      )}
    </Transition>
  );
};