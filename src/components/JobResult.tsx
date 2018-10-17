import { Icon } from "office-ui-fabric-react";
import * as React from "react";
import { JobFailed, JobSucceded } from 'WorkPointStrings';
import styles from './JobResult.module.scss';

export interface IJobResultProps {
  result: "succeded" | "failed";
}

export default class JobResult extends React.Component<IJobResultProps, null> {
  public render ():JSX.Element {

    const { result } = this.props;

    let displayElement: JSX.Element = <span>{result}</span>;

    try {

      if (typeof result === "string") {

        const loweredResult:string = result.toLowerCase();

        switch (loweredResult) {
          case "succeded": {
            displayElement = <Icon className={styles.success} title={JobSucceded} iconName="Accept" />;
            break;
          }
          case "failed": {
            displayElement = <Icon className={styles.warning} title={JobFailed} iconName="Warning" />;
          }
        }

      }
    } catch (exception) {}

    return displayElement;
  }
}