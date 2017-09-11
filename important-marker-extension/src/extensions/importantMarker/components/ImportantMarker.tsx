import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import * as React from 'react';
import { Icon } from "office-ui-fabric-react";

import styles from './ImportantMarker.module.scss';

export interface IImportantMarkerProps {
  text: string;
}

const LOG_SOURCE: string = 'ImportantMarker';

export default class ImportantMarker extends React.Component<IImportantMarkerProps, {}> {
  @override
  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: ImportantMarker mounted');
  }

  @override
  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: ImportantMarker unmounted');
  }

  @override
  public render(): React.ReactElement<{}> {
    return (
      <div className={styles.cell}>
        <span className="ms-fontColor-themePrimary">
          <Icon name="Important" />
        </span>
        {this.props.text}
      </div>
    );
  }
}
