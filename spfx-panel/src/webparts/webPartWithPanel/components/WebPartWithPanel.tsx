import * as React from 'react';
import styles from './WebPartWithPanel.module.scss';
import { IWebPartWithPanelProps } from './IWebPartWithPanelProps';
import { Button, PrimaryButton, DefaultButton } from "office-ui-fabric-react/lib/Button";
import { Panel } from "office-ui-fabric-react/lib/Panel";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { autobind } from '@uifabric/utilities';

export interface IWebPartWithPanelState {
  showPanel: boolean;
  message: string;
  messageDraft: string;
}

const DEFAULT_MESSAGE = "Hello PnP !";

export default class WebPartWithPanel extends React.Component<IWebPartWithPanelProps, IWebPartWithPanelState> {

  constructor(props: IWebPartWithPanelProps) {
    super(props);
    this.state = {
      message: DEFAULT_MESSAGE,
      messageDraft: DEFAULT_MESSAGE,
      showPanel: false
    };
  }

  @autobind
  private _openPanel() {
    this.setState({ showPanel: true });
  }

  @autobind
  private _onCancel() {
    this.setState({
      showPanel: false,
      messageDraft: this.state.message
    })
  }

  @autobind
  private _onMessageChanged(messageDraft: string) {
    this.setState({ messageDraft });
  }

  @autobind
  private _onApply() {
    this.setState({
      message: this.state.messageDraft,
      showPanel: false
    });
  }

  private _onRenderFooterContent = (): JSX.Element => {
    return (
      <div>
        <PrimaryButton onClick={this._onApply} style={{ marginRight: '8px' }}>
          Apply
        </PrimaryButton>
        <DefaultButton onClick={this._onCancel}>Cancel</DefaultButton>
      </div>
    );
  };

  public render(): React.ReactElement<IWebPartWithPanelProps> {
    let { showPanel, message, messageDraft } = this.state;
    return (
      <div className={styles.webPartWithPanel}>
        <div className={styles.container}>
          <Panel isOpen={showPanel}
            onDismiss={this._onCancel}
            headerText="My Custom panel"
            onRenderFooterContent={this._onRenderFooterContent}>
            <TextField label="Message" value={messageDraft} onChanged={this._onMessageChanged} />
          </Panel>
          <h1>{message}</h1>
          <Button text="Open panel" onClick={this._openPanel} />
        </div>
      </div>
    );
  }
}
