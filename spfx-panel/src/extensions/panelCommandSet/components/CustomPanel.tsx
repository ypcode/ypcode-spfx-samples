import * as React from 'react';
import { TextField, DefaultButton, PrimaryButton, DialogFooter, Panel, Spinner, SpinnerType, PanelType } from "office-ui-fabric-react";
import { sp } from "@pnp/sp";
import { autobind } from '@uifabric/utilities';

const pnpSuperHero = require("../../../assets/pnphero.png");

export interface ICustomPanelState {
  saving: boolean;
  editedTitle: string;
}

export interface ICustomPanelProps {
  onClose: () => void;
  isOpen: boolean;
  currentTitle: string;
  itemId: number;
  listId: string;
}

export default class CustomPanel extends React.Component<ICustomPanelProps, ICustomPanelState> {

  constructor(props: ICustomPanelProps) {
    super(props);
    this.state = {
      saving: false,
      editedTitle: props.currentTitle
    };
  }

  @autobind
  private _onTitleChanged(title: string) {
    this.setState({editedTitle: title});
  }

  @autobind
  private _onCancel() {
    this.props.onClose();
  }

  @autobind
  private _onApply() {
    this.setState({ saving: true });
    // Update the list item title using PnP JS
    sp.web.lists.getById(this.props.listId).items.getById(this.props.itemId).update({
      'Title': this.state.editedTitle
    }).then(() => {
      this.setState({ saving: false });
      this.props.onClose();
    });
  }

  private _onRenderFooterContent = (): JSX.Element => {
    return (
      <DialogFooter>
        <PrimaryButton onClick={this._onApply} style={{ marginRight: '8px' }}>
          Apply
          </PrimaryButton>
        <DefaultButton onClick={this._onCancel}>Cancel</DefaultButton>
      </DialogFooter>
    );
  }

  public render(): React.ReactElement<ICustomPanelProps> {
    let { isOpen } = this.props;
    return (
      <Panel isOpen={isOpen} type={PanelType.medium} onRenderFooterContent={this._onRenderFooterContent}>
        <h2>Item editor</h2>
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md6" style={{"textAlign":"center"}}>
            <img style={{ width: "100px" }} src={`${pnpSuperHero}`} />
          </div>
          <div className="ms-Grid-col ms-sm12 ms-md6">
            <TextField value={this.state.editedTitle} onChanged={this._onTitleChanged} label="Item title" placeholder="Choose the new title" />
            <p>
              From this form you can simply update the title of the selected item.
              Change the name in the field above and click the "Apply" button.
              If you prefer to discard your changes, click the "Cancel" button or the upper right corner cross
            </p>
          </div>
        </div>
        {this.state.saving && <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12">
            <Spinner type={SpinnerType.large} label="Saving..." />
          </div>
        </div>}
      </Panel>
    );
  }
}
