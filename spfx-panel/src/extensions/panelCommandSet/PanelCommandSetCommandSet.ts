import * as React from 'react';
import * as ReactDom from 'react-dom';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { sp } from "@pnp/sp";

import * as strings from 'PanelCommandSetCommandSetStrings';
import { autobind, assign } from '@uifabric/utilities';

import CustomPanel, { ICustomPanelProps } from "./components/CustomPanel";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IPanelCommandSetCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'PanelCommandSetCommandSet';

export default class PanelCommandSetCommandSet extends BaseListViewCommandSet<IPanelCommandSetCommandSetProperties> {

  private panelDomElement: HTMLDivElement;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized PanelCommandSetCommandSet');

    // Setup the PnP JS with SPFx context
    sp.setup({
      spfxContext: this.context
    });

    this.panelDomElement = document.body.appendChild(document.createElement("div"));

    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const openEditorCommand: Command = this.tryGetCommand('CMD_PANEL');
    openEditorCommand.visible = event.selectedRows.length === 1;
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {

    switch (event.itemId) {
      case 'CMD_PANEL':
        let selectedItem = event.selectedRows[0];
        const listItemId = selectedItem.getValueByName('ID') as number;
        const title = selectedItem.getValueByName("Title");
        this._showPanel(listItemId, title);
        break;
      default:
        throw new Error('Unknown command');
    }
  }


  private _showPanel(itemId: number, currentTitle: string) {
    this._renderPanelComponent({
      isOpen: true,
      currentTitle,
      itemId,
      listId: this.context.pageContext.list.id.toString(),
      onClose: this._dismissPanel
    });
  }

  @autobind
  private _dismissPanel() {
    this._renderPanelComponent({ isOpen: false });
  }

  private _renderPanelComponent(props: any) {
    const element: React.ReactElement<ICustomPanelProps> = React.createElement(CustomPanel, assign({
      onClose: null,
      currentTitle: null,
      itemId: null,
      isOpen: false,
      listId: null
    }, props));
    ReactDom.render(element, this.panelDomElement);
  }

}
