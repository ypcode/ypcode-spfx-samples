import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import {sp} from "@pnp/sp";

import * as strings from 'PanelCommandSetCommandSetStrings';

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

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized PanelCommandSetCommandSet');

    // Setup the PnP JS with SPFx context
    sp.setup({
      spfxContext: this.context
    });


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
        alert("The command is executed !");
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
