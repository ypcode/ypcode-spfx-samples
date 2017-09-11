import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';

import pnp from "sp-pnp-js";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IMarkAsImportantCommandSetProperties {
  // This is an example; replace with your own property
  disabledCommandIds: string[] | undefined;
}

const LOG_SOURCE: string = 'MarkAsImportantCommandSet';

export default class MarkAsImportantCommandSet
  extends BaseListViewCommandSet<IMarkAsImportantCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized MarkAsImportantCommandSet');

    pnp.setup({
      spfxContext: this.context
    });

    return Promise.resolve<void>();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    if (this.properties.disabledCommandIds) {
      for (const commandId of this.properties.disabledCommandIds) {
        const command: Command | undefined = this.tryGetCommand(commandId);
        if (command && command.visible) {
          Log.info(LOG_SOURCE, `Hiding command ${commandId}`);
          command.visible = false;
        }
      }
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.commandId) {
      case 'CMD_MARK_AS_IMPORTANT':
        console.log(event);
        for (let i = 0; i < event.selectedRows.length; i++) {
          let row = event.selectedRows[i];
          let id = row.getValueByName("ID");
          console.log("Found ID: ");
          console.log(id);
          pnp.sp.web.lists.getByTitle(this.context.pageContext.list.title)
            .items.getById(id)
            .update({
              IsImportant: true
            })
            .then(result =>
              console.log("Successfully updated item"))
            .catch(error => console.log(error));

        }
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
