import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderName,
  PlaceholderContent
} from '@microsoft/sp-application-base';

import * as strings from 'MotdApplicationCustomizerStrings';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from "./motd.module.scss";
import pnp from "sp-pnp-js";

const LOG_SOURCE: string = 'MotdApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IMotdApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class MotdApplicationCustomizer
  extends BaseApplicationCustomizer<IMotdApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    return new Promise((resolve, reject) => {
      pnp.setup({
        spfxContext: this.context
      });

      pnp.sp.web.lists.getByTitle("MOTD")
        .items.select("Title", "Message")
        .orderBy("Id", false)
        .top(1)
        .get().then(items => {

          let motd = items && items.length && items[0];
          if (!motd) {
            resolve();
            return;
          }

          let topPlaceholder: PlaceholderContent = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);
          if (topPlaceholder) {
            topPlaceholder.domElement.innerHTML = `
            <div class="${styles.motd}">
              <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.top}">
                <div class="${styles.title}">
                  ${escape(motd.Title)}
                </div>&nbsp;-&nbsp;
                ${escape(motd.Message)}</div>
              </div>
            </div>`;
          }
          resolve();
        }).catch(error => {
          reject();
        });

    });
  }
}
