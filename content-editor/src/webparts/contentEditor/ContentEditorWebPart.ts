import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IWebPartPropertiesMetadata,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';

import * as strings from 'ContentEditorWebPartStrings';
import ContentEditor from './components/ContentEditor';
import { IContentEditorProps } from './components/IContentEditorProps';
import { IContentEditorWebPartProps } from './IContentEditorWebPartProps';

import pnp from "sp-pnp-js";
import {
  IContentService,
  ContentService,
  IContentServiceConfiguration,
  SourceFormat,
  SourceType
} from "./services/ContentService";

export default class ContentEditorWebPart extends BaseClientSideWebPart<IContentEditorWebPartProps> {

  private contentService: IContentService;

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {

      pnp.setup({
        spfxContext: this.context
      });

      this.contentService = new ContentService();
      this.contentService.configure(this.properties);

      console.log(this.properties);

    });
  }

  public render(): void {
    const element: React.ReactElement<IContentEditorProps> = React.createElement(
      ContentEditor,
      {
        contentService: this.contentService,
        displayMode: this.displayMode,
        showCaption: this.properties.showCaption
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      'sourceContent': { isHtmlString: true },
      'sourceLink': { isLink: true }
    };
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let groupSourceConfigFields: any = [
      PropertyPaneDropdown('sourceType', {
        label: strings.SourceTypeLabel,
        selectedKey: this.properties.sourceType != null ? this.properties.sourceType : SourceType.Content,
        options: [
          { key: SourceType.Content, text: strings.ContentSourceTypeOption },
          { key: SourceType.Link, text: strings.LinkSourceTypeOption }
        ]
      })
    ];

    if (this.properties.sourceType == SourceType.Link) {
      groupSourceConfigFields.push(PropertyPaneTextField('sourceLink', {
        label: strings.SourceLinkLabel
      }));
    }

    groupSourceConfigFields.push(PropertyPaneDropdown('sourceFormat', {
      label: strings.SourceFormatLabel,
      selectedKey: this.properties.sourceFormat != null ? this.properties.sourceFormat  : SourceFormat.Markdown,
      options: [
        { key: SourceFormat.Auto, text: "Auto" },
        { key: SourceFormat.Html, text: "HTML" },
        { key: SourceFormat.Markdown, text: "Markdown" }
      ]
    }));

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneHeader
          },
          groups: [
            {
              groupName: strings.SourceConfigGroup,
              groupFields: groupSourceConfigFields
            },
            {
              groupName: strings.displaySettingsGroupLabel,
              groupFields: [
                PropertyPaneToggle("showCaption", {
                  label: strings.showCaptionSwitchLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
