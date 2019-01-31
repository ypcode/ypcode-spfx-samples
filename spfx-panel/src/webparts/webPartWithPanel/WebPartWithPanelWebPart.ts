import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'WebPartWithPanelWebPartStrings';
import WebPartWithPanel from './components/WebPartWithPanel';
import { IWebPartWithPanelProps } from './components/IWebPartWithPanelProps';

export interface IWebPartWithPanelWebPartProps {
  description: string;
}

export default class WebPartWithPanelWebPart extends BaseClientSideWebPart<IWebPartWithPanelWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IWebPartWithPanelProps > = React.createElement(
      WebPartWithPanel,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
