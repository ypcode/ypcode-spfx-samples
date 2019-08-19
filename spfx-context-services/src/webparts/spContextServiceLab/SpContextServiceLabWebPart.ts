import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, ServiceScope } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import { FirstLevelSubComponent, IComponentProps } from './components/componentsHierarchy';
import { configure } from '../../startup/configure';
import { ComponentContextServiceKey } from '../../services/ComponentContextService';

export interface ISpContextServiceLabWebPartProps {
  documentLibraryName: string;
}

export default class SpContextServiceLabWebPart extends BaseClientSideWebPart<ISpContextServiceLabWebPartProps> {

  private componentServiceScope: ServiceScope;

  public onInit(): Promise<void> {
    // Make sure to return a resolved promise when configuration is done
    // This will ensure all services are properly configured so we can safely call serviceScope.consume() in any component 
    return configure(this.context, this.properties)
      .then(serviceScope => {
        this.componentServiceScope = serviceScope;
      }).catch(error => {
        console.error('An error occured while trying to initialize WebPart', error);
      });
  }

  public render(): void {
    const element: React.ReactElement<IComponentProps> = React.createElement(
      FirstLevelSubComponent,
      {
        serviceScope: this.componentServiceScope
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public onPropertyPaneFieldChanged(property: string, oldValue: any, newValue: any) {
    // Update the value from configuration when changed
    const componentContextService = this.componentServiceScope.consume(ComponentContextServiceKey);
    componentContextService.properties[property] = newValue;
    this.render();
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
            description: "Settings"
          },
          groups: [
            {
              groupName: "Configuration",
              groupFields: [
                PropertyPaneTextField('documentLibraryName', {
                  label: "Name of the Documents library"
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
