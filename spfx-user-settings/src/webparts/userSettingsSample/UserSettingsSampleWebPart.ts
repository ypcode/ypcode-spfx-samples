import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, ServiceScope } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'UserSettingsSampleWebPartStrings';
import UserSettingsSample from './components/UserSettingsSample';
import { IUserSettingsSampleProps } from './components/IUserSettingsSampleProps';
import { UserPreferencesServiceKey, UserPreferencesService } from '../../services/UserPreferencesService';


export interface IUserSettingsSampleWebPartProps {
  description: string;
}

export default class UserSettingsSampleWebPart extends BaseClientSideWebPart<IUserSettingsSampleWebPartProps> {

  private _usedServiceScope: ServiceScope;

  public onInit(): Promise<void> {
    return super.onInit().then(() => {
      // Create a child scope for the current WebPart
      this._usedServiceScope = this.context.serviceScope.startNewChild();
      // Configure it with the current WebPart instance Id
      const serviceInstance = this._usedServiceScope.createAndProvide(UserPreferencesServiceKey, UserPreferencesService);
      serviceInstance.configure(this.instanceId);
      this._usedServiceScope.finish();
    });
  }

  public render(): void {

    const element: React.ReactElement<IUserSettingsSampleProps> = React.createElement(
      UserSettingsSample,
      {
        serviceScope: this._usedServiceScope
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
