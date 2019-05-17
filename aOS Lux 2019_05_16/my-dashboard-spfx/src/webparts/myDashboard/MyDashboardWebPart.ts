import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, PropertyPaneToggle } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'MyDashboardWebPartStrings';
import MyDashboard, {IMyDashboardProps} from './components/MyDashboard';

import * as microsoftTeams from '@microsoft/teams-js';

export interface IMyDashboardWebPartProps {
  welcomeMessage: string;
  showDebug: boolean;
}

export default class MyDashboardWebPart extends BaseClientSideWebPart<IMyDashboardWebPartProps> {

  protected onInit(): Promise<any> {
    let retVal: Promise<any> = Promise.resolve();
    // NOTE Work around for SPFx 1.8.1 missing microsoftTeams typings
    let microsoftTeamsContext = ((this.context) as any).microsoftTeams;
    if (microsoftTeamsContext) {
      retVal = new Promise((resolve, reject) => {
        microsoftTeamsContext.getContext(context => {
          this._teamsContext = context;
          resolve();
        });
      });
    }
    return retVal;
  }

  private _teamsContext: microsoftTeams.Context;

  public render(): void {
  
    let subTitle: string = '';
    if (this._teamsContext) {
      // We have teams context for the web part
      subTitle = `[Microsoft Teams] ${this._teamsContext.teamName} > ${this._teamsContext.channelName}`;
    }
    else
    {
      // We are rendered in normal SharePoint context
      subTitle = "[Microsoft SharePoint]";
    }

    const element: React.ReactElement<IMyDashboardProps > = React.createElement(
      MyDashboard,
      {
        title: this.properties.welcomeMessage,
        subTitle,
        graphClientFactory: this.context.msGraphClientFactory,
        teamsContext: this._teamsContext,
        showDebug: this.properties.showDebug
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
            description: "Settings"
          },
          groups: [
            {
              groupName: "basic settings",
              groupFields: [
                PropertyPaneTextField('welcomeMessage', {
                  label: "Welcome message"
                }),
                PropertyPaneToggle('showDebug', {
                  label: "Show debug"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
