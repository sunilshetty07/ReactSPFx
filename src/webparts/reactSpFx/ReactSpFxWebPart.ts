import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ReactSpFxWebPartStrings';
import ReactSpFx from './components/ReactSpFx';
import { IReactSpFxProps } from './components/IReactSpFxProps';
import { LogLevel, PnPLogging } from "@pnp/logging";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { SPHttpClient } from "@microsoft/sp-http";

export interface IReactSpFxWebPartProps {
  selectedList: string;
  description: string;
}
export interface IMyWebPartProps {
  selectedList: string; // holds selected list name
}

let _sp: SPFI | undefined;



export const getSP = (context: WebPartContext): SPFI => {
  if (context) {
     //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    _sp = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
  }
  return _sp!;
};


export default class ReactSpFxWebPart extends BaseClientSideWebPart<IReactSpFxWebPartProps> {
  private listOptions: IPropertyPaneDropdownOption[] = [];
  
  public render(): void {
    const element: React.ReactElement<IReactSpFxProps> = React.createElement(
      ReactSpFx,
      {
        userDisplayName: this.context.pageContext.user.displayName,
        context:this.context,
        selectedList: this.properties.selectedList
      }
    );
    
    
    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
  return this._getEnvironmentMessage().then(async (message) => {
    // 1. Fetch all lists in the site
    const listsResp = await this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Title&$filter=Hidden eq false`,
      SPHttpClient.configurations.v1
    );
    const listsJson = await listsResp.json();

    // 2. Map to property pane dropdown options
    this.listOptions = listsJson.value.map((list: any) => {
      return { key: list.Title, text: list.Title };
    });

    // 3. Refresh property pane so dropdown updates
    this.context.propertyPane.refresh();
  });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    //this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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
              // Existing text field
              PropertyPaneTextField('description', {
                label: strings.DescriptionFieldLabel
              }),

              // New dropdown field for selecting a list
              PropertyPaneDropdown('selectedList', {
                label: "Choose a SharePoint List",
                options: this.listOptions.length > 0 ? this.listOptions : [
                  { key: "loading", text: "Loading..." }
                ]
              })
            ]
          }
        ]
      }
    ]
  };
}


protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
  super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

  if (propertyPath === 'selectedList' && oldValue !== newValue) {
    this.render();
  }
}


}
