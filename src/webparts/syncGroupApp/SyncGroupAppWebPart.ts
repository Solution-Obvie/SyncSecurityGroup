import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SyncGroupAppWebPartStrings';
import SyncGroupApp from './components/SyncGroupApp';
import { ISyncGroupAppProps } from './components/ISyncGroupAppProps';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { sp } from "@pnp/sp/presets/all";

export interface ISyncGroupAppWebPartProps {
  description: string;
}

export default class SyncGroupAppWebPart extends BaseClientSideWebPart<ISyncGroupAppWebPartProps> {

  protected onInit(): Promise<void> {

    return super.onInit().then(_ => {
  
      // other init code may be present
  
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    // this.context.msGraphClientFactory
    // .getClient()
    // .then((client: MSGraphClient): void => {
    //   // get information about the current user from the Microsoft Graph
    //   client
    //     .api('/me')
    //     .get((error, response: any, rawResponse?: any) => {
    //       // handle the response
    //       console.log(JSON.stringify(response));
    //       console.log(error)
    //       console.log(rawResponse)
    //     })
    //   })
    //   this.context.msGraphClientFactory
    //   .getClient()
    //   .then((client: MSGraphClient) => {
    //     // get information about the current user from the Microsoft Graph
    //     client
    //     //  .api("/groups?$filter=groupTypes/any(c:c eq ' ') ") // and mailEnabled eq 'false'
    //     .api("/groups?$filter=mailEnabled eq 'false' ")
    //       .get((error, response: any, rawResponse?: any) => {
    //         // handle the response
    //         console.log(JSON.stringify(response));
    //         console.log(error)
    //         console.log(rawResponse)
    //       })
    //     })
    const element: React.ReactElement<ISyncGroupAppProps> = React.createElement(
      SyncGroupApp,
      {
        description: this.properties.description,
        context: this.context
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
