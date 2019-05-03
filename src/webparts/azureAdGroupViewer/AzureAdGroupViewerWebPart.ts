import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'AzureAdGroupViewerWebPartStrings';
import AzureAdGroupViewer from './components/AzureAdGroupViewer';
import { IAzureAdGroupViewerProps } from './components/IAzureAdGroupViewerProps';
import { setup as pnpSetup } from '@pnp/common';

export interface IAzureAdGroupViewerWebPartProps {
  description: string;
  groupName: string;
}

export default class AzureAdGroupViewerWebPart extends BaseClientSideWebPart<IAzureAdGroupViewerWebPartProps> {

  public onInit(): Promise<void> {

    pnpSetup({
      spfxContext: this.context
    });

    return Promise.resolve();
  }
  public render(): void {
    const element: React.ReactElement<IAzureAdGroupViewerProps > = React.createElement(
      AzureAdGroupViewer,
      {
        description: this.properties.description,
        groupName:this.properties.groupName
        
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
