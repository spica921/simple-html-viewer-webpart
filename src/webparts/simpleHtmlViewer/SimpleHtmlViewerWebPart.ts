import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SimpleHtmlViewerWebPartStrings';
import SimpleHtmlViewer from './components/SimpleHtmlViewer';
import { ISimpleHtmlViewerProps } from './components/ISimpleHtmlViewerProps';

export interface ISimpleHtmlViewerWebPartProps {
  html: string;
}

export default class SimpleHtmlViewerWebPart extends BaseClientSideWebPart<ISimpleHtmlViewerWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISimpleHtmlViewerProps> = React.createElement(
      SimpleHtmlViewer,
      {
        html: this.properties.html
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
                PropertyPaneTextField('html', {
                  label: strings.HTMLLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
