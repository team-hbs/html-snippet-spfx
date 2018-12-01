import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'HtmlSnippetSpfxWebPartStrings';
import HtmlSnippetSpfx from './components/HtmlSnippetSpfx';
import { IHtmlSnippetSpfxProps } from './components/IHtmlSnippetSpfxProps';

export interface IHtmlSnippetSpfxWebPartProps {
  description: string;
}

export default class HtmlSnippetSpfxWebPart extends BaseClientSideWebPart<IHtmlSnippetSpfxWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IHtmlSnippetSpfxProps > = React.createElement(
      HtmlSnippetSpfx,
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
