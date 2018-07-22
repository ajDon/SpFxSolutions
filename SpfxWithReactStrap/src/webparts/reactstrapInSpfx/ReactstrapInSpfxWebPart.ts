import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import * as jquery from "jquery";
require('bootstrap/dist/js/bootstrap.js');
import { SPComponentLoader } from '@microsoft/sp-loader';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ReactstrapInSpfxWebPartStrings';
import ReactstrapInSpfx from './components/ReactstrapInSpfx';
import { IReactstrapInSpfxProps } from './components/IReactstrapInSpfxProps';

export interface IReactstrapInSpfxWebPartProps {
  description: string;
}

export default class ReactstrapInSpfxWebPart extends BaseClientSideWebPart<IReactstrapInSpfxWebPartProps> {

  public render(): void {
    // SPComponentLoader.loadScript(modalJs);
    const element: React.ReactElement<IReactstrapInSpfxProps> = React.createElement(
      ReactstrapInSpfx,
      {
        description: this.properties.description,
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl
      }
    );
    

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    // ReactDom.unmountComponentAtNode(this.domElement);
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
