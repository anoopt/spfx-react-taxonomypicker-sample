import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SimpleTaxonomyPickerWebPartStrings';
import SimpleTaxonomyPicker from './components/SimpleTaxonomyPicker';
import { ISimpleTaxonomyPickerProps } from './components/ISimpleTaxonomyPickerProps';

export interface ISimpleTaxonomyPickerWebPartProps {
  description: string;
}

export default class SimpleTaxonomyPickerWebPart extends BaseClientSideWebPart<ISimpleTaxonomyPickerWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISimpleTaxonomyPickerProps > = React.createElement(
      SimpleTaxonomyPicker,
      {
        description: this.properties.description,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
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
