import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration, PropertyPaneToggle, PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SlickWebPartStrings';
import Slick from './components/Slick';
import { ISlickProps } from './components/ISlickProps';

export interface ISlickWebPartProps {
  description: string;
  sliderSelection:number;
}

export default class SlickWebPart extends BaseClientSideWebPart<ISlickWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISlickProps> = React.createElement(
      Slick,
      {
        description: this.properties.description,
        level: this.properties.sliderSelection
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

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
            description: strings.FirstPageDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneToggle("showAdvancedSettings", {
                  label: strings.Controls.AdvancedToggle.Label,
                  onText: strings.Controls.AdvancedToggle.OnText,
                  offText: strings.Controls.AdvancedToggle.OffText,
                }),
              ]
            }
          ]
        },
        {
          header: {
            description: strings.SecondPageDescription
          },
          groups: [
            {
              groupName: strings.AdvancedGroupName,
              groupFields: [
                PropertyPaneSlider("sliderSelection", {
                  label: strings.Controls.Slider.Label,
                  min: 0,
                  max: 100,
                  value: 50,
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
