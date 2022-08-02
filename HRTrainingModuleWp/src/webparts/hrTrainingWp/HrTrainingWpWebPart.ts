import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'HrTrainingWpWebPartStrings';
import HrTrainingWp from './components/HrTrainingWp';
import { IHrTrainingWpProps } from './components/IHrTrainingWpProps';
import { sp } from "@pnp/sp/presets/all";
export interface IHrTrainingWpWebPartProps {
  description: string;
}

export default class HrTrainingWpWebPart extends BaseClientSideWebPart<IHrTrainingWpProps> {
  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      // other init code may be present  
      sp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    const element: React.ReactElement<IHrTrainingWpProps> = React.createElement(
      HrTrainingWp,
      {
        description: this.properties.description,
        context: this.context,
        siteUrl: this.context.pageContext.web.serverRelativeUrl,
        hrTrainingTitles: this.properties.hrTrainingTitles,
        hrTrainingReports: this.properties.hrTrainingReports,
        TrainingModuleLibrary: this.properties.TrainingModuleLibrary,
        webPartTitle: this.properties.webPartTitle,
        errorMessage: this.properties.errorMessage,
        statementIfNoItems: this.properties.statementIfNoItems,
        labelForInstructions: this.properties.labelForInstructions,
        messageBar: this.properties.messageBar
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
                }),
                PropertyPaneTextField('hrTrainingReports', {
                  label: "Hr Training Titles List Name"
                }),
                PropertyPaneTextField('hrTrainingTitles', {
                  label: "Hr Training Reports List Name"
                }),
                PropertyPaneTextField('TrainingModuleLibrary', {
                  label: "Training Module Library Name"
                }),
                PropertyPaneTextField('webPartTitle', {
                  label: "WebPart Title"
                }),
                PropertyPaneTextField('errorMessage', {
                  label: "Error Message"
                }),
                PropertyPaneTextField('statementIfNoItems', {
                  label: "If no items"
                }),
                PropertyPaneTextField('labelForInstructions', {
                  label: "Label For Instructions"
                }),
                PropertyPaneTextField('messageBar', {
                  label: "Message Bar After Submit"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
