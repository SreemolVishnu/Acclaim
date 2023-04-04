import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'NboDetailListWebPartStrings';
import NboDetailList from './components/NboDetailList';
import { INboDetailListProps } from './components/INboDetailListProps';
import { sp } from "@pnp/sp/presets/all";
import $ from 'jquery';
export interface INboDetailListWebPartProps {
  description: string;
}

export default class NboDetailListWebPart extends BaseClientSideWebPart<INboDetailListProps> {
  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      // other init code may be present
      sp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    try {
      $(".ControlZone").parent().parent().css("max-width", "100%");
    }
    catch (err) {
      console.log("Couldnot update the max-width of the page");
    }
    const element: React.ReactElement<INboDetailListProps> = React.createElement(
      NboDetailList,
      {
        description: this.properties.description,
        context: this.context,
        siteUrl: this.context.pageContext.web.serverRelativeUrl,
        sourceListName: this.properties.sourceListName,
        nboStage: this.properties.nboStage,
        nboListName: this.properties.nboListName,
        classOfInsurance: this.properties.classOfInsurance,
        industry: this.properties.industry,
        brokeragePercentage: this.properties.brokeragePercentage,
        teamList: this.properties.teamList,
        pageSizeForPagination: this.properties.pageSizeForPagination,
        emailNotificationSettings: this.properties.emailNotificationSettings,
        complianceGroupEmail: this.properties.complianceGroupEmail,
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
                PropertyPaneTextField('sourceListName', {
                  label: "Source List Name"
                }),
                PropertyPaneTextField('nboListName', {
                  label: "NBO List Name"
                }),
                PropertyPaneTextField('classOfInsurance', {
                  label: "Class Of Insurance List Name"
                }),
                PropertyPaneTextField('industry', {
                  label: "Industry List Name"
                }),
                PropertyPaneTextField('nboStage', {
                  label: "NBO Stage List Name"
                }),
                PropertyPaneTextField('brokeragePercentage', {
                  label: "Brokerage Percentage List Name"
                }),
                PropertyPaneTextField('teamList', {
                  label: "Team List Name"
                }),
                PropertyPaneTextField('pageSizeForPagination', {
                  label: "Page Size For Pagination"
                }),
                PropertyPaneTextField('emailNotificationSettings', {
                  label: "Email Notification Settings List"
                }),
                PropertyPaneTextField('complianceGroupEmail', {
                  label: "Compliance Group Email"
                })

              ]
            }
          ]
        }
      ]
    };
  }
}
