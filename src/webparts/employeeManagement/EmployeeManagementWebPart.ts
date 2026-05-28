import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'EmployeeManagementWebPartStrings';
import { EmployeeManagementPanel } from './components/EmployeeManagementPanel';
import { IEmployeeManagementProps } from './components/IEmployeeManagementProps';

export interface IEmployeeManagementWebPartProps {
  description: string;
  apiBaseUrl: string;
}

export default class EmployeeManagementWebPart extends BaseClientSideWebPart<IEmployeeManagementWebPartProps> {
  public onInit(): Promise<void> {
    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<IEmployeeManagementProps> = React.createElement(
      EmployeeManagementPanel,
      {
        apiBaseUrl: this.properties.apiBaseUrl,
        userEmail: this.context.pageContext.user.email,
        userName: this.context.pageContext.user.displayName,
        webUrl: this.context.pageContext.web.absoluteUrl,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('apiBaseUrl', {
                  label: 'API Base URL',
                  placeholder: 'https://your-api.com/api',
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
