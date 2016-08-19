import {
  EnvironmentType
} from '@microsoft/sp-client-base';

// SharePoint Framework Imports
import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-client-preview';

import importableModuleLoader from '@microsoft/sp-module-loader';

// App Imports
import * as strings from 'mystrings';
import MockHttpClient from './tests/MockHttpClient';
import { IJQueryWebPartProps } from './IJQueryWebPartProps';
import * as myjQuery from 'jquery';
import './JQueryWebPart.css';

require('jqueryui');

// Define List Models
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export default class JQueryWebPart extends BaseClientSideWebPart<IJQueryWebPartProps> {

  // Define and retrieve mock List data
  private _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get(this.context.pageContext.web.absoluteUrl).then(() => {
        const listData: ISPLists = {
            value:
            [
                { Title: 'Mock List 1', Id: '1' },
                { Title: 'Mock List 2', Id: '2' },
                { Title: 'Mock List 3', Id: '3' },
                { Title: 'Mock List 4', Id: '4' }
            ]
            };
        return listData;
    }) as Promise<ISPLists>;
  }

  // Retrieve Lists from SharePoint
  private _getListData(): Promise<ISPLists> {
    return this.context.httpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`)
      .then((response: Response) => {
      return response.json();
      });
  }

  // Call methods for List data retrieval
  private _renderListAsync(): void {
  // Mock List data
  if (this.context.environment.type === EnvironmentType.Local) {
    this._getMockListData().then((response) => {
      this._renderList(response.value);
    }); }
    else {
    this._getListData()
      .then((response) => {
        this._renderList(response.value);
      });
    }
  }

  // Render the List data with the values fetched from the REST API
  private _renderList(items: ISPList[]): void {
    // Remove accordion to handle property changes
    $('.accordion').remove();

    // Set up html for the jQuery UI Accordion Widget to display collapsible content panels
    // Learn more about the Accordion Widget at http://jqueryui.com/accordion/
    let html: string = '';

    html += `<div class='accordion'>`;

    items.forEach((item: ISPList) => {
        html += `
          <h3>${item.Title}</h3>
            <div>
                <p> ${item.Id} </p>
            </div>`;
    });

    this.domElement.innerHTML += html;

    html += `</div>`;

    // Set up base accordion options
    const accordionOptions: JQueryUI.AccordionOptions = {
      animate: this.properties.speed,
      collapsible: true,
      icons: {
        header: 'ui-icon-circle-arrow-e',
        activeHeader: 'ui-icon-circle-arrow-s'
      }
    };

    // Set up configurable jQueryUI effects and interactions
    if (this.properties.resize == false) {
      myjQuery(this.domElement).children('.accordion').accordion(accordionOptions);
    } else {
      myjQuery(this.domElement).children('.accordion').accordion(accordionOptions).resizable();
    }

    if (this.properties.sort == false) {
      myjQuery(this.domElement).children('.accordion').accordion(accordionOptions);
    } else {
      myjQuery(this.domElement).children('.accordion').accordion(accordionOptions).sortable();
    }
  }

  public constructor(context: IWebPartContext) {
    super(context);

    // Load remote stylesheet
    importableModuleLoader.loadCss('//code.jquery.com/ui/1.12.0/themes/base/jquery-ui.css');
  }

  // Render the jQuery Widget Web Part
  public render(): void {
    this._renderListAsync();
  }

  // Set up the Web Part Property Pane
  protected get propertyPaneSettings(): IPropertyPaneSettings {
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
                PropertyPaneSlider('speed', {
                  label: 'Animation Speed',
                  min: 1,
                  max: 500
                }),
                PropertyPaneToggle('resize', {
	                label: 'Resizable',
                  onText: 'Enable',
                  offText: 'Disable'
                }),
                PropertyPaneToggle('sort', {
                  label: 'Sortable',
                  onText: 'Enable',
                  offText: 'Disable'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
