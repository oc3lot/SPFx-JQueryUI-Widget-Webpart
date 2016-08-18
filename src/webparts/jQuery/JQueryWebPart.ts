import {
  DisplayMode
} from '@ms/sp-client-base';

// SharePoint Framework imports
import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  IWebPartData,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle,
  HostType
} from '@ms/sp-client-platform';

// App imports
import strings from './loc/Strings.resx';
import HttpClient from './tests/HttpClient';
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

// Web Part properties
export interface IJQueryWebPartProps {
  description: string;
  speed: number;
  resize: boolean;
  sort: boolean;
}

export default class JQueryWebPart extends BaseClientSideWebPart<IJQueryWebPartProps> {

  // Define and retrieve mock List data
  private _getMockListData(): Promise<ISPLists> {
    return HttpClient.get(this.host.pageContext.webAbsoluteUrl).then(() => {
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

  // Retrieve List data from SharePoint
  private _getListData(): Promise<ISPLists> {
    return this.host.httpClient.get(this.host.pageContext.webAbsoluteUrl + `/_api/web/lists?$filter=Hidden eq false`)
      .then((response: Response) => {
      return response.json();
    });
  }

  // Call methods for List data retrieval
  private _renderListAsync(): void {
    // Mock List data
    if (this.host.hostType === HostType.TestPage) {
        this._getMockListData().then((response) => {
            this._renderList(response.value);
        });

    // SharePoint List data on Modern Page
    } else if (this.host.hostType === HostType.ModernPage) {
        this._getListData()
            .then((response) => {
                this._renderList(response.value);
            });

    // SharePoint List data on Classic Page
    } else if (this.host.hostType == HostType.ClassicPage) {
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
    this.host.resourceLoader.loadCSS('//code.jquery.com/ui/1.12.0/themes/base/jquery-ui.css');
  }

  // Render the jQuery Widget Web Part
  public render(mode: DisplayMode, data?: IWebPartData): void {
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
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}