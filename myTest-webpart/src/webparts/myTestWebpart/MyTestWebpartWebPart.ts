import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MyTestWebpartWebPart.module.scss';
import * as strings from 'MyTestWebpartWebPartStrings';
import { PropertyPaneDropdown } from '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneDropdown/PropertyPaneDropdown';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IMyTestWebpartWebPartProps {
  description: string;
  listTitle: string;
  formHTML: string;
}

export interface spList {
  Title: string;
  id: string;
}
export interface spLists {
  value: spList[];
}

export default class MyTestWebpartWebPart extends BaseClientSideWebPart<IMyTestWebpartWebPartProps> {
  private dropDownOptions: IPropertyPaneDropdownOption[] = [];

  public render(): void {
    this.getAllFieldsFromList().then((response) => {
      this.getFieldsArray(response.value);
    })
    this.domElement.innerHTML = `
      <div class="${ styles.myTestWebpart}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">${escape(this.properties.listTitle)} - Form</span>
              <p class="${ styles.description}">${escape(this.properties.description)}</p>
              <p class="${ styles.description}">${this.properties.formHTML}</p>              
            </div>
          </div>
        </div>
      </div>`;
  }

  private getAllFieldsFromList(): Promise<any> {
    if (this.properties.listTitle != '' && this.properties.listTitle != undefined) {
      let listFieldUrl: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('" + this.properties.listTitle + "')/Fields?$filter=Hidden eq false and Group ne '_Hidden' and TypeAsString ne 'Computed' and ReadOnlyField eq false";
      return this.context.spHttpClient.get(listFieldUrl, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          return response.json();
        })
        .catch((error) => {
          console.error("Error Occurred while loading Fields", error);
        })
    }
  }

  allFields: Array<any>;
  private getFieldsArray(fieldArray: Array<JSON>) {
    console.info("Data:", fieldArray);
    this.allFields = new Array<any>();
    this.properties.formHTML = "";
    fieldArray.forEach(field => {
      this.allFields.push({
        Title: field["Title"],
        InternalName: field["InternalName"],
        Type: field["TypeAsString"]
      });
      if (field["TypeAsString"] === "Text") {
        this.properties.formHTML += `
        <p class="${ styles.description}">
          <label for='${field["Id"]}${field["InternalName"]}'>${field["Title"]}</label>
          <input type='text' id='${field["Id"]}${field["InternalName"]}' name='${field["Id"]}${field["InternalName"]}'/>
        </p>
        `;
      }
    });

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
            },
            {
              groupName: "Lists",
              groupFields: [
                PropertyPaneDropdown('listTitle', {
                  label: "Select List",
                  options: this.dropDownOptions,
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected onPropertyPaneConfigurationStart(): void {
    if (this.dropDownOptions.length > 0) {
      return;
    }

    this.GetLists();
  }

  protected onPropertyPaneFieldChanged(): void {
    this.render();
  }
  protected get disableReactivePropertyChanges():boolean{
    return true;
  }
  private GetLists(): void {
    let listRestUrl: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists?$select=Id,Title&$filter=Hidden eq false";
    this.LoadLists(listRestUrl)
      .then((response) => {
        this.LoadDropDownValues(response.value);
      })
      .catch((error) => {
        console.error("Error Occurred while loading Lists", error);
      })
  }

  /**
   * @name LoadLists
   * @description Load list from sharepoint site
   * @param listresturl
   * @returns Promise 
   */
  private LoadLists(listresturl: string): Promise<spLists> {
    return this.context.spHttpClient.get(listresturl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
      return response.json();
    });
  }

  /**
   * @name LoadDropDownValues
   * @description Load List names from sharepoint site
   * @param lists
   */
  private LoadDropDownValues(lists: spList[]): void {
    lists.forEach((list: spList) => {
      // Loads the drop down values  
      this.dropDownOptions.push({ key: list.Title, text: list.Title });
    });
    this.context.propertyPane.refresh();
  }
}
