import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';
import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
var fabric: any = require('fabric');

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  onInit(): Promise<any> {
    let cssURL = "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.2.0/css/fabric.min.css";
    let _cssURL = "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.2.0/css/fabric.components.min.css";
    SPComponentLoader.loadCss(cssURL);
    SPComponentLoader.loadCss(_cssURL);
    return Promise.resolve();
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.helloWorld}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle}">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description}">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button}">
                <span class="${ styles.label}">Learn more</span>
              </a>
              <div class="ms-CheckBox"> 
                  <input tabindex="-1" type="checkbox" class="ms-CheckBox-input">
                  <label role="checkbox"
                      class="ms-CheckBox-field"
                      tabindex="0"
                      aria-checked="false"
                      name="checkboxa">
                      <span class="ms-Label">Checkbox</span>
                  </label>
              </div>
            </div>
          </div>
        </div>
      </div>`;
    var CheckBoxElements = document.querySelectorAll(".ms-CheckBox");
    for (var i = 0; i < CheckBoxElements.length; i++) {
      new fabric['CheckBox'](CheckBoxElements[i]);
    }
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
