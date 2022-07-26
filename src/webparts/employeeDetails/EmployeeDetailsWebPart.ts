import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './EmployeeDetailsWebPart.module.scss';
import * as strings from 'EmployeeDetailsWebPartStrings';

import { SPHttpClient, SPHttpClientResponse,
ISPHttpClientOptions } from '@microsoft/sp-http';


export interface IEmployeeDetailsWebPartProps {
  description: string;
}

export default class EmployeeDetailsWebPart extends BaseClientSideWebPart<IEmployeeDetailsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit()
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.employeeDetails} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      
      <div>
        <form>
          <label for="fname">Full Name:</label>
          <input type="text" id="fname" name="fname"></br></br>

          <label for="age">Age:</label>
          <input type="text" id="age" name="age"></br></br>

          <button type="button" id="bttnSave">Save</button>
        </form>
      </div>
    </section>`;
    this._bindSave();
  }

  private _bindSave(): void {
    this.domElement.querySelector('#bttnSave').addEventListener('click', () => {
      this.addListItem();
    });
  }

  private addListItem(): void {
    const fname = document.getElementById("fname")["value"];
    const age = document.getElementById("age")["value"];
    const siteUrl: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('EmployeeDetails')/items";
    const itemBody: any = {
      "Title": fname,
      "Age": age,
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      "body": JSON.stringify(itemBody)
    };

    this.context.spHttpClient.post(siteUrl, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        alert('success');
      });
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

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
