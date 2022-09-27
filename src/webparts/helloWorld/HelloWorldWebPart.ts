import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as pnp from '@pnp/sp'
import { Environment, EnvironmentType

} from '@microsoft/sp-core-library';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import styles from ./HelloWorld.module.scss;
import * as strings from 'HelloWorldWebPartStrings';
import HelloWorld from './components/HelloWorld';
import { IHelloWorldProps } from './components/IHelloWorldProps';
import { head } from 'lodash';

export interface IHelloWorldWebPartProps {
  description: string;
}

export interface ISPList {
  Description: string;
  Priority: string;
  Status: string;
  Assignedto: string; 
DateReported:string;
IssueSource:string;
Images:any;
Issueloggedby:string;
lookupcol:string;
IsActive:boolean
}
export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <div class="${styles.helloWorld}">
    <div class="${styles.container}">
    <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
    <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
      <span class="ms-font-xl ms-fontColor-white" style="font-size:28px">Welcome to SharePoint Framework Development using PnP JS Library</span>
     
      <p class="ms-font-l ms-fontColor-white" style="text-align: left">Demo : Retrieve Employee Data from SharePoint List</p>
    </div>
    </div>
    <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
      <div style="background-color:Black;color:white;text-align: center;font-weight: bold;font-size:18px;">Employee Details</div>
    <br>
    <div id="spListContainer" />
    </div>
    </div> 94. </div>`;
    this._renderListAsync();
    
     } 

  private _getListData(): Promise<ISPList[]> {
    return pnp.sp.web.lists.getByTitle("Sample List").items.get().then((response) => {
    return response;
    
     });
    
    
    }

    private _renderListAsync(): void { 
      this._getListData()
      .then((response) => {
      this._renderList(response);
   });
      
    
      }

      private _renderList(items: ISPList[]): void {

         let html: string = '<table class="TFtable" border=1 width=100% style="border-collapse: collapse;">';
        
        html +=
        
        `<th>Description</th><th>Priority</th><th>Status</th><th>DateReported</th><th>IsActive</th><th>Assignedto</th><th>Issueloggedby</th><th>Images</th>`;
        
   items.forEach((item: ISPList) => {
        
        html += `
        <tr>
        <td>${item.Description?item.Description:""}</td>
        <td>${item.Priority?item.Priority:""}</td>
        <td>${item.Status?item.Status:""}</td>
        <td>${item.DateReported?item.DateReported:""}</td>
        <td>${item.IsActive?item.IsActive:""}</td>
        <td>${item.Assignedto?item.Assignedto:""}</td>
        <td>${item.Issueloggedby?item.Issueloggedby:""}</td>
        <td>${item.Images?item.Images:""}</td>

        </tr>
        
        `;
        
        });
        
        html += `</table>`;
        const listContainer: Element = this.domElement.querySelector('#spListContainer');
        listContainer.innerHTML = html;
         } 

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
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
