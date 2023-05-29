import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GetSpListItemsWebPart.module.scss';
// import * as strings from 'GetSpListItemsWebPartStrings';
import MockHttpClient from './MockHttpClient';
import {  
  SPHttpClient  
} from '@microsoft/sp-http';
import {  
  Environment,  
  EnvironmentType  
} from '@microsoft/sp-core-library';
import { AppLocalEnvironmentSharePoint, AppLocalEnvironmentTeams, AppSharePointEnvironment, AppTeamsTabEnvironment, BasicGroupName, DescriptionFieldLabel, PropertyPaneDescription } from 'GetSpListItemsWebPartStrings';



export interface ISPLists {  
  value: ISPList[];  
}  
export interface ISPList {  
  Title: string;  
  Descripcion: string;  
  Estado: string;  
  codigo: string;  
  Fecha: Date;  
} 

export interface IGetSpListItemsWebPartProps {
  description: string;
}

export default class GetSpListItemsWebPart extends BaseClientSideWebPart<IGetSpListItemsWebPartProps> {

  private _getMockListData(): Promise<ISPLists> {  
    return MockHttpClient.get(this.context.pageContext.web.absoluteUrl).then(() => {  
        const listData: ISPLists = {  
            value:  
            [  
                { Title: 'Prueba 1', Descripcion: 'John', codigo: 'SharePoint',Estado: 'India' , Fecha: new Date("2023-05-25")},  
                 { Title: 'Prueba 2', Descripcion: 'Martin', codigo: '.NET',Estado: 'Qatar' , Fecha: new Date("2023-05-25")},  
                { Title: 'Prueba 3', Descripcion: 'Luke', codigo: 'JAVA',Estado: 'UK', Fecha: new Date("2023-05-25") }  
            ]  
            };  
        return listData;  
    }) as Promise<ISPLists>;  
  }


  private _getListData(): Promise<ISPLists> {  
    console.log("absoluteUrl: ", this.context.pageContext.web.absoluteUrl);
    
    return this.context.spHttpClient.get(`https://secol.sharepoint.com/sites/AppTareasSharepoint/_api/web/lists/GetByTitle('Tareas')/Items`, SPHttpClient.configurations.v1)  
        .then((response) => {   
          debugger;  
          return response.json();  
        });  
  }

  private _renderListAsync(): void {  
      
    if (Environment.type === EnvironmentType.Local) {  
      this._getMockListData().then((response) => {  
        this._renderList(response.value);  
      });  
    }  
    else {  
       this._getListData()  
      .then((response) => {  
        this._renderList(response.value);  
      });  
    }  
  }

  private _renderList(items: ISPList[]): void {  
    let html: string = '<table class="TFtable" border=1 width=100% style="border-collapse: collapse;">';  
    html += `<th>Nombre</th><th>Descripción</th><th>Código</th><th>Estado</th><th>Fecha</th>`;  
    items.forEach((item: ISPList) => {  
      html += `  
           <tr>  
          <td>${item.Title}</td>  
          <td>${item.Descripcion}</td>  
          <td>${item.codigo}</td>  
          <td>${item.Estado}</td>
          <td>${item.Fecha}</td>
          </tr>  
          `;  
    });  
    html += `</table>`;  
    const listContainer = this.domElement.querySelector('#spListContainer');  
    listContainer.innerHTML = html;  
  }   

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';


  public render(): void {  
    this.domElement.innerHTML = `  
      <div class="">  
  <div class="">  
    <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white">  
      <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">  
        <span class="ms-font-xl ms-fontColor-white" style="font-size:28px">Welcome to SharePoint Framework Development</span>  
          
        <p class="ms-font-l ms-fontColor-white" style="text-align: center">Demo : Retrieve Employee Data from SharePoint List</p>  
      </div>  
    </div>  
    <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white">  
    <div style="background-color:Black;color:white;text-align: center;font-weight: bold;font-size:18px;">Employee Details</div>  
    <br>  
  <div id="spListContainer" />  
    </div>  
  </div>  
  </div>`;  
  this._renderListAsync();  
 }  

  // public render(): void {
  //   this.domElement.innerHTML = `
  //   <section class="${styles.getSpListItems} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
  //     <div class="${styles.welcome}">
  //       <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
  //       <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
  //       <div>${this._environmentMessage}</div>
  //       <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
  //     </div>
  //     <div>
  //       <h3>Welcome to SharePoint Framework!</h3>
  //       <p>
  //       The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
  //       </p>
  //       <h4>Learn more about SPFx development:</h4>
  //         <ul class="${styles.links}">
  //           <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
  //           <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
  //           <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
  //           <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
  //           <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
  //           <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
  //           <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
  //         </ul>
  //     </div>
  //   </section>`;
  // }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }



  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? AppLocalEnvironmentTeams : AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? AppLocalEnvironmentSharePoint : AppSharePointEnvironment;
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

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: PropertyPaneDescription
          },
          groups: [
            {
              groupName: BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
