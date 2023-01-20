import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'CarouselWithSharepointWebPartStrings';
import CarouselWithSharepoint from './components/CarouselWithSharepoint';
import { ICarouselWithSharepointProps } from './components/ICarouselWithSharepointProps';

export interface ICarouselWithSharepointWebPartProps {
 description : string;
 NumberOfNews : number;
 NumberOfIntervelInSeconds : number;
  StopScroll : boolean;
  listName : string;
  absoluteURL : any;
  spHttpClient : any;
}

export default class CarouselWithSharepointWebPart extends BaseClientSideWebPart<ICarouselWithSharepointWebPartProps> {


  public render(): void {
    const element: React.ReactElement<ICarouselWithSharepointProps> = React.createElement(
      CarouselWithSharepoint,
      {
        description: this.properties.description,
        NumberOfNews: this.properties.NumberOfNews,
        NumberOfIntervelInSeconds : this.properties.NumberOfIntervelInSeconds,
        StopScroll : this.properties.StopScroll,
        listName: this.properties.listName,
        absoluteURL: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
      }
      
    );
      
    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(() => {
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    // this._isDarkTheme = !!currentTheme.isInverted;
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
            description: "Add list Name"
          },
          groups: [
            {
              groupName: "News",
              groupFields: [
                PropertyPaneTextField('NumberOfNews', {
                  label: "Please add the number of Posts to Show"
                }),
                PropertyPaneToggle("StopScroll", {
                  label: "Disable Automatically cycle news",
                  onText: "on",
                  offText: "off"
                }),
                PropertyPaneTextField('NumberOfIntervelInSeconds', {
                  label: "Second between each change of news posts"
                }),

              ]
            }
          ],
          
        }
      ]
    };
  }
}
