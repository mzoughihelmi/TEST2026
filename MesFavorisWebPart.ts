import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "MesFavorisWebPartStrings";
import { App, IAppProps } from "./components/App/App";
import { IMesFavorisWebPartProps } from "./models/IMesFavorisWebPartProps";
import { DEFAULT_FAVORITES_LIST_TITLE } from "./services/FavoritesService";

export default class MesFavorisWebPart extends BaseClientSideWebPart<IMesFavorisWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IAppProps> = React.createElement(App, {
      context: this.context,
      // Emplacement choisi dans le volet ; "famille" par défaut.
      displayMode: this.properties.displayMode === "accueil" ? "accueil" : "famille",
      themeVariant: undefined,
      listTitle: DEFAULT_FAVORITES_LIST_TITLE,
      title: this.properties.title,
      maxItems: this.properties.maxItems || 5,
      voirPlusUrl: this.properties.voirPlusUrl,
      voirPlusLabel: this.properties.voirPlusLabel
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneChoiceGroup("displayMode", {
                  label: strings.DisplayModeFieldLabel,
                  options: [
                    {
                      key: "accueil",
                      text: strings.DisplayModeAccueil,
                      iconProps: { officeFabricIconFontName: "Home" }
                    },
                    {
                      key: "famille",
                      text: strings.DisplayModeFamille,
                      iconProps: { officeFabricIconFontName: "Product" }
                    }
                  ]
                }),
                PropertyPaneTextField("title", {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneSlider("maxItems", {
                  label: strings.MaxItemsFieldLabel,
                  min: 1,
                  max: 10,
                  value: 5
                }),
                PropertyPaneTextField("voirPlusUrl", {
                  label: strings.VoirPlusUrlFieldLabel,
                  description: strings.VoirPlusUrlFieldDescription
                }),
                PropertyPaneTextField("voirPlusLabel", {
                  label: strings.VoirPlusLabelFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
