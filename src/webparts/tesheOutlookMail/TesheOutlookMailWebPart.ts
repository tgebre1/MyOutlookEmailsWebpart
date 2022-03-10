import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { PropertyFieldIconPicker } from "@pnp/spfx-property-controls/lib/PropertyFieldIconPicker";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import * as strings from "TesheOutlookMailWebPartStrings";
import { TesheOutlookMail, ITesheOutlookMailProps } from "./components";
import { CalloutTriggers } from "@pnp/spfx-property-controls/lib/PropertyFieldHeader";
import { PropertyFieldToggleWithCallout } from "@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout";
import { MSGraphClient } from "@microsoft/sp-http";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { loadTheme } from "office-ui-fabric-react";

import {
  IReadonlyTheme,
  ThemeChangedEventArgs,
  ThemeProvider,
} from "@microsoft/sp-component-base";

export interface ITesheOutlookMailWebPartProps {
  title: string;
  nrOfMessages: number;
  showInboxOnly: boolean;
  iconPicker: string;
}

export default class TesheOutlookMailWebPart extends BaseClientSideWebPart<ITesheOutlookMailWebPartProps> {
  private graphClient: MSGraphClient;
  private propertyFieldNumber;
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  public onInit(): Promise<void> {
    return new Promise<void>(
      (resolve: () => void, reject: (error: any) => void): void => {
        this.context.msGraphClientFactory.getClient().then(
          (client: MSGraphClient): void => {
            this.graphClient = client;
            resolve();
          },
          (err) => reject(err)
        );

        this._themeProvider = this.context.serviceScope.consume(
          ThemeProvider.serviceKey
        );
        // If it exists, get the theme variant
        this._themeVariant = this._themeProvider.tryGetTheme();
        // Register a handler to be notified if the theme variant changes
        this._themeProvider.themeChangedEvent.add(
          this,
          this._handleThemeChangedEvent
        );
      }
    );
  }

  public render(): void {
    const element: React.ReactElement<ITesheOutlookMailProps> =
      React.createElement(TesheOutlookMail, {
        title: this.properties.title,
        nrOfMessages: this.properties.nrOfMessages,
        showInboxOnly: this.properties.showInboxOnly,
        // pass the current display mode to determine if the title should be
        // editable or not
        displayMode: this.displayMode,
        themeVariant: this._themeVariant,
        iconPicker: this.properties.iconPicker,
        // pass the reference to the MSGraphClient
        graphClient: this.graphClient,
        // handle updated web part title
        updateProperty: (value: string): void => {
          // store the new title in the title web part property
          this.properties.title = value;
        },
      });

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  //executes only before property pane is loaded.
  protected async loadPropertyPaneResources(): Promise<void> {
    // import additional controls/components

    const { PropertyFieldNumber } = await import(
      /* webpackChunkName: 'pnp-propcontrols-number' */
      "@pnp/spfx-property-controls/lib/propertyFields/number"
    );

    this.propertyFieldNumber = PropertyFieldNumber;
  }
  protected get disableReactivePropertyChanges(): boolean {
    return true;
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
              groupFields: [
                this.propertyFieldNumber("nrOfMessages", {
                  key: "nrOfMessages",
                  label: strings.NrOfMessagesToShow,
                  value: this.properties.nrOfMessages,
                  minValue: 1,
                  maxValue: 20,
                }),
                PropertyFieldToggleWithCallout("showInboxOnly", {
                  calloutTrigger: CalloutTriggers.Click,
                  key: "showInboxOnly",
                  label: strings.ShowInboxOnly,
                  calloutContent: React.createElement(
                    "p",
                    {},
                    strings.ShowInboxOnlyCallout
                  ),
                  checked: this.properties.showInboxOnly,
                }),
                PropertyFieldIconPicker("iconPicker", {
                  currentIcon: this.properties.iconPicker,
                  key: "iconPickerId",
                  onSave: (icon: string) => {
                    console.log(icon);
                    this.properties.iconPicker = icon;
                  },
                  onChanged: (icon: string) => {
                    console.log(icon);
                  },
                  buttonLabel: "Icon",
                  renderOption: "panel",
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  label: "Icon Picker",
                }),
              ],
            },
          ],
        },
      ],
    };
  }

  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;

    this.render();
  }
}
