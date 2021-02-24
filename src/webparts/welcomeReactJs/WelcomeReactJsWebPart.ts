import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";

import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
	PropertyPaneToggle,
	PropertyPaneTextField,
	IPropertyPaneConfiguration,
	PropertyPaneChoiceGroup,
	PropertyPaneSlider,
  } from "@microsoft/sp-property-pane";
import * as strings from "WelcomeReactJsWebPartStrings";
import WelcomeReactJs from "./components/WelcomeReactJs";
import { IWelcomeReactJsProps } from "./components/IWelcomeReactJsProps";
import {
  ThemeProvider,
  ThemeChangedEventArgs,
  IReadonlyTheme,
} from "@microsoft/sp-component-base";


export interface IWelcomeReactJsWebPartProps {
  title: string;
  messagestyle: string;
  textalignment: string;
  showtimebasedmessage: boolean;
  morningmessage: string;
  afternoonmessage: string;
  afternoonbegintime: number;
  eveningmessage: string;
  eveningbegintime: number;
  message: string;
  showname: string;
  showfirstname: boolean;
}

export default class WelcomeReactJsWebPart extends BaseClientSideWebPart<IWelcomeReactJsWebPartProps> {
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;



  protected onInit(): Promise<void> {
    return super.onInit().then((_) => {
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
    });
  }

  /**
   * Update the current theme variant reference and re-render.
   *
   * @param args The new theme
   */
  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }

  protected get disableReactivePropertyChanges(): boolean {
    //return true;
    return false; // Use true to show Apply button
  }

  public render(): void {
    const element: React.ReactElement<IWelcomeReactJsProps> = React.createElement(
      WelcomeReactJs,
      {
        title: this.properties.title,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
        message: this.properties.message,
        themeVariant: this._themeVariant,
        showname: this.properties.showname,
        context: this.context,
        showtimebasedmessage: this.properties.showtimebasedmessage,
        morningmessage: this.properties.morningmessage,
        afternoonmessage: this.properties.afternoonmessage,
        eveningmessage: this.properties.eveningmessage,
        textalignment: this.properties.textalignment,
        messagestyle: this.properties.messagestyle,
        afternoonbegintime: this.properties.afternoonbegintime,
        eveningbegintime: this.properties.eveningbegintime,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let messagefields = [];

    if (this.properties.showtimebasedmessage) {
      messagefields.push(
        PropertyPaneTextField("morningmessage", {
          label: strings.MorningMessageLabel,
        })
      );
      messagefields.push(
        PropertyPaneTextField("afternoonmessage", {
          label: strings.AfternoonMessageLabel,
        })
      );

      messagefields.push(
        PropertyPaneTextField("eveningmessage", {
          label: strings.EveningMessageLabel,
        })
      );
      messagefields.push(
        PropertyPaneSlider("afternoonbegintime", {
          label: strings.AfternoonBeginTimeLabel,
          min: 11,
          max: 14,
        })
      );
      messagefields.push(
        PropertyPaneSlider("eveningbegintime", {
          label: strings.EveningBeginTimeLabel,
          min: 16,
          max: 19,
        })
      );
    } else {
      messagefields.push(
        PropertyPaneTextField("message", {
          label: strings.MessageLabel,
        })
      );
    }

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.NamePropertiesGroupName,
              groupFields: [
                PropertyPaneChoiceGroup("showname", {
                  label: strings.ShowNameLabel,
                  options: [
                    {
                      key: "full",
                      text: "Full name",
                    },
                    {
                      key: "first",
                      text: "First name only",
                    },
                    {
                      key: "none",
                      text: "No name",
                    },
                  ],
                }),
                PropertyPaneToggle("showtimebasedmessage", {
                  label: strings.ShowTimeBasedMessageLabel,
                }),
                ...messagefields,
              ],
            },
            {
              groupName: strings.StylePropertiesGroupName,
              groupFields: [
                PropertyPaneChoiceGroup("messagestyle", {
                  label: strings.MessageStyleLabel,
                  options: [
                    {
                      key: "h1",
                      text: "H1",
                    },
                    {
                      key: "h2",
                      text: "H2",
                    },
                    {
                      key: "h3",
                      text: "H3",
                    },
                    {
                      key: "h4",
                      text: "H4",
                    },
                    {
                      key: "p",
                      text: "P",
                    },
                  ],
                }),
                PropertyPaneChoiceGroup("textalignment", {
                  label: strings.TextAlignmentLabel,
                  options: [
                    {
                      key: "left",
                      text: "Left",
                      iconProps: {
                        officeFabricIconFontName: "AlignLeft",
                      },
                    },
                    {
                      key: "centre",
                      text: "Center",
                      iconProps: {
                        officeFabricIconFontName: "AlignCenter",
                      },
                    },
                    {
                      key: "right",
                      text: "Right",
                      iconProps: {
                        officeFabricIconFontName: "AlignRight",
                      },
                    },
                  ],
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
