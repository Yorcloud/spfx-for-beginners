import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";
import styles from "./WelcomeWebPart.module.scss";
import * as strings from "WelcomeWebPartStrings";

import {
  ThemeProvider,
  ThemeChangedEventArgs,
  IReadonlyTheme,
} from "@microsoft/sp-component-base";

export interface IWelcomeWebPartProps {
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

export default class WelcomeWebPart extends BaseClientSideWebPart<IWelcomeWebPartProps> {
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

  public render(): void {
    const { semanticColors }: IReadonlyTheme = this._themeVariant;

    let message = this.properties.message;

    if (this.properties.showtimebasedmessage) {
      const today: Date = new Date();
      if (today.getHours() >= this.properties.eveningbegintime) {
        message = this.properties.eveningmessage;
      }
      if (
        today.getHours() >= this.properties.afternoonbegintime &&
        today.getHours() <= this.properties.eveningbegintime
      ) {
        message = this.properties.afternoonmessage;
      }
      if (today.getHours() < this.properties.afternoonbegintime) {
        message = this.properties.morningmessage;
      }
    }
    const nameparts = this.context.pageContext.user.displayName.split(" ");

    const textalign =
      this.properties.textalignment === "left"
        ? styles.left
        : this.properties.textalignment === "right"
        ? styles.right
        : styles.center;

    let messagecontent = null;
    let name = "";
    switch (this.properties.showname) {
      case "full": {
        name = this.context.pageContext.user.displayName;
        break;
      }
      case "first": {
        name = nameparts[0];
        break;
      }
    }
    switch (this.properties.messagestyle) {
      case "h3":
        messagecontent = `<h3
              class='${textalign}'
              style='color: ${semanticColors.bodyText}'
            >
              ${message} ${name}
            </h3>`;
        break;
      case "h2":
        messagecontent = `<h2
				  class='${textalign}'
				  style='color: ${semanticColors.bodyText}'
				>
				  ${message} ${name}
				</h2>`;
        break;
      case "h1":
        messagecontent = `<h1
				  class='${textalign}'
				  style='color: ${semanticColors.bodyText}'
				>
				  ${message} ${name}
				</h1>`;
        break;
      case "h4":
        messagecontent = `<h4
				  class='${textalign}'
				  style='color: ${semanticColors.bodyText}'
				>
				  ${message} ${name}
				</h4>`;
        break;
      default:
        messagecontent = `<p class='${textalign}' style='color: ${semanticColors.bodyText}'>
              ${message} ${name}
            </p>`;
        break;
    }

    this.domElement.innerHTML = `
      <div class=${styles.welcome} style=${{
      backgroundColor: semanticColors.accentButtonBackground,
    }}>
	  <div class=${styles.webpartTitle}>${this.properties.title}</div>
	  ${messagecontent}
	  </div>`;
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
                PropertyPaneTextField("title", {
                  label: strings.TitleLabel,
                }),
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
                      key: "normal",
                      text: "Normal",
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
