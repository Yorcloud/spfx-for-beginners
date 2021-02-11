import * as React from "react";
import styles from "./WelcomeReactJs.module.scss";
import { IWelcomeReactJsProps } from "./IWelcomeReactJsProps";
import { MessageBar, MessageBarType, Shimmer, Label } from 'office-ui-fabric-react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { WebPartTitle } from "@pnp/spfx-controls-react";

export default class WelcomeReactJs extends React.Component<
  IWelcomeReactJsProps,
  {}
> {
  private _onConfigure() {
    // Context of the web part

    this.props.context.propertyPane.open();
  }

  public render(): React.ReactElement<IWelcomeReactJsProps> {
    const { semanticColors }: IReadonlyTheme = this.props.themeVariant;

      let message = this.props.message;

      if (this.props.showtimebasedmessage) {
        const today: Date = new Date();
        if (today.getHours() >= this.props.eveningbegintime) {
          message = this.props.eveningmessage;
        }
        if (
          today.getHours() >= this.props.afternoonbegintime &&
          today.getHours() <= this.props.eveningbegintime
        ) {
          message = this.props.afternoonmessage;
        }
        if (today.getHours() < this.props.afternoonbegintime) {
          message = this.props.morningmessage;
        }
      }
      const nameparts = this.props.context.pageContext.user.displayName.split(
        " "
      );

      const textalign =
        this.props.textalignment === "left"
          ? styles.left
          : this.props.textalignment === "right"
          ? styles.right
          : styles.center;

      let messagecontent = null;
      let name = "";
      switch (this.props.showname) {
        case "full": {
          name = this.props.context.pageContext.user.displayName;
          break;
        }
        case "first": {
          name = nameparts[0];
          break;
        }
      }
      switch (this.props.messagestyle) {
        case "h3":
          messagecontent = (
            <h3
              className={textalign}
              style={{ color: semanticColors.bodyText }}
            >
              {message} {name}
            </h3>
          );
          break;
        case "h2":
          messagecontent = (
            <h2
              className={textalign}
              style={{ color: semanticColors.bodyText }}
            >
              {message} {name}
            </h2>
          );
          break;
        case "h1":
          messagecontent = (
            <h1
              className={textalign}
              style={{ color: semanticColors.bodyText }}
            >
              {message} {name}
            </h1>
          );
          break;
        case "h4":
          messagecontent = (
            <h4
              className={textalign}
              style={{ color: semanticColors.bodyText }}
            >
              {message} {name}
            </h4>
          );
          break;
        default:
          messagecontent = (
            <p className={textalign} style={{ color: semanticColors.bodyText }}>
              {message} {name}
            </p>
          );
          break;
      }

      return (
        <div
          className={styles.welcomeReactJs}
          style={{ backgroundColor: semanticColors.bodyBackground }}
        >
          <WebPartTitle
            displayMode={this.props.displayMode}
            title={this.props.title}
            updateProperty={this.props.updateProperty}
            themeVariant={this.props.themeVariant}
          />
          {messagecontent}
        </div>
      );
    }
  
}
