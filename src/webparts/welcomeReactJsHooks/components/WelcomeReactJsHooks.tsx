import * as React from "react";
import styles from "./WelcomeReactJsHooks.module.scss";
import { IWelcomeReactJsHooksProps } from "./IWelcomeReactJsHooksProps";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { WebPartTitle } from "@pnp/spfx-controls-react";
import { FunctionComponent, useEffect, useState } from "react";

const WelcomeReactJsHooks: FunctionComponent<IWelcomeReactJsHooksProps> = (
  props
) => {
  const [messageContent, setMessageContent] = useState("");

  const { semanticColors }: IReadonlyTheme = props.themeVariant;

  useEffect(() => {
    let message = props.message;

    if (props.showtimebasedmessage) {
      const today: Date = new Date();
      if (today.getHours() >= props.eveningbegintime) {
        message = props.eveningmessage;
      }
      if (
        today.getHours() >= props.afternoonbegintime &&
        today.getHours() <= props.eveningbegintime
      ) {
        message = props.afternoonmessage;
      }
      if (today.getHours() < props.afternoonbegintime) {
        message = props.morningmessage;
      }
    }
    const nameparts = props.context.pageContext.user.displayName.split(" ");

    const textalign =
      props.textalignment === "left"
        ? styles.left
        : props.textalignment === "right"
        ? styles.right
        : styles.center;

    let messagecontent = null;
    let name = "";
    switch (props.showname) {
      case "full": {
        name = props.context.pageContext.user.displayName;
        break;
      }
      case "first": {
        name = nameparts[0];
        break;
      }
    }
    switch (props.messagestyle) {
      case "h3":
        messagecontent = (
          <h3 className={textalign} style={{ color: semanticColors.bodyText }}>
            {message} {name}
          </h3>
        );
        break;
      case "h2":
        messagecontent = (
          <h2 className={textalign} style={{ color: semanticColors.bodyText }}>
            {message} {name}
          </h2>
        );
        break;
      case "h1":
        messagecontent = (
          <h1 className={textalign} style={{ color: semanticColors.bodyText }}>
            {message} {name}
          </h1>
        );
        break;
      case "h4":
        messagecontent = (
          <h4 className={textalign} style={{ color: semanticColors.bodyText }}>
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

    setMessageContent(messagecontent);
  }, [props]);

  return (
    <div
      className={styles.welcomeReactJsHooks}
      style={{ backgroundColor: semanticColors.bodyBackground }}
    >
      <WebPartTitle
        displayMode={props.displayMode}
        title={props.title}
        updateProperty={props.updateProperty}
        themeVariant={props.themeVariant}
      />
      {messageContent}
    </div>
  );
};

export default WelcomeReactJsHooks;
