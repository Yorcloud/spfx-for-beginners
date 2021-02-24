import * as React from "react";
import styles from "./WelcomeReactJsHooks.module.scss";
import { IWelcomeReactJsHooksProps } from "./IWelcomeReactJsHooksProps";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { WebPartTitle } from "@pnp/spfx-controls-react";
import { FunctionComponent, useEffect, useState } from "react";

const WelcomeReactJsHooks: FunctionComponent<IWelcomeReactJsHooksProps> = (
  props
) => {
  const [messageContent, setMessageContent] = useState<React.ReactElement>(null);

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

    const messagecontent = React.createElement(
      props.messagestyle,
      {
        className: textalign,
        style: { color: semanticColors.bodyText },
      },
      `${message} ${name}`
    );

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
