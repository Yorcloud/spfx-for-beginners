import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "AlertReactJsWebPartStrings";
import Welcome from "./components/AlertReactJs";
import { IAlertReactJsProps } from "./components/IAlertReactJsProps";
import { PropertyFieldDateTimePicker, DateConvention, TimeConvention, IDateTimeFieldValue } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy} from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages} from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';
import {
  ThemeProvider,
  ThemeChangedEventArgs,
  IReadonlyTheme,
} from "@microsoft/sp-component-base";
import { sp } from "@pnp/sp";

export interface IWelcomeWebPartProps {
  title: string;
  showtitle: boolean;
  body: string;
  showbodyashtml: boolean;
  enablestartdate: boolean;
  startdate: IDateTimeFieldValue;
  enableenddate: boolean;
  enddate: IDateTimeFieldValue;
  enablereadmessage: boolean;
  readmessage: string;
  readlist: string;
  showasmodal: boolean;
  messagelistid: string;
}

export default class WelcomeWebPart extends BaseClientSideWebPart<IWelcomeWebPartProps> {
  private _themeProvider: ThemeProvider;

  private _themeVariant: IReadonlyTheme;

  private _loading: boolean = false;

  protected onInit(): Promise<void> {
    return super.onInit().then((_) => {
      sp.setup({
        spfxContext: this.context,
        sp: {
          headers: {
            Accept: "application/json; odata=nometadata",
          },
        },
      });

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

  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;

    //return false; // Use true to show Apply button
  }

  public render(): void {
    const element: React.ReactElement<IAlertReactJsProps> = React.createElement(
      Welcome,
      {
        title: this.properties.title,
        showtitle: this.properties.showtitle,
        body: this.properties.body,
        showbodyashtml: this.properties.showbodyashtml,
        enablestartdate: this.properties.enablestartdate,
        startdate: this.properties.startdate,
        enableenddate: this.properties.enableenddate,
        enddate: this.properties.enddate,
        enablereadmessage: this.properties.enablereadmessage,
        readmessage: this.properties.readmessage,
        readlist: this.properties.readlist,
		displayMode: this.displayMode,
		showasmodal: this.properties.showasmodal,

        updateProperty: (value: string) => {
          this.properties.title = value;
        },
        loading: this._loading,
        themeVariant: this._themeVariant,
        context: this.context,
		messagelistid: this.properties.messagelistid
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

	let bodyEditor = this.properties.showbodyashtml ? 
	PropertyFieldCodeEditor('body', {
		label: 'Edit HTML Code',
		panelTitle: 'Edit HTML Code',
		initialValue: this.properties.body,
		onPropertyChange: this.onPropertyPaneFieldChanged,
		properties: this.properties,
		disabled: false,
		key: 'codeEditorFieldId',
		language: PropertyFieldCodeEditorLanguages.HTML,
		options: {
		  wrap: true,
		  fontSize: 12,
		  // more options
		}
	  }) : PropertyPaneTextField("body", {
		label: strings.BodyLabel,
		multiline: true,
		rows: 10,
	  });

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
				PropertyFieldListPicker('messagelistid', {
					label: 'Select a list',
					selectedList: this.properties.messagelistid,
					includeHidden: false,
					orderBy: PropertyFieldListPickerOrderBy.Title,
					disabled: false,
					onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
					properties: this.properties,
					context: this.context,
					onGetErrorMessage: null,
					deferredValidationTime: 0,
					key: 'listPickerFieldId'
				  }),
                PropertyPaneToggle("showtitle", {
                  label: strings.ShowTitleFieldLabel,
                }),
                bodyEditor,
                PropertyPaneToggle("showbodyashtml", {
                  label: strings.ShowBodyAsHtmlLabel,
                }),
                PropertyPaneToggle("enablereadmessage", {
                  label: strings.EnableReadMessageLabel,
                }),
                ,
                PropertyPaneTextField("readmessage", {
                  label: strings.ReadMessageLabel,
				}),
				PropertyPaneToggle("showasmodal", {
					label: strings.ShowAsModalLabel,
				  }),
                PropertyPaneToggle("enablestartdate", {
                  label: strings.EnableStartDateLabel,
				}),
				PropertyFieldDateTimePicker("startdate", {
					initialDate: this.properties.startdate,
					label: strings.StartDateLabel,
					properties: this.properties,
					onPropertyChange: this.onPropertyPaneFieldChanged,
					key: 'startdateFieldId'}
				),
                PropertyPaneToggle("enableenddate", {
                  label: strings.EnableEndDateLabel,
				}),
				PropertyFieldDateTimePicker("enddate", {
					initialDate: this.properties.enddate,
					label: strings.EndDateLabel,
					properties: this.properties,
					onPropertyChange: this.onPropertyPaneFieldChanged,
					key: 'enddateFieldId'}
				)
			
              ],
            },
          ],
        },
      ],
    };
  }
}

