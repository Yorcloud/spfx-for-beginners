import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { IDateTimeFieldValue } from "@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker";

export interface IAlertReactJsProps {
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
	displayMode: DisplayMode;
	messagelistid: string;
  
	showasmodal: boolean;
  
	updateProperty: (value: string) => void;
	context: WebPartContext;
  
	themeVariant: IReadonlyTheme;
  
	loading?: boolean;
}