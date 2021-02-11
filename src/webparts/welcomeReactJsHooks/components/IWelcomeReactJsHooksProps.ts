import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IWelcomeReactJsHooksProps {
	title: string;
	displayMode: DisplayMode;
	updateProperty: (value: string) => void;
	context: WebPartContext;
	themeVariant: IReadonlyTheme | undefined;
	message: string;
	showname: string;
	showtimebasedmessage: boolean;
	morningmessage: string;
	afternoonmessage: string;
	afternoonbegintime: number;
	eveningmessage: string;
	eveningbegintime: number;
	textalignment: string;
	messagestyle: string;
}
