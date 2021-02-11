declare interface IWelcomeWebPartStrings {
	EveningBeginTimeLabel: string;
	AfternoonBeginTimeLabel: string;
	TextAlignmentLabel: string;
	MessageStyleLabel: string;
	StylePropertiesGroupName: string;
	NamePropertiesGroupName: string;
	MessagePropertiesGroupName: string;
	MorningMessageLabel: string;
	AfternoonMessageLabel: string;
	EveningMessageLabel: string;
	ShowTimeBasedMessageLabel: string;
	ShowFirstNameLabel: string;
	ShowNameLabel: string;
	PropertyPaneDescription: string;
	PropertiesGroupName: string;
	MessageLabel: string;
	TitleLabel: string;
}

declare module 'WelcomeWebPartStrings' {
  const strings: IWelcomeWebPartStrings;
  export = strings;
}
