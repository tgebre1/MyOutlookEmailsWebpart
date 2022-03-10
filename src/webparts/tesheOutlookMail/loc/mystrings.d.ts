declare interface ITesheOutlookMailWebPartStrings {
  NoMessages: ReactNode;
  Loading: string;
  ViewAll: string;
  NewEmail: string;
  Error: string;
  ShowInboxOnlyCallout(
    arg0: string,
    arg1: {},
    ShowInboxOnlyCallout: any
  ): React.ReactNode;
  ShowInboxOnly:
    | string
    | ReactElement<any, string | JSXElementConstructor<any>>;
  NrOfMessagesToShow: any;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module "TesheOutlookMailWebPartStrings" {
  const strings: ITesheOutlookMailWebPartStrings;
  export = strings;
}
