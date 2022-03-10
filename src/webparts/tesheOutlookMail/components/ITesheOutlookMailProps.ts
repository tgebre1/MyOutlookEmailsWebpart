import { ITesheOutlookMailWebPartProps } from "../TesheOutlookMailWebPart";
import { MSGraphClient } from "@microsoft/sp-http";
import { DisplayMode } from "@microsoft/sp-core-library";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

export interface ITesheOutlookMailProps extends ITesheOutlookMailWebPartProps {
  displayMode: DisplayMode;
  graphClient: MSGraphClient;
  themeVariant: IReadonlyTheme | undefined;
  iconPicker: string;

  updateProperty: (value: string) => void;
}
