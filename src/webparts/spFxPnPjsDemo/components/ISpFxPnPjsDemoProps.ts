import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISpFxPnPjsDemoProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}
