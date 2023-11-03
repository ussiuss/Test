import { WebPartContext } from "@microsoft/sp-webpart-base";
export type BoxSize = "small" | "medium" | "large" | "auto";

export interface ITestProps {
  context: WebPartContext;
  selectedList: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  birthdays: any[];
  boxSize: BoxSize;
}
