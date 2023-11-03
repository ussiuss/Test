import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import * as strings from "TestWebPartStrings";
import Test from "./components/Test";
import { ITestProps, BoxSize } from "./components/ITestProps";
import { SPHttpClient } from "@microsoft/sp-http";

export interface ITestWebPartProps {
  boxSize: BoxSize;
  description: string;
  selectedList: string;
  birthdays: any[]; // changed from any[] | IError to any[]
}

export default class TestWebPart extends BaseClientSideWebPart<ITestWebPartProps> {
  private _lists: { key: string; text: string }[] = [];
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";
  private _birthdays: any[] = []; // Local variable to store birthdays
  private dropdownOptions: Array<{ key: BoxSize; text: string }> = [
    { key: "small", text: "Small" },
    { key: "medium", text: "Medium" },
    { key: "large", text: "Large" },
    { key: "auto", text: "Auto" },
  ];

  public render(): void {
    const element: React.ReactElement<ITestProps> = React.createElement(Test, {
      context: this.context,
      selectedList: this.properties.selectedList,
      isDarkTheme: this._isDarkTheme,
      environmentMessage: this._environmentMessage,
      hasTeamsContext: !!this.context.sdks.microsoftTeams,
      userDisplayName: this.context.pageContext.user.displayName,
      birthdays: this._birthdays,
      boxSize: this.properties.boxSize,
    });

    ReactDom.render(element, this.domElement);
  }

  private _getBirthdays(): Promise<any> {
    if (!this.properties.selectedList) {
      return Promise.resolve({ error: "No list selected" });
    }

    const url =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/getbyid(guid'${this.properties.selectedList}')/items?$select=Person/Id,Person/Title,Person/EMail,Birthdate,Department&$expand=Person`;

    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response) => {
        if (!response.ok) {
          throw new Error("Failed to fetch data");
        }
        return response.json();
      })

      .then((jsonResponse) => {
        console.log("JSON Response:", jsonResponse); // Log the JSON response
        if (jsonResponse.error) {
          console.error("Error in response:", jsonResponse.error);
          return { error: "The list does not contain the necessary columns" };
        } else if (jsonResponse.value.length === 0) {
          return { error: "The list is empty" };
        } else {
          return jsonResponse.value;
        }
      })
      .catch((error) => {
        console.error("Error fetching data:", error);
        return { error: "An error occurred while fetching data" };
      });
  }

  private _getLists(): Promise<any[]> {
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists?$filter=Hidden eq false`,
        SPHttpClient.configurations.v1
      )
      .then((response) => response.json())
      .then((jsonResponse) => jsonResponse.value)
      .catch((error) => {
        console.error("Error fetching lists:", error);
        throw error; // re-throw error to be caught in onInit
      });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error("Unknown host");
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected onInit(): Promise<void> {
    return super.onInit().then((_) => {
      return this._getEnvironmentMessage().then((message) => {
        this._environmentMessage = message;
        return this._getLists().then((lists) => {
          this._lists = lists.map((list) => ({
            key: list.Id,
            text: list.Title,
          }));
          return this._getBirthdays().then((birthdays) => {
            this._birthdays = birthdays;
            this.render(); // Re-render after initialization is done
          });
        });
      });
    });
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): void {
    console.log("PropertyPaneFieldChanged called", propertyPath, newValue);

    if (propertyPath === "selectedList" && newValue !== oldValue) {
      // Update the selectedList property
      this.properties.selectedList = newValue;

      // Fetch the birthdays for the new list
      this._getBirthdays()
        .then((birthdays) => {
          // Update the local state with the new birthdays
          this._birthdays = birthdays;
          // Refresh the property pane and re-render the component
          this.context.propertyPane.refresh();
          this.render();
        })
        .catch((error) => {
          console.error("Error fetching new data:", error);
          this.render();
        });
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
                PropertyPaneDropdown("selectedList", {
                  label: "Select a list",
                  options: this._lists,
                }),
                PropertyPaneDropdown("boxSize", {
                  label: "Webpart size",
                  options: this.dropdownOptions,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
