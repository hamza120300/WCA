import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  // PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "MyRequestsWebPartStrings";
// import MyRequests from './components/MyRequests';  // event recever
// import { IMyRequestsProps } from './components/IMyRequestsProps';
import { MyRequests } from "./components/MyRequests"; // notice we use named export

export interface IMyRequestsWebPartProps {
  description: string;
  mode: string; // "my" or "approval"
  showServiceType: boolean;
  showCreated: boolean;
  showStatus: boolean;
  showAssignedTo: boolean;
  showActions: boolean;
}

export default class MyRequestsWebPart extends BaseClientSideWebPart<IMyRequestsWebPartProps> {
  public render(): void {
    const initialTab: "my" | "approval" =
      this.properties.mode === "approval" ? "approval" : "my";
    const element = React.createElement(MyRequests, {
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      defaultMode: initialTab, // pass initial tab
      columnVisibility: {
        serviceType: this.properties.showServiceType !== false,
        created: this.properties.showCreated !== false,
        status: this.properties.showStatus !== false,
        assignedTo: this.properties.showAssignedTo !== false,
        actions: this.properties.showActions !== false,
      },
    });

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then((message) => {
      // this._environmentMessage = message;
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
            case "TeamsModern":
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
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

    //this._isDarkTheme = !!currentTheme.isInverted;
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
          header: { description: "Settings" },
          groups: [
            {
              groupName: "Mode",
              groupFields: [
                PropertyPaneDropdown("mode", {
                  label: "Select Mode",
                  options: [
                    { key: "my", text: "My Requests" },
                    { key: "approval", text: "My Approvals" },
                  ],
                }),
              ],
            },
            {
              groupName: "Column Visibility",
              groupFields: [
                PropertyPaneToggle("showServiceType", {
                  label: "Show Service Type",
                  onText: "Shown",
                  offText: "Hidden",
                }),
                PropertyPaneToggle("showCreated", {
                  label: "Show Created Date",
                  onText: "Shown",
                  offText: "Hidden",
                }),
                PropertyPaneToggle("showStatus", {
                  label: "Show Status",
                  onText: "Shown",
                  offText: "Hidden",
                }),
                PropertyPaneToggle("showAssignedTo", {
                  label: "Show Assigned To",
                  onText: "Shown",
                  offText: "Hidden",
                }),
                PropertyPaneToggle("showActions", {
                  label: "Show Actions",
                  onText: "Shown",
                  offText: "Hidden",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
