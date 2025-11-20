import * as React from "react";
import { IColumn } from "@fluentui/react";
import { StatusBadge } from "./StatusBadge"; // if you still want a colored circle
import { ActionButton } from "./ActionButton"; // icon-only button

export const columns: IColumn[] = [
  {
    key: "RequestID",
    name: "Request Type",
    fieldName: "RequestID",
    minWidth: 100,
    maxWidth: 150,
    isResizable: true,
  },
  // {
  //   key: "ServiceType",
  //   name: "Service Type",
  //   fieldName: "ServiceType",
  //   minWidth: 150,
  //   maxWidth: 200,
  //   isResizable: true,
  // },
  {
    key: "Status",
    name: "Status",
    fieldName: "Status",
    minWidth: 120,
    maxWidth: 150,
    isResizable: true,
    onRender: (item) => <StatusBadge status={item.Status} />,
  },
  {
    key: "AssignedTo",
    name: "Assigned Approver",
    fieldName: "AssignedTo",
    minWidth: 150,
    maxWidth: 200,
    isResizable: true,
  },
  {
    key: "Created",
    name: "Creation Date",
    fieldName: "Created",
    minWidth: 150,
    maxWidth: 200,
    isResizable: true,
  },
  {
    key: "Actions",
    name: "Actions",
    minWidth: 80,
    maxWidth: 100,
    isResizable: false,
    onRender: (item) => (
      <ActionButton
        onClick={() =>
          window.open(
            `https://ejadasharepoint.sharepoint.com/sites/WCA-DEV`,
            "_blank"
          )
        }
      />
    ),
  },
];
