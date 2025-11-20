import * as React from "react";
import { useEffect, useState } from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import { Text } from "@fluentui/react";

import {
  DetailsList,
  IColumn,
  SelectionMode,
  DetailsListLayoutMode
} from "@fluentui/react/lib/DetailsList";
import { Pivot, PivotItem } from "@fluentui/react/lib/Pivot";
import { Stack } from "@fluentui/react/lib/Stack";

import { ActionButton } from "./ActionButton";
import { StatusBadge } from "./StatusBadge";

interface IMyRequestsProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
   defaultMode?: "my" | "approval"; // optional, from property pane
}

interface IRequestItem {
  id: number;
  RequestID: string;
 // ServiceType: string;
  Status: string;
  AssignedTo: string;
  Created: string;
}

export const MyRequests: React.FC<IMyRequestsProps> = ({
  spHttpClient,
  siteUrl,
  defaultMode,
}) => {
  //const [tab, setTab] = useState<"my" | "approval">("my");
  const [tab, setTab] = useState<"my" | "approval">(defaultMode || "my");
  const [data, setData] = useState<IRequestItem[]>([]);

  // Columns for Fluent UI DetailsList
  const columns: IColumn[] = [
    { key: "RequestID", name: "Request Type", fieldName: "RequestID", minWidth: 100, maxWidth: 120, isResizable: true },
   // { key: "ServiceType", name: "Service Type", fieldName: "ServiceType", minWidth: 150, maxWidth: 180, isResizable: true },
    {
      key: "Status",
      name: "Status",
      fieldName: "Status",
      minWidth: 120,
      maxWidth: 140,
      isResizable: true,
      onRender: (item: IRequestItem) => <StatusBadge status={item.Status} />
    },
    { key: "AssignedTo", name: "Assigned To", fieldName: "AssignedTo", minWidth: 150, maxWidth: 180, isResizable: true },
    { key: "Created", name: "Created", fieldName: "Created", minWidth: 150, maxWidth: 160, isResizable: true },
    {
      key: "Actions",
      name: "Actions",
      minWidth: 80,
      maxWidth: 100,
      isResizable: false,
      onRender: (item: IRequestItem) => <ActionButton onClick={() => window.open(`${siteUrl}/Lists/Requests/DispForm.aspx?ID=${item.id}`, "_blank")} />
    }
  ];

  useEffect(() => {
    const fetchData = async () => {
      // Get current user
      const userResponse = await spHttpClient.get(
        `${siteUrl}/_api/web/currentuser`,
        SPHttpClient.configurations.v1
      );
      const currentUser = await userResponse.json();
      const email = currentUser.Email;

      const filter =
        tab === "approval"
          ? `AssignedTo/EMail eq '${email}'`
          : `Author/EMail eq '${email}'`;
 
      const response = await spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('Requests')/items` +
          `?$select=Id,RequestID,Status/Title,AssignedTo/Title,Created` +
          `&$expand=AssignedTo,Status` +
          `&$filter=${filter}`,
        SPHttpClient.configurations.v1
      );

      const items = await response.json();
      const mappedItems = items.value.map((item: any) => ({
        id: item.Id,
        RequestID: item.RequestID,
       // ServiceType: item.ServiceType ? item.ServiceType.Title : "",
        Status: item.Status ? item.Status.Title : "",
        AssignedTo: item.AssignedTo ? item.AssignedTo.Title : "",
        Created: item.Created,
      }));

      setData(mappedItems);
    };

    fetchData();
  }, [spHttpClient, siteUrl, tab]);

  return (
    <Stack tokens={{ childrenGap: 10 }}>
       {/* Title */}
    <Text variant="xLarge" block>
      Requests & Approvals
    </Text>

      {/* Pivot Tabs */}
      <Pivot
        selectedKey={tab}
        onLinkClick={(item) => setTab(item?.props.itemKey as "my" | "approval")}
      >
        <PivotItem headerText="My Requests" itemKey="my" />
        <PivotItem headerText="My Approvals" itemKey="approval" />
      </Pivot>

      {/* Table */}
      <DetailsList
        items={data}
        columns={columns}
        selectionMode={SelectionMode.none}
        layoutMode={DetailsListLayoutMode.fixedColumns}
        isHeaderVisible={true}
      />
    </Stack>
  );
};
