import * as React from "react";
import { useEffect, useState } from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import { Text, DefaultButton } from "@fluentui/react";

import {
  DetailsList,
  IColumn,
  SelectionMode,
  DetailsListLayoutMode,
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
  Status: string;
  AssignedTo: string;
  Created: string;
}

export const MyRequests: React.FC<IMyRequestsProps> = ({
  spHttpClient,
  siteUrl,
  defaultMode,
}) => {
  const [tab, setTab] = useState<"my" | "approval">(defaultMode || "my");
  const [data, setData] = useState<IRequestItem[]>([]);

  // Pagination state
  const [currentPage, setCurrentPage] = useState<number>(1);
  const pageSize = 10;

  // Columns for Fluent UI DetailsList
  const columns: IColumn[] = [
    {
      key: "RequestID",
      name: "Request Type",
      fieldName: "RequestID",
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
    },
    {
      key: "Status",
      name: "Status",
      fieldName: "Status",
      minWidth: 120,
      maxWidth: 140,
      isResizable: true,
      onRender: (item: IRequestItem) => <StatusBadge status={item.Status} />,
    },
    {
      key: "AssignedTo",
      name: "Assigned Approver",
      fieldName: "AssignedTo",
      minWidth: 150,
      maxWidth: 180,
      isResizable: true,
    },
    {
      key: "Created",
      name: "Creation Date",
      fieldName: "Created",
      minWidth: 150,
      maxWidth: 160,
      isResizable: true,
    },
    {
      key: "Actions",
      name: "Actions",
      minWidth: 80,
      maxWidth: 100,
      isResizable: false,
      onRender: (item: IRequestItem) => (
        <ActionButton
          onClick={() =>
            window.open(
              `${siteUrl}/Lists/Requests/DispForm.aspx?ID=${item.id}`,
              "_blank"
            )
          }
        />
      ),
    },
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
        Status: item.Status ? item.Status.Title : "",
        AssignedTo: item.AssignedTo ? item.AssignedTo.Title : "",
        Created: item.Created,
      }));

      setData(mappedItems);
      setCurrentPage(1); // reset to first page when tab changes
    };

    fetchData();
  }, [spHttpClient, siteUrl, tab]);

  // Slice data for current page
  const pagedItems = data.slice(
    (currentPage - 1) * pageSize,
    currentPage * pageSize
  );
  const totalPages = Math.ceil(data.length / pageSize);

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

      {/* Table or empty message */}
      {pagedItems.length > 0 ? (
        <>
          <DetailsList
            items={pagedItems}
            columns={columns}
            selectionMode={SelectionMode.none}
            layoutMode={DetailsListLayoutMode.fixedColumns}
            isHeaderVisible={true}
          />

          {/* Pagination */}
          {data.length > pageSize && (
            <Stack
              horizontal
              tokens={{ childrenGap: 10 }}
              horizontalAlign="center"
              style={{ marginTop: 10 }}
            >
              <DefaultButton
                text="Previous"
                disabled={currentPage === 1}
                onClick={() => setCurrentPage(currentPage - 1)}
              />
              <span>
                Page {currentPage} of {totalPages}
              </span>
              <DefaultButton
                text="Next"
                disabled={currentPage === totalPages}
                onClick={() => setCurrentPage(currentPage + 1)}
              />
            </Stack>
          )}
        </>
      ) : (
        <Text
          variant="large"
          styles={{
            root: { textAlign: "center", marginTop: 20, color: "#666" },
          }}
        >
          No data found
        </Text>
      )}
    </Stack>
  );
};
