import * as React from "react";
import { useEffect, useState } from "react";
import Paper from "@mui/material/Paper";
import { DataGrid } from "@mui/x-data-grid";
import { columns } from "./columns";
import { SPHttpClient } from "@microsoft/sp-http";
import { Tabs, Tab, Box } from "@mui/material";

interface IMyRequestsProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  mode: string;
}

export const MyRequests: React.FC<IMyRequestsProps> = ({
  spHttpClient,
  siteUrl,
  mode,
}) => {
  const [tab, setTab] = useState<"my" | "approval">("my");
  const [data, setData] = useState<any[]>([]);

  useEffect(() => {
    const fetchData = async () => {
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
          `?$select=Id,RequestID,ServiceType/Title,Status/Title,AssignedTo/Title,Created` +
          `&$expand=AssignedTo,ServiceType,Status` +
          `&$filter=${filter}`,
        SPHttpClient.configurations.v1
      );

      const items = await response.json();
      const mappedItems = items.value.map((item: any) => ({
        id: item.Id,
        RequestID: item.RequestID,
        ServiceType: item.ServiceType ? item.ServiceType.Title : "",
        Status: item.Status ? item.Status.Title : "",
        AssignedTo: item.AssignedTo ? item.AssignedTo.Title : "",
        Created: item.Created,
      }));

      setData(mappedItems);
    };

    fetchData();
  }, [spHttpClient, siteUrl, tab]);

  return (
    <Box>
      {/* Tabs */}
      <Tabs
        value={tab}
        onChange={(_, newValue) => setTab(newValue)}
        sx={{ marginBottom: 2 }}
      >
        <Tab value="my" label="My Requests" />
        <Tab value="approval" label="My Approvals" />
      </Tabs>

      {/* Data Table */}
      <Paper sx={{ height: 500, width: "100%", padding: 2 }}>
        <DataGrid
          rows={data}
          columns={columns}
          pageSizeOptions={[5, 10, 20]}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          sx={{ border: 0 }}
        />
      </Paper>
    </Box>
  );
};
