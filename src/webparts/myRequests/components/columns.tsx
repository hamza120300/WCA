import { GridColDef } from "@mui/x-data-grid";
import * as React from "react";
import ActionButton from "./ActionButton";

export const columns: GridColDef[] = [
  { field: "RequestID", headerName: "Request ID", width: 120 },
  { field: "ServiceType", headerName: "Service Type", width: 180 },
  { field: "Status", headerName: "Status", width: 140 },
  { field: "AssignedTo", headerName: "Assigned To", width: 180 },
  { field: "Created", headerName: "Created", width: 160 },

  {
    field: "Actions",
    headerName: "Actions",
    width: 150,
    sortable: false,
    renderCell: (params) => {
      return <ActionButton />;
    },
  },
];
