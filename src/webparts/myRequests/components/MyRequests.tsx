import * as React from "react";
import { useEffect, useState } from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import { Text } from "@fluentui/react";
import styles from "./MyRequests.module.scss";
import { Spinner, SpinnerSize } from "@fluentui/react";

//, DefaultButton // start add arrow

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
  columnVisibility?: {
    serviceType?: boolean;
    created?: boolean;
    status?: boolean;
    assignedTo?: boolean;
    actions?: boolean;
  };
}

const isArabic =
  window.location.pathname.toLowerCase().indexOf("/sitepages/ar/") !== -1;

const tabs = [
  { key: "my", text: isArabic ? "طلباتي" : "Created by me" },
  { key: "approval", text: isArabic ? "مسند الي" : "Assigned to me" },
];

interface IRequestItem {
  id: number;
  RequestID: string;
  ServiceType: string;
  Status: string;
  AssignedTo: string;
  Created: string;
  AssignedToEmail?: string;
  DetailsPage: string; // NEW
}

export const MyRequests: React.FC<IMyRequestsProps> = ({
  spHttpClient,
  siteUrl,
  defaultMode,
  columnVisibility = {},
}) => {
  const [tab, setTab] = useState<"my" | "approval">(defaultMode || "my");
  const [detailsRootURL, setDetailsRootURL] = React.useState("");
  const [data, setData] = useState<IRequestItem[]>([]);
  const [sortedColumn, setSortedColumn] = useState<string | undefined>();
  const [isSortedDescending, setIsSortedDescending] = useState<
    boolean | undefined
  >();
  const [loading, setLoading] = useState(true);
  //const [loaded, setLoaded] = useState(false);

  const onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ) => {
    const newDesc = sortedColumn === column.key ? !isSortedDescending : false;

    const sorted = [...data].sort((a, b) => {
      const aVal = a[column.fieldName as keyof IRequestItem] || "";
      const bVal = b[column.fieldName as keyof IRequestItem] || "";

      return newDesc
        ? ("" + bVal).localeCompare("" + aVal)
        : ("" + aVal).localeCompare("" + bVal);
    });

    setSortedColumn(column.key);
    setIsSortedDescending(newDesc);
    setData(sorted);
  };

  // Pagination state
  const [currentPage, setCurrentPage] = useState<number>(1);
  const pageSize = 6;

  const getUserPhoto = (email: string) =>
    `${siteUrl}/_layouts/15/userphoto.aspx?size=S&accountname=${email}`;

  interface IPaginationProps {
    currentPage: number;
    totalPages: number;
    onChange: (page: number) => void;
  }

  const Pagination: React.FC<IPaginationProps> = ({
    currentPage,
    totalPages,
    onChange,
  }) => {
    const renderPages = () => {
      const pages: (number | string)[] = [];

      pages.push(1);

      if (currentPage > 3) pages.push("...");

      for (let p = currentPage - 1; p <= currentPage + 1; p++) {
        if (p > 1 && p < totalPages) pages.push(p);
      }

      if (currentPage < totalPages - 2) pages.push("...");

      if (totalPages > 1) pages.push(totalPages);

      return pages;
    };

    const pageStyle = (isCurrent: boolean) => ({
      display: "inline-block", // make numbers block-like
      padding: "4px 8px", // space around number
      margin: "0 4px",
      cursor: isCurrent ? "default" : "pointer",
      fontWeight: isCurrent ? 600 : 400,
      color: isCurrent ? "#3aa272ff" : "#000",
      borderBottom: isCurrent ? "2px solid #3aa272ff" : "2px solid transparent", // green underline
      // borderRadius: "2px", // optional: small rounding
      // textAlign: "center",
    });

    return (
      <Stack
        horizontal
        horizontalAlign="start"
        tokens={{ childrenGap: 2 }}
        style={{ marginTop: 10 }}
      >
        <span
          style={{
            cursor: currentPage === 1 ? "default" : "pointer",
            marginRight: 6,
          }}
          onClick={() => currentPage > 1 && onChange(currentPage - 1)}
        >
          &lt;
        </span>

        {renderPages().map((p, idx) =>
          p === "..." ? (
            <span key={idx} style={{ margin: "0 4px" }}>
              …
            </span>
          ) : (
            <span
              key={idx}
              style={pageStyle(p === currentPage)}
              onClick={() => p !== currentPage && onChange(Number(p))}
            >
              {p}
            </span>
          )
        )}

        <span
          style={{
            cursor: currentPage === totalPages ? "default" : "pointer",
            marginLeft: 6,
          }}
          onClick={() => currentPage < totalPages && onChange(currentPage + 1)}
        >
          &gt;
        </span>
      </Stack>
    );
  };

  // const renderRow = (props: any) => {
  //   if (!props) return null;

  //   const item: IRequestItem = props.item;

  //   // Only make row clickable if Actions column is hidden
  //   if (!columnVisibility.actions && item) {
  //     const encodedRequestId = btoa(item.RequestID);
  //     const url = `${detailsRootURL}${
  //       item.DetailsPage
  //     }?requestId=${encodedRequestId}${isArabic ? "&locale=ar-sa" : ""}`;

  //     return (
  //       <div
  //         {...props}
  //         style={{ cursor: "pointer" }}
  //         onClick={() => window.open(url, "_blank")}
  //       />
  //     );
  //   }

  //   return <div {...props} />;
  // };

  // Columns for Fluent UI DetailsList
  // Define columns based on visibility settings h
  // const serviceTypeColumn: IColumn = {
  //   key: "ServiceType",
  //   name: isArabic ? "نوع الخدمة" : "Request Type",
  //   fieldName: "ServiceType",
  //   minWidth: columnVisibility.assignedTo ? 115 : 270,
  //   maxWidth: columnVisibility.assignedTo ? 135 : 310,
  //   isResizable: true,
  //   isSorted: sortedColumn === "ServiceType",
  //   isSortedDescending: isSortedDescending,
  //   onColumnClick: onColumnClick,
  // };
  const serviceTypeColumn: IColumn = {
    key: "ServiceType",
    name: isArabic ? "نوع الخدمة" : "Request Type",
    fieldName: "ServiceType",
    minWidth: columnVisibility.assignedTo ? 115 : 270,
    maxWidth: columnVisibility.assignedTo ? 135 : 310,
    isResizable: true,
    isSorted: sortedColumn === "ServiceType",
    isSortedDescending: isSortedDescending,
    onColumnClick: onColumnClick,

    onRender: (item: IRequestItem) => {
      // If Actions is hidden → make ServiceType clickable
      if (!columnVisibility.actions) {
        const encodedRequestId = btoa(item.RequestID);
        const url = `${detailsRootURL}${
          item.DetailsPage
        }?requestId=${encodedRequestId}${isArabic ? "&locale=ar-sa" : ""}`;

        return (
          <span
            style={{
              // color: "#0078d4",
              cursor: "pointer",
              // textDecoration: "underline",
            }}
            onClick={() => window.open(url, "_blank")}
          >
            {item.ServiceType}
          </span>
        );
      }

      // Default (normal text)
      return <span>{item.ServiceType}</span>;
    },
  };

  const createdColumn: IColumn = {
    key: "Created",
    name: isArabic ? "تاريخ الإنشاء" : "Created Date",
    fieldName: "Created",
    // minWidth: 80,
    //maxWidth: 100,
    minWidth: columnVisibility.assignedTo ? 80 : 200,
    maxWidth: columnVisibility.assignedTo ? 100 : 220,
    isResizable: true,
    isSorted: sortedColumn === "Created",
    isSortedDescending: isSortedDescending,
    onColumnClick: onColumnClick,
    onRender: (item: IRequestItem) => {
      if (!item.Created) return "-";

      const formatted = new Date(item.Created)
        .toLocaleDateString("en-GB", {
          day: "2-digit",
          month: "short",
          year: "numeric",
        })
        .replace(/ /g, " ");

      return (
        <span style={{ direction: "ltr", unicodeBidi: "isolate" }}>
          {formatted}
        </span>
      );
    },
  };

  const statusColumn: IColumn = {
    key: "Status",
    name: isArabic ? "الحالة" : "Status",
    fieldName: "Status",
    minWidth: 80,
    maxWidth: 90,
    isResizable: true,
    onRender: (item: IRequestItem) => (
      <StatusBadge status={item.Status} lang={isArabic ? "ar" : "en"} />
    ),
    isSorted: sortedColumn === "Status",
    isSortedDescending: isSortedDescending,
    onColumnClick: onColumnClick,
  };

  const assignedToColumn: IColumn = {
    key: "AssignedTo",
    name: isArabic ? "الجهة المكلفة بالموافقة" : "Assigned Approver",
    fieldName: "AssignedTo",
    minWidth: 150,
    maxWidth: 200,
    isResizable: true,
    onRender: (item: IRequestItem) => {
      if (!item.AssignedTo) return "-";

      return (
        <div
          className={styles["person-pill"]}
          style={{ flexDirection: isArabic ? "row-reverse" : "row" }}
        >
          <img
            src={getUserPhoto(item.AssignedToEmail || "")}
            alt={item.AssignedTo}
            className="personDisplayCoin_d125512b"
          />
          <span className="personDisplayName_d125512b">{item.AssignedTo}</span>
        </div>
      );
    },
  };

  const actionsColumn: IColumn = {
    key: "Actions",
    name: isArabic ? "الإجراءات" : "Actions",
    minWidth: 60,
    maxWidth: 60,
    isResizable: false,
    onRender: (item: IRequestItem) => (
      <ActionButton
        onClick={() => {
          const encodedRequestId = btoa(item.RequestID);
          window.open(
            `${detailsRootURL}${
              item.DetailsPage
            }?requestId=${encodedRequestId}${isArabic ? "&locale=ar-sa" : ""}`,
            "_blank"
          );
        }}
      />
    ),
  };

  const columns: IColumn[] = [];

  if (columnVisibility.serviceType) columns.push(serviceTypeColumn);
  if (columnVisibility.created) columns.push(createdColumn);
  if (columnVisibility.status) columns.push(statusColumn);
  if (columnVisibility.assignedTo) columns.push(assignedToColumn);
  if (columnVisibility.actions) columns.push(actionsColumn);
  // Apply column visibility from props
  //  //  "https://ejadasharepoint.sharepoint.com/sites/WCA-DEV/_layouts/15/workbench.aspx"
  useEffect(() => {
    const fetchData = async () => {
      setLoading(true);
      // setLoaded(false);
      try {
        //fetch from Configurations
        const configsResponse = await spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('Configurations')/items?$select=Title,Value`,
          SPHttpClient.configurations.v1
        );
        const configsJson = await configsResponse.json();

        let root = "";
        configsJson.value.forEach((cfg: any) => {
          if (cfg.Title === "DetailsRootURL") root = cfg.Value || "";
        });

        setDetailsRootURL(root);

        // Get current user
        const userResponse = await spHttpClient.get(
          `${siteUrl}/_api/web/currentuser`,
          SPHttpClient.configurations.v1
        );
        const currentUser = await userResponse.json();
        const email = currentUser.Email.toLowerCase();

        const groupsResponse = await spHttpClient.get(
          `${siteUrl}/_api/web/currentuser/groups`,
          SPHttpClient.configurations.v1
        );
        const groupsJson = await groupsResponse.json();

        const userGroups: string[] = (groupsJson.value || []).map((g: any) =>
          String(g.Title || "")
        );

        const response = await spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('Requests')/items` +
            `?$select=Id,RequestID,Requestor,Author/EMail,ServiceType/DetailsPage,ServiceType/Title,ServiceType/Title_Ar,Status/StatusSummary,Status/StatusSummary_Ar,AssignedTo/Title,AssignedTo/EMail,Created` +
            `&$expand=AssignedTo,Status,ServiceType,Author` +
            `&$orderby=Created desc`,
          SPHttpClient.configurations.v1
        );

        const items = await response.json();

        //   const mappedItems = items.value.map((item: any) => ({
        const mappedItems = (items.value || []).map((item: any) => ({
          id: item.Id,
          RequestID: item.RequestID,
          DetailsPage: item.ServiceType?.DetailsPage || "",
          ServiceType: item.ServiceType
            ? isArabic
              ? item.ServiceType.Title_Ar
              : item.ServiceType.Title
            : "",
          Status: item.Status ? item.Status.StatusSummary : "",
          AssignedTo: item.AssignedTo ? item.AssignedTo.Title : "",
          AssignedToEmail: item.AssignedTo?.EMail || "",
          AuthorEmail: item.Author?.EMail || "",
          Created: item.Created,
          Requestor: item.Requestor || "",
        }));
        console.log("DetailsPage", mappedItems[0].Requestor);
        let filteredItems: IRequestItem[] = [];

        if (tab === "my") {
          filteredItems = mappedItems.filter((mi: any) => {
            return (mi.Requestor || "").toLowerCase() === email;
          });
        } else {
          filteredItems = mappedItems.filter((mi: any) => {
            if ((mi.AssignedToEmail || "").toLowerCase() === email) return true;
            if (mi.AssignedTo && userGroups.indexOf(mi.AssignedTo) !== -1)
              return true;
            return false;
          });
        }

        setData(filteredItems);
        setCurrentPage(1);
      } finally {
        setLoading(false);
        //  setLoaded(true);
      }
    };

    fetchData();
  }, [spHttpClient, siteUrl, tab]);

  // ------------

  // Slice data for current page
  const pagedItems = data.slice(
    (currentPage - 1) * pageSize,
    currentPage * pageSize
  );
  const totalPages = Math.ceil(data.length / pageSize);

  return (
    <Stack
      tokens={{ childrenGap: window.innerWidth < 600 ? 4 : 10 }}
      style={{
        direction: isArabic ? "rtl" : "ltr",
        textAlign: isArabic ? "right" : "left",
      }}
    >
      {/* Title */}
      <div
        style={{
          direction: isArabic ? "rtl" : "ltr",
          textAlign: isArabic ? "right" : "left",
        }}
      >
        <Text variant="xLarge" block>
          {isArabic ? "الطلبات والموافقات" : "Requests & Approvals"}
        </Text>
      </div>

      {/* Pivot Tabs */}
      <Pivot
        selectedKey={tab}
        onLinkClick={(item) => setTab(item?.props.itemKey as "my" | "approval")}
        style={{
          direction: isArabic ? "rtl" : "ltr",
          textAlign: isArabic ? "right" : "left",
        }}
        // linkFormat="tabs"
        //headersOnly={false}
      >
        {tabs.map((tab) => (
          <PivotItem
            key={tab.key}
            itemKey={tab.key}
            headerText={tab.text}
            headerButtonProps={{
              style: { textAlign: isArabic ? "right" : "left" },
            }}
          />
        ))}
      </Pivot>

      {/* Table + Header */}
      <div style={{ height: "100%", overflowY: "auto", overflowX: "auto" }}>
        <DetailsList
          items={pagedItems.length === 0 ? [] : pagedItems}
          columns={columns}
          selectionMode={SelectionMode.none}
          layoutMode={DetailsListLayoutMode.fixedColumns}
          isHeaderVisible={true}
          onRenderMissingItem={() => null} // prevents SharePoint from rendering empty row
          //  onRenderRow={renderRow}
          onRenderItemColumn={(item, index, column) => {
            if (!column) return null;

            const value = item[column.fieldName as keyof IRequestItem];

            if (column.key === "Status")
              return <StatusBadge status={value as string} />;

            if (column.key === "Actions")
              return (
                <ActionButton
                  onClick={() => {
                    const encodedRequestId = btoa(item.RequestID);
                    window.open(
                      `${detailsRootURL}${item.DetailsPage}?requestId=${encodedRequestId}`,
                      "_blank"
                    );
                  }}
                />
              );

            return <span>{value}</span>;
          }}
        />
      </div>
      {/* Loading Spinner loading */}
      {loading && pagedItems.length === 0 && (
        <div
          style={{ textAlign: "center", padding: "40px 0", color: "#CAF0CC" }}
        >
          <Spinner
            size={SpinnerSize.large}
            label="Loading..."
            //  labelClassName="spinner-label-green"
          />
        </div>
      )}

      {/* Empty State BELOW HEADERS */}
      {!loading && pagedItems.length === 0 && (
        <div
          style={{
            textAlign: "center",
            padding: "60px 0",
            color: "#555",
          }}
        >
          <img
            src={require("../assets/noRequests.svg")}
            alt="No Requests"
            style={{ width: 140, opacity: 0.95 }}
          />
          <div style={{ marginTop: 16, fontSize: 16 }}>
            You don’t have any request
          </div>
        </div>
      )}

      {/* Pagination */}
      {data.length > pageSize && (
        <Pagination
          currentPage={currentPage}
          totalPages={totalPages}
          onChange={setCurrentPage}
        />
      )}
    </Stack>
  );
};
