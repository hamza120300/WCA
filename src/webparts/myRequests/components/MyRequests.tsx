import * as React from "react";
import { useEffect, useState } from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import { Text } from "@fluentui/react";
import styles from "./MyRequests.module.scss";

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
}

const isArabic =
  window.location.pathname.toLowerCase().indexOf("/sitepages/ar/") !== -1;

const tabs = [
  { key: "my", text: isArabic ? "طلباتي" : "Created by me" },
  { key: "approval", text: isArabic ? "مسند الي" : "Assigned to me" },
];



interface IRequestItem {
  id: number;
  // RequestID: string;
  ServiceType: string;
  Status: string;    
  AssignedTo: string;
  Created: string;
  AssignedToEmail?: string;
}

export const MyRequests: React.FC<IMyRequestsProps> = ({
  spHttpClient,
  siteUrl,
  defaultMode,
}) => {
  const [tab, setTab] = useState<"my" | "approval">(defaultMode || "my");
  const [data, setData] = useState<IRequestItem[]>([]);
  const [sortedColumn, setSortedColumn] = useState<string | undefined>();
  const [isSortedDescending, setIsSortedDescending] = useState<
    boolean | undefined
  >();

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
      margin: "0 6px",
      cursor: isCurrent ? "default" : "pointer",
      fontWeight: isCurrent ? 600 : 400,
      color: isCurrent ? "#0078D4" : "#000",
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

  // Columns for Fluent UI DetailsList
  const columns: IColumn[] = [
    {
      key: "ServiceType",
      name: isArabic ? "نوع الخدمة" : "Request Type",
      fieldName: "ServiceType",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      isSorted: sortedColumn === "ServiceType",
      isSortedDescending: isSortedDescending,
      onColumnClick: onColumnClick,
    },

    {
      key: "Created",
      name: isArabic ? "تاريخ الإنشاء" : "Created Date",
      fieldName: "Created",
      minWidth: 150,
      maxWidth: 160,
      isResizable: true,
      isSorted: sortedColumn === "Created",
      isSortedDescending: isSortedDescending,
      onColumnClick: onColumnClick,
      onRender: (item: IRequestItem) => {
        if (!item.Created) return "-";
        // Take only the date part
        return item.Created.split("T")[0];
      },
    },
    {
      key: "Status",
      name: isArabic ? "الحالة" : "Status",
      fieldName: "Status",
      minWidth: 120,
      maxWidth: 220,
      isResizable: true,
      onRender: (item: IRequestItem) => (
        <StatusBadge status={item.Status} lang={isArabic ? "ar" : "en"} />
      ),
      isSorted: sortedColumn === "Status",
      isSortedDescending: isSortedDescending,
      onColumnClick: onColumnClick,
    },

    {
      key: "AssignedTo",
      name: isArabic ? "الجهة المكلفة بالموافقة" : "Assigned Approver",
      fieldName: "AssignedTo",
      minWidth: 200,
      maxWidth: 220,
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
            <span className="personDisplayName_d125512b">
              {item.AssignedTo}
            </span>
          </div>
        );
      },
    },

    {
      key: "Actions",
      name: isArabic ? "الإجراءات" : "Actions",
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
          `?$select=Id,RequestID,ServiceType/Title,ServiceType/Title_Ar,Status/Title,Status/Title_Ar,AssignedTo/Title,AssignedTo/EMail,Created` +
          `&$expand=AssignedTo,Status,ServiceType` +
          `&$filter=${filter}` +
          `&$orderby=Created desc`,
        SPHttpClient.configurations.v1
      );

      const items = await response.json();
      const mappedItems = items.value.map((item: any) => ({
        id: item.Id,
        RequestID: item.RequestID,
        ServiceType: item.ServiceType
          ? isArabic
            ? item.ServiceType.Title_Ar
            : item.ServiceType.Title
          : "",
        Status: item.Status ? item.Status.Title : "", // ALWAYS English
        AssignedTo: item.AssignedTo ? item.AssignedTo.Title : "",
        AssignedToEmail: item.AssignedTo?.EMail || "",
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
    <Stack
      tokens={{ childrenGap: 10 }}
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

      {/* Table */}
      <DetailsList
        items={pagedItems.length > 0 ? pagedItems : [{} as IRequestItem]} // dummy row if no data
        columns={columns}
        selectionMode={SelectionMode.none}
        layoutMode={DetailsListLayoutMode.fixedColumns}
        isHeaderVisible={true}
        onRenderItemColumn={(item, index, column) => {
          if (!column) return null;

          // Show "No data found" if pagedItems is empty
          if (pagedItems.length === 0) {
            return column.key === "ServiceType" ? (
              <span style={{ fontStyle: "italic", color: "#666" }}>
                No data found
              </span>
            ) : null;
          }
          
          const value = item[column.fieldName as keyof IRequestItem];

          // Custom render for Status
          if (column.key === "Status")
            return <StatusBadge status={value as string} />;

          // Custom render for Actions
          if (column.key === "Actions")
            return (
              <ActionButton
                onClick={() =>
                  window.open(
                    `${siteUrl}/Lists/Requests/DispForm.aspx?ID=${item.id}`,
                    "_blank"
                  )
                }
              />
            );

          return <span>{value}</span>;
        }}
      />

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
