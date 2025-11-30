import * as React from "react";

interface IStatusBadgeProps {
  status: string; // English key
  lang?: "en" | "ar";
}

export const StatusBadge: React.FC<IStatusBadgeProps> = ({
  status,
  lang = "en",
}) => {
  // Background colors for English keys
  const bgColors: Record<string, string> = {
    Pending: "#FFEBC0",
    Approved: "#CAF0CC",
    Reject: "#FED5D1",
    Submitted: "#E6E6E6",
    Completed: "#D6EAF8",
    "Pending Chief Approval": "#FFEBC0",
    "Pending HR Approval": "#FFEBC0",
    "Pending CEO Approval": "#FFEBC0",
    "Pending Manager Approval": "#FFEBC0",
    "Pending IT Approval": "#FFEBC0",
  };

  // Text colors for English keys
  const textColors: Record<string, string> = {
    Pending: "#8F6200",
    Approved: "#437406",
    Reject: "#A0410D",
    Submitted: "#444444",
    Completed: "#1B4F72",
    "Pending Chief Approval": "#8F6200",
    "Pending HR Approval": "#8F6200",
    "Pending CEO Approval": "#8F6200",
    "Pending IT Approval": "#8F6200",
    "Pending Manager Approval": "#8F6200",
  };

  // Arabic display mapping
  const arMap: Record<string, string> = {
    Submitted: "مقدم",
    Pending: "قيد الانتظار",
    Completed: "اكتملت",
    "Pending Chief Approval": "في انتظار موافقة الرئيس",
    "Pending HR Approval": "في انتظار موافقة الموارد البشرية",
    "Pending CEO Approval": "في انتظار موافقة الرئيس التنفيذي",
    "Pending Manager Approval": "في انتظار موافقة المدير المباشر",
    "Pending IT Approval": "في انتظار موافقة قسم تكنولوجيا المعلومات",
    Approved: "موافقة",
    Reject: "رفض",
  };

  // Colors always use English keys
  const backgroundColor = bgColors[status] || "#E0E0E0";
  const textColor = textColors[status] || "#444";

  // Label for display only
  const label = lang === "ar" ? arMap[status] || status : status;

  return (
    <div
      style={{
        display: "inline-flex",
        justifyContent: "center",
        alignItems: "center",
        padding: "6px 12px",
        minWidth: lang === "ar" ? 110 : 90, // wider for Arabic
        borderRadius: "50px",
        backgroundColor,
        color: textColor,
        fontWeight: 700,
        fontSize: 12,
        direction: lang === "ar" ? "rtl" : "ltr",
      }}
    >
      {label}
    </div>
  );
};
