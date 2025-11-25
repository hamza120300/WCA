import * as React from "react";

interface IStatusBadgeProps {
  status: string;
  lang?: "en" | "ar"; // optional, default English
}

export const StatusBadge: React.FC<IStatusBadgeProps> = ({
  status,
  lang = "en",
}) => {
  // Figma COLORS
  const bgColors: Record<string, string> = {
    Pending: "#FFEBC0",
    Approved: "#CAF0CC",
    Reject: "#FED5D1",

    // Not provided → coordinated palette
    Submitted: "#E6E6E6", // soft gray
    Completed: "#D6EAF8", // soft blue
    "Pending On Chief Approval": "#FFEBC0",
    "Pending On HR Approval": "#FFEBC0",
    "Pending On CEO Approval": "#FFEBC0",
  };

  const textColors: Record<string, string> = {
    Pending: "#8F6200",
    Approved: "#437406",
    Reject: "#A0410D",

    // Coordinated with palettes
    Submitted: "#444444",
    Completed: "#1B4F72",
    "Pending On Chief Approval": "#8F6200",
    "Pending On HR Approval": "#8F6200",
    "Pending On CEO Approval": "#8F6200",
  };

  // Arabic translations
  const arMap: Record<string, string> = {
    Submitted: "مُقَدَّم",
    Pending: "قيد الانتظار",
    Completed: "اكتملت",
    "Pending On Chief Approval": "في انتظار موافقة الرئيس",
    "Pending On HR Approval": "في انتظار موافقة الموارد البشرية",
    "Pending On CEO Approval": "في انتظار موافقة الرئيس التنفيذي",
    Approved: "موافقه",
    Reject: "رفض",
  };

  const backgroundColor = bgColors[status] || "#E0E0E0";
  const textColor = textColors[status] || "#444";

  const label = lang === "ar" ? arMap[status] : status;

  return (
    <div
      style={{
        display: "inline-flex",
        justifyContent: "center",
        alignItems: "center",

        padding: "6px 12px",
        minWidth: 90,

        borderRadius: "50px",
        backgroundColor: backgroundColor,
        color: textColor,

        fontWeight: 700,
        fontSize: 12,

        //direction: lang === "ar" ? "rtl" : "ltr",
      }}
    >
      {label}
    </div>
  );
};
