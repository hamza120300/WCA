import * as React from "react";
//import { Stack } from "@fluentui/react";

interface IStatusBadgeProps {
  status: string;
}

export const StatusBadge: React.FC<IStatusBadgeProps> = ({ status }) => {
  const colorMap: Record<string, string> = {
    Submitted: "#605E5C",
    Pending: "#FCE100",
    Completed: "#107C10",
    "Pending On Chief Approval": "#005A9E",
    "Pending On HR Approval": "#0078D4",
    "Pending On CEO Approval": "#8A8886",
    Approved: "#107C10",
    Reject: "#A4262C",
  };

  const color = colorMap[status] || "#666";

  return (
    <div
      style={{
        display: "inline-flex",
        justifyContent: "center",
        alignItems: "center",

        padding: "6px 12px",
        minWidth: 80,

        borderRadius: "50px", // full circle/rounded pill
        backgroundColor: color,
        color: "#fff",
        fontWeight: 600,
        fontSize: 12,
      }}
      className="circular-span"
    >
      {status}
    </div>
  );
};
