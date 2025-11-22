import * as React from "react";

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

  // Darken hex color by a percentage
  const darkenColor = (hex: string, percent: number) => {
    const num = parseInt(hex.replace("#", ""), 16);
    const r = Math.max(0, (num >> 16) - 255 * percent);
    const g = Math.max(0, ((num >> 8) & 0x00ff) - 255 * percent);
    const b = Math.max(0, (num & 0x0000ff) - 255 * percent);
    return `rgb(${Math.round(r)}, ${Math.round(g)}, ${Math.round(b)})`;
  };

  const textColor = darkenColor(color, 0.3); // 30% darker for readability

  return (
    <div
      style={{
        display: "inline-flex",
        justifyContent: "center",
        alignItems: "center",

        padding: "6px 12px",
        minWidth: 80,

        borderRadius: "50px",
        backgroundColor: color, // status background
        color: textColor, // darker text for readability
        fontWeight: 700,
        fontSize: 12,
      }}
    >
      {status}
    </div>
  );
};
