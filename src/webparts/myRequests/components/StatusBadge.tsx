import * as React from "react";
import { Stack, Text } from "@fluentui/react";

interface IStatusBadgeProps {
  status: string;
}

export const StatusBadge: React.FC<IStatusBadgeProps> = ({ status }) => {
  const colorMap: Record<string, string> = {
    Approved: "#107C10",
    Pending: "#FCE100",
    Rejected: "#A4262C",
    InProgress: "#005A9E",
  };

  const color = colorMap[status] || "#666";

  return (
    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
      <div
        style={{
          width: 10,
          height: 10,
          borderRadius: "50%",
          backgroundColor: color,
        }}
      />
      <Text>{status}</Text>
    </Stack>
  );
};
