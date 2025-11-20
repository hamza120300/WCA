import * as React from "react";
import Button from "@mui/material/Button";
import VisibilityIcon from "@mui/icons-material/Visibility";

// interface IActionButtonProps {
//   requestId: number | string;
// }

const ActionButton: React.FC = () => {
  const url = `https://ejadasharepoint.sharepoint.com/sites/WCA-DEV`;

  return (
    <Button
      variant="contained"
      size="small"
      onClick={() => window.open(url, "_blank")}
      startIcon={<VisibilityIcon />}
    >
      View
    </Button>
  );
};

export default ActionButton;
