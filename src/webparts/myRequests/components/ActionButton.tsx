import * as React from "react";
import { IconButton, IIconProps } from "@fluentui/react";

interface IActionButtonProps {
  onClick?: () => void;
}

// Use a Fluent UI eye icon
const viewIcon: IIconProps = { iconName: "View" };

export const ActionButton: React.FC<IActionButtonProps> = ({ onClick }) => {
  return <IconButton iconProps={viewIcon} title="View" onClick={onClick} />;
};
