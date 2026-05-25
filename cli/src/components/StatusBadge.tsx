import React from "react";
import { Text } from "ink";

export type StatusBadgeProps = {
  status: "ok" | "warn" | "error" | "info";
  label: string;
};

export function StatusBadge({ status, label }: StatusBadgeProps) {
  const colors = {
    ok: "green",
    warn: "yellow",
    error: "red",
    info: "cyan",
  };
  const symbols = {
    ok: "●",
    warn: "●",
    error: "●",
    info: "●",
  };
  return (
    <Text color={colors[status]}>
      {symbols[status]} {label}
    </Text>
  );
}
