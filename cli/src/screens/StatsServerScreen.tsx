import React, { useState, useEffect } from "react";
import { Box, Text } from "ink";
import SelectInput from "ink-select-input";
import { CommandRunner } from "../components/CommandRunner.js";
import { ROOT_DIR, STATS_MANAGER_PATH } from "../utils/paths.js";
import { StatusBadge } from "../components/StatusBadge.js";

export type StatsServerScreenProps = {
  onBack: () => void;
};

export function StatsServerScreen({ onBack }: StatsServerScreenProps) {
  const [status, setStatus] = useState<"unknown" | "running" | "stopped">("unknown");
  const [action, setAction] = useState<"menu" | "starting" | "stopping" | "done">("menu");

  useEffect(() => {
    const check = () => {
      import("node:child_process").then(({ spawn }) => {
        const child = spawn("node", [STATS_MANAGER_PATH, "status"], { cwd: ROOT_DIR });
        let out = "";
        child.stdout?.on("data", (d) => { out += d.toString(); });
        child.on("close", () => {
          setStatus(out.includes("running") ? "running" : "stopped");
        });
      });
    };
    check();
  }, [action]);

  if (action === "starting") {
    return (
      <Box flexDirection="column">
        <CommandRunner
          cmd="node"
          args={[STATS_MANAGER_PATH, "start"]}
          cwd={ROOT_DIR}
          label="Starting stats server..."
          onComplete={() => setAction("done")}
        />
      </Box>
    );
  }

  if (action === "stopping") {
    return (
      <Box flexDirection="column">
        <CommandRunner
          cmd="node"
          args={[STATS_MANAGER_PATH, "stop"]}
          cwd={ROOT_DIR}
          label="Stopping stats server..."
          onComplete={() => setAction("done")}
        />
      </Box>
    );
  }

  if (action === "done") {
    return (
      <Box flexDirection="column">
        <StatusBadge status="ok" label="Action complete" />
        <Box marginTop={1}>
          <SelectInput
            items={[
              { label: "← Back to menu", value: "menu" },
              { label: "← Main menu", value: "main" },
            ]}
            onSelect={(item) => {
              if (item.value === "menu") setAction("menu");
              else onBack();
            }}
          />
        </Box>
      </Box>
    );
  }

  const items = [
    status === "stopped" && { label: "▶️  Start (silent)", value: "start" },
    status === "running" && { label: "⏹️  Stop", value: "stop" },
    { label: "🌐 Open Dashboard", value: "open" },
    { label: "← Back", value: "back" },
  ].filter(Boolean) as Array<{ label: string; value: string }>;

  return (
    <Box flexDirection="column">
      <Text bold>Stats Dashboard</Text>
      <Box marginTop={1}>
        {status === "running" ? (
          <StatusBadge status="ok" label="Server is running" />
        ) : status === "stopped" ? (
          <StatusBadge status="warn" label="Server is stopped" />
        ) : (
          <StatusBadge status="info" label="Checking status..." />
        )}
      </Box>
      <Box marginTop={1}>
        <Text color="gray">URL: http://127.0.0.1:9876</Text>
      </Box>
      <Box marginTop={1}>
        <SelectInput
          items={items}
          onSelect={(item) => {
            if (item.value === "back") onBack();
            else if (item.value === "start") setAction("starting");
            else if (item.value === "stop") setAction("stopping");
            else if (item.value === "open") {
              import("node:child_process").then(({ spawn }) => {
                spawn("open", ["http://127.0.0.1:9876"]);
              });
            }
          }}
        />
      </Box>
    </Box>
  );
}
