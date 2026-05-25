import React from "react";
import { Box, Text } from "ink";
import SelectInput from "ink-select-input";
import type { Screen } from "../types.js";

export type MainMenuProps = {
  onSelect: (screen: Screen) => void;
  onSelectNotebook: (name: string) => void;
  activeNotebook?: string;
};

export function MainMenu({ onSelect, activeNotebook }: MainMenuProps) {
  const items = [
    { label: "📚 List OneNote Notebooks", value: "notebook-list" as Screen },
    { label: "📊 Check Sync Status", value: "sync-status" as Screen },
    { label: "🔄 Diff & Resync", value: "diff-resync" as Screen },
    { label: "📤 Export Notebook", value: "export" as Screen },
    { label: "🔧 Post-process Export", value: "postprocess" as Screen },
    { label: "📄 Create Markdown-Only Copy", value: "markdown-export" as Screen },
    { label: "📈 Stats Dashboard", value: "stats-server" as Screen },
    { label: "❓ Help", value: "help" as Screen },
    { label: "👋 Quit", value: "quit" as Screen },
  ];

  return (
    <Box flexDirection="column" marginTop={1}>
      {activeNotebook ? (
        <Text color="gray">
          Active notebook: <Text color="yellow">{activeNotebook}</Text>
        </Text>
      ) : (
        <Text color="yellow">No notebook selected — choose "List OneNote Notebooks" first</Text>
      )}
      <Box marginTop={1}>
        <SelectInput
          items={items}
          onSelect={(item) => {
            if (item.value === "quit") {
              process.exit(0);
            } else {
              onSelect(item.value);
            }
          }}
        />
      </Box>
    </Box>
  );
}
