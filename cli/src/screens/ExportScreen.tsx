import React from "react";
import { Box, Text } from "ink";
import { CommandRunner } from "../components/CommandRunner.js";
import { ROOT_DIR } from "../utils/paths.js";
import SelectInput from "ink-select-input";

export type ExportScreenProps = {
  notebook: string;
  onBack: () => void;
};

export function ExportScreen({ notebook, onBack }: ExportScreenProps) {
  const [mode, setMode] = React.useState<"menu" | "running">("menu");

  if (mode === "menu") {
    const items = [
      { label: `📤 Full Export — ${notebook}`, value: "full" },
      { label: "← Back", value: "back" },
    ];
    return (
      <Box flexDirection="column">
        <Text bold>Export Notebook</Text>
        <Box marginTop={1}>
          <SelectInput
            items={items}
            onSelect={(item) => {
              if (item.value === "back") onBack();
              else setMode("running");
            }}
          />
        </Box>
      </Box>
    );
  }

  return (
    <Box flexDirection="column">
      <CommandRunner
        cmd="node"
        args={[
          "src/onenote-interactor.js",
          "graph-export",
          "--notebook",
          notebook,
          "--out",
          `exports/graph/${notebook}`,
        ]}
        cwd={ROOT_DIR}
        label={`Exporting ${notebook}...`}
        maxLines={30}
        onComplete={(code) => {
          if (code !== 0) {
            setTimeout(() => setMode("menu"), 2000);
          }
        }}
      />
      <Box marginTop={1}>
        <SelectInput items={[{ label: "← Back", value: "back" }]} onSelect={onBack} />
      </Box>
    </Box>
  );
}
