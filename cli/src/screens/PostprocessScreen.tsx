import React from "react";
import { Box, Text } from "ink";
import { CommandRunner } from "../components/CommandRunner.js";
import { ROOT_DIR } from "../utils/paths.js";
import SelectInput from "ink-select-input";

export type PostprocessScreenProps = {
  notebook: string;
  onBack: () => void;
};

export function PostprocessScreen({ notebook, onBack }: PostprocessScreenProps) {
  return (
    <Box flexDirection="column">
      <Text bold>Post-process — {notebook}</Text>
      <Box marginTop={1}>
        <CommandRunner
          cmd="node"
          args={["src/onenote-interactor.js", "graph-postprocess", "--root", `exports/graph/${notebook}`]}
          cwd={ROOT_DIR}
          label="Downloading assets & generating Markdown..."
          maxLines={30}
        />
      </Box>
      <Box marginTop={1}>
        <SelectInput items={[{ label: "← Back", value: "back" }]} onSelect={onBack} />
      </Box>
    </Box>
  );
}
