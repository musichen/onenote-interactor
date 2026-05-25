import React from "react";
import { Box, Text } from "ink";
import SelectInput from "ink-select-input";

export type HelpScreenProps = {
  onBack: () => void;
};

export function HelpScreen({ onBack }: HelpScreenProps) {
  return (
    <Box flexDirection="column">
      <Text bold>OneNote Interactor CLI — Help</Text>
      <Box flexDirection="column" marginTop={1}>
        <Text bold color="cyan">Setup Commands</Text>
        <Box marginLeft={2} flexDirection="column">
          <Text>graph-login — Authenticate with Microsoft device-code flow</Text>
          <Text>graph-list — List notebooks from OneNote Graph API</Text>
        </Box>

        <Text bold color="cyan" marginTop={1}>Export Commands</Text>
        <Box marginLeft={2} flexDirection="column">
          <Text>graph-export — Export notebook pages as HTML + JSON</Text>
          <Text>graph-postprocess — Download assets, rewrite HTML, generate Markdown</Text>
          <Text>graph-sync — Run export + postprocess in one shot</Text>
        </Box>

        <Text bold color="cyan" marginTop={1}>Maintenance Commands</Text>
        <Box marginLeft={2} flexDirection="column">
          <Text>graph-manifest — Build local page manifest from exported files</Text>
          <Text>graph-diff — Compare local state with Graph (detects changes)</Text>
          <Text>graph-resync — Incremental update: export changed pages + cleanup deleted</Text>
          <Text>graph-audit — Quality check for malformed HTML, missing assets, etc.</Text>
          <Text>graph-status — Disk counts and export summary</Text>
        </Box>

        <Text bold color="cyan" marginTop={1}>Utility Commands</Text>
        <Box marginLeft={2} flexDirection="column">
          <Text>create-markdown-only-book — Copy clean .md files with stripped IDs</Text>
          <Text>stats-server — Progress dashboard on http://127.0.0.1:9876</Text>
          <Text>local-index — Scan Mac OneNote backup .one files</Text>
        </Box>
      </Box>
      <Box marginTop={1}>
        <SelectInput items={[{ label: "← Back", value: "back" }]} onSelect={onBack} />
      </Box>
    </Box>
  );
}
