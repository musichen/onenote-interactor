import React, { useState, useEffect } from "react";
import { Box, Text } from "ink";
import SelectInput from "ink-select-input";
import { CommandRunner } from "../components/CommandRunner.js";
import { EXPORTS_DIR, MARKDOWN_EXPORTER_PATH, ROOT_DIR } from "../utils/paths.js";

export type MarkdownExportScreenProps = {
  onBack: () => void;
};

export function MarkdownExportScreen({ onBack }: MarkdownExportScreenProps) {
  const [books, setBooks] = useState<string[]>([]);
  const [loading, setLoading] = useState(true);
  const [selected, setSelected] = useState<string | null>(null);

  useEffect(() => {
    import("node:fs/promises").then(async (fs) => {
      try {
        const entries = await fs.readdir(EXPORTS_DIR, { withFileTypes: true });
        const dirs = entries.filter((e) => e.isDirectory()).map((e) => e.name);
        setBooks(dirs);
      } catch {
        setBooks([]);
      }
      setLoading(false);
    });
  }, []);

  if (loading) {
    return (
      <Box>
        <Text color="cyan">Loading available books...</Text>
      </Box>
    );
  }

  if (selected) {
    return (
      <Box flexDirection="column">
        <CommandRunner
          cmd="node"
          args={[MARKDOWN_EXPORTER_PATH]}
          cwd={ROOT_DIR}
          env={{ FORCE_BOOK: selected }}
          label={`Creating markdown-only copy of ${selected}...`}
          maxLines={20}
        />
        <Box marginTop={1}>
          <SelectInput items={[{ label: "← Back", value: "back" }]} onSelect={onBack} />
        </Box>
      </Box>
    );
  }

  if (books.length === 0) {
    return (
      <Box flexDirection="column">
        <Text color="gray">No exported books found. Run export first.</Text>
        <Box marginTop={1}>
          <SelectInput items={[{ label: "← Back", value: "back" }]} onSelect={onBack} />
        </Box>
      </Box>
    );
  }

  const items = books.map((b) => ({ label: `📓 ${b}`, value: b }));
  items.push({ label: "← Back", value: "__back__" });

  return (
    <Box flexDirection="column">
      <Text bold>Create Markdown-Only Copy</Text>
      <Text color="gray">Select a book to export clean markdown files:</Text>
      <Box marginTop={1}>
        <SelectInput
          items={items}
          onSelect={(item) => {
            if (item.value === "__back__") onBack();
            else setSelected(item.value);
          }}
        />
      </Box>
    </Box>
  );
}
