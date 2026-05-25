import React, { useState, useEffect } from "react";
import { Box, Text } from "ink";
import SelectInput from "ink-select-input";
import Spinner from "ink-spinner";
import { useCommand } from "../hooks/useCommand.js";
import { EXPORTS_DIR, ROOT_DIR } from "../utils/paths.js";

export type NotebookListProps = {
  onSelect: (name: string) => void;
  onBack: () => void;
};

export function NotebookList({ onSelect, onBack }: NotebookListProps) {
  const { result, run } = useCommand();
  const [localBooks, setLocalBooks] = useState<string[]>([]);
  const [remoteBooks, setRemoteBooks] = useState<string[]>([]);
  const [loadingLocal, setLoadingLocal] = useState(true);
  const [fetchError, setFetchError] = useState("");

  // 1. Instantly load local books
  useEffect(() => {
    import("node:fs/promises").then(async (fs) => {
      try {
        const entries = await fs.readdir(EXPORTS_DIR, { withFileTypes: true });
        const dirs = entries.filter((e) => e.isDirectory()).map((e) => e.name);
        setLocalBooks(dirs);
      } catch {
        setLocalBooks([]);
      }
      setLoadingLocal(false);
    });
  }, []);

  // 2. Background fetch from Graph
  useEffect(() => {
    run("node", ["src/onenote-interactor.js", "graph-list-all"], { cwd: ROOT_DIR });
  }, []);

  useEffect(() => {
    if (result.status === "success") {
      const allOutput = result.output.join("\n");
      const jsonStart = allOutput.indexOf("{");
      if (jsonStart !== -1) {
        try {
          const data = JSON.parse(allOutput.slice(jsonStart));
          const names = (data.notebooks || []).map((nb: any) => nb.displayName as string);
          setRemoteBooks(names);
        } catch {
          // ignore parse error
        }
      }
    }
    if (result.status === "error") {
      setFetchError("Graph API rate-limited. Showing locally known notebooks.");
    }
  }, [result.status]);

  // Merge: all remote + any local-only books
  const allBooks = Array.from(new Set([...remoteBooks, ...localBooks]));

  const items = allBooks.map((b) => {
    const isRemote = remoteBooks.includes(b);
    const isLocal = localBooks.includes(b);
    let icon = "📓";
    if (isRemote && isLocal) icon = "📓";
    else if (isLocal) icon = "💾";
    else if (isRemote) icon = "☁️";
    return { label: `${icon} ${b}`, value: b };
  });
  items.push({ label: "← Back", value: "__back__" });

  if (loadingLocal) {
    return (
      <Box>
        <Text color="cyan">
          <Spinner type="dots" /> Loading...
        </Text>
      </Box>
    );
  }

  return (
    <Box flexDirection="column">
      <Text bold>Select a notebook:</Text>
      {fetchError && (
        <Text color="yellow" dimColor>
          {fetchError}
        </Text>
      )}
      {result.status === "running" && (
        <Text color="gray" dimColor>
          <Spinner type="dots" /> Refreshing from Graph...
        </Text>
      )}
      <Box marginTop={1}>
        <SelectInput
          items={items}
          onSelect={(item) => {
            if (item.value === "__back__") onBack();
            else onSelect(item.value);
          }}
        />
      </Box>
    </Box>
  );
}
