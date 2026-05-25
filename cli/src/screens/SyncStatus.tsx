import React, { useState, useEffect } from "react";
import { Box, Text } from "ink";
import SelectInput from "ink-select-input";
import { EXPORTS_DIR } from "../utils/paths.js";
import { StatusBadge } from "../components/StatusBadge.js";

export type SyncStatusProps = {
  onBack: () => void;
};

type BookStatus = {
  name: string;
  htmlCount: number;
  mdCount: number;
  jsonCount: number;
  assetCount: number;
  hasManifest: boolean;
};

export function SyncStatus({ onBack }: SyncStatusProps) {
  const [books, setBooks] = useState<BookStatus[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const fs = import("node:fs/promises");
    const path = import("node:path");
    Promise.all([fs, path]).then(async ([fsMod, pathMod]) => {
      const entries = await fsMod.readdir(EXPORTS_DIR, { withFileTypes: true });
      const dirs = entries.filter((e) => e.isDirectory()).map((e) => e.name);
      const results: BookStatus[] = [];
      for (const name of dirs) {
        const bookDir = pathMod.join(EXPORTS_DIR, name);
        const pagesDir = pathMod.join(bookDir, "pages");
        let htmlCount = 0;
        let mdCount = 0;
        let jsonCount = 0;
        let assetCount = 0;
        const hasManifest = await fsMod.access(pathMod.join(bookDir, "pages-manifest.json")).then(() => true).catch(() => false);
        try {
          const walk = async (dir: string) => {
            const items = await fsMod.readdir(dir, { withFileTypes: true });
            for (const item of items) {
              const p = pathMod.join(dir, item.name);
              if (item.isDirectory()) {
                if (item.name.endsWith(".assets")) assetCount++;
                else await walk(p);
              } else if (item.name.endsWith(".html")) htmlCount++;
              else if (item.name.endsWith(".md")) mdCount++;
              else if (item.name.endsWith(".json")) jsonCount++;
            }
          };
          await walk(pagesDir);
        } catch {
          // no pages dir
        }
        results.push({ name, htmlCount, mdCount, jsonCount, assetCount, hasManifest });
      }
      setBooks(results);
      setLoading(false);
    });
  }, []);

  if (loading) {
    return (
      <Box>
        <StatusBadge status="info" label="Scanning exports..." />
      </Box>
    );
  }

  if (books.length === 0) {
    return (
      <Box flexDirection="column">
        <Text color="gray">No exported notebooks found in exports/graph/</Text>
        <Box marginTop={1}>
          <SelectInput items={[{ label: "← Back", value: "back" }]} onSelect={onBack} />
        </Box>
      </Box>
    );
  }

  return (
    <Box flexDirection="column">
      <Text bold>Sync Status</Text>
      <Box flexDirection="column" marginTop={1}>
        {books.map((b) => (
          <Box key={b.name} flexDirection="column" marginBottom={1}>
            <Text bold>{b.name}</Text>
            <Box marginLeft={2}>
              <Text color="gray">
                HTML: <Text color="cyan">{b.htmlCount}</Text> | MD: <Text color="cyan">{b.mdCount}</Text> | Assets: <Text color="cyan">{b.assetCount}</Text>
              </Text>
            </Box>
            <Box marginLeft={2}>
              {b.hasManifest ? (
                <StatusBadge status="ok" label="Manifest present (ready for diff/resync)" />
              ) : (
                <StatusBadge status="warn" label="No manifest (run export first)" />
              )}
            </Box>
          </Box>
        ))}
      </Box>
      <Box marginTop={1}>
        <SelectInput items={[{ label: "← Back", value: "back" }]} onSelect={onBack} />
      </Box>
    </Box>
  );
}
