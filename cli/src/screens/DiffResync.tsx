import React, { useState } from "react";
import { Box, Text } from "ink";
import SelectInput from "ink-select-input";
import { CommandRunner } from "../components/CommandRunner.js";
import { ROOT_DIR } from "../utils/paths.js";

type DiffSummary = {
  isQuick?: boolean;
  totals: {
    added: number;
    updated: number;
    deleted: number;
    missingLocal: number;
    unchanged: number;
    remotePages: number;
    localManifestPages: number;
    protectedSections: number;
    failedSections: number;
  };
  added: Array<{ title: string; sectionPath: string }>;
  updated: Array<{ title: string; sectionPath: string }>;
  deleted: Array<{ title: string; sectionPath: string }>;
  missingLocal: Array<{ title: string; sectionPath: string }>;
  protectedSections: Array<{ path: string }>;
  failedSections: Array<{ path: string }>;
};

type ResyncSummary = {
  diffReused?: boolean;
  diffGeneratedAt?: string;
  stats: {
    pagesExported: number;
    pagesFailed: number;
    deletedMarked: number;
  };
  cleanup: {
    deletedDetected: number;
    deletedCleaned: number;
    cleanupFailures: number;
  };
};

export type DiffResyncProps = {
  notebook: string;
  onBack: () => void;
};

export function DiffResync({ notebook, onBack }: DiffResyncProps) {
  const [step, setStep] = useState<"menu" | "diff" | "resync" | "done">("menu");
  const [mode, setMode] = useState<{ action: "diff" | "resync"; full: boolean } | null>(null);
  const [diffSummary, setDiffSummary] = useState<DiffSummary | null>(null);
  const [resyncSummary, setResyncSummary] = useState<ResyncSummary | null>(null);
  const [readError, setReadError] = useState("");
  const [forceFresh, setForceFresh] = useState(false);

  const loadSummaries = async () => {
    const fs = await import("node:fs/promises");
    const path = await import("node:path");
    const root = path.join(ROOT_DIR, "exports", "graph", notebook);

    try {
      const diffRaw = await fs.readFile(path.join(root, "diff-summary.json"), "utf-8");
      setDiffSummary(JSON.parse(diffRaw));
    } catch {
      // ignore
    }

    try {
      const resyncRaw = await fs.readFile(path.join(root, "resync-summary.json"), "utf-8");
      setResyncSummary(JSON.parse(resyncRaw));
    } catch {
      // ignore
    }
  };

  const handleComplete = (code: number | null) => {
    if (code === 0) {
      loadSummaries().then(() => setStep("done"));
    } else {
      setReadError(`Command failed with exit code ${code}`);
      setStep("done");
    }
  };

  if (step === "menu") {
    const items = [
      { label: "🔍 Quick Diff", value: "quick-diff" },
      { label: "🔄 Quick Resync", value: "quick-resync" },
      { label: "🔎 Full Diff (catches deletions)", value: "full-diff" },
      { label: "🔄 Full Resync (catches deletions)", value: "full-resync" },
      { label: "← Back", value: "back" },
    ];
    return (
      <Box flexDirection="column">
        <Text bold>Diff & Resync — {notebook}</Text>
        <Text color="gray">Quick = time-filtered scan (~30-60s). Full = scans all pages (~10-15min).</Text>
        <Box marginTop={1}>
          <SelectInput
            items={items}
            onSelect={(item) => {
              if (item.value === "back") onBack();
              else {
                const full = item.value.startsWith("full-");
                const action = item.value.endsWith("diff") ? "diff" : "resync";
                setMode({ action, full });
                setStep(action);
              }
            }}
          />
        </Box>
      </Box>
    );
  }

  if (step === "diff") {
    const args = [
      "src/onenote-interactor.js",
      "graph-diff",
      "--notebook",
      notebook,
      "--root",
      `exports/graph/${notebook}`,
    ];
    if (mode?.full) args.push("--full");
    return (
      <Box flexDirection="column">
        <CommandRunner
          cmd="node"
          args={args}
          cwd={ROOT_DIR}
          label={`Running ${mode?.full ? "full" : "quick"} diff on ${notebook}...`}
          onComplete={handleComplete}
        />
      </Box>
    );
  }

  if (step === "resync") {
    const args = [
      "src/onenote-interactor.js",
      "graph-resync",
      "--notebook",
      notebook,
      "--root",
      `exports/graph/${notebook}`,
    ];
    if (!forceFresh) args.push("--use-diff");
    if (mode?.full) args.push("--full");
    return (
      <Box flexDirection="column">
        <CommandRunner
          cmd="node"
          args={args}
          cwd={ROOT_DIR}
          label={`Running ${mode?.full ? "full" : "quick"} resync on ${notebook}${forceFresh ? " (fresh scan)" : ""}...`}
          onComplete={handleComplete}
        />
      </Box>
    );
  }

  return (
    <Box flexDirection="column">
      {readError ? (
        <Text color="red">✗ {readError}</Text>
      ) : (
        <>
          <Text bold color="green">
            ✓ {mode?.full ? "Full" : "Quick"} {mode?.action === "diff" ? "Diff" : "Resync"} complete
          </Text>

          {diffSummary && (
            <Box flexDirection="column" marginTop={1}>
              <Text bold color="cyan">Diff Results{diffSummary.isQuick ? " (Quick — deletions not scanned)" : ""}</Text>
              <Box marginLeft={2} flexDirection="column">
                <Text>
                  Remote pages: <Text color="cyan">{diffSummary.totals.remotePages}</Text>
                  {"  "}Local manifest: <Text color="cyan">{diffSummary.totals.localManifestPages}</Text>
                </Text>
                <Text>
                  Added: <Text color="green">{diffSummary.totals.added}</Text>
                  {"  "}Updated: <Text color="yellow">{diffSummary.totals.updated}</Text>
                  {"  "}Deleted: <Text color="red">{diffSummary.totals.deleted}</Text>
                  {"  "}Missing: <Text color="magenta">{diffSummary.totals.missingLocal}</Text>
                  {"  "}Unchanged: <Text color="gray">{diffSummary.totals.unchanged}</Text>
                </Text>
                <Text>
                  Protected sections: <Text color="yellow">{diffSummary.totals.protectedSections}</Text>
                  {"  "}Failed sections: <Text color="red">{diffSummary.totals.failedSections}</Text>
                </Text>
              </Box>

              {diffSummary.added.length > 0 && (
                <Box marginLeft={2} marginTop={1} flexDirection="column">
                  <Text bold color="green">Added ({diffSummary.added.length})</Text>
                  {diffSummary.added.slice(0, 20).map((p, i) => (
                    <Text key={`a-${i}`} color="gray">  • {p.sectionPath}/{p.title}</Text>
                  ))}
                  {diffSummary.added.length > 20 && (
                    <Text color="gray">  … and {diffSummary.added.length - 20} more</Text>
                  )}
                </Box>
              )}

              {diffSummary.updated.length > 0 && (
                <Box marginLeft={2} marginTop={1} flexDirection="column">
                  <Text bold color="yellow">Updated ({diffSummary.updated.length})</Text>
                  {diffSummary.updated.slice(0, 20).map((p, i) => (
                    <Text key={`u-${i}`} color="gray">  • {p.sectionPath}/{p.title}</Text>
                  ))}
                  {diffSummary.updated.length > 20 && (
                    <Text color="gray">  … and {diffSummary.updated.length - 20} more</Text>
                  )}
                </Box>
              )}

              {diffSummary.deleted.length > 0 && (
                <Box marginLeft={2} marginTop={1} flexDirection="column">
                  <Text bold color="red">Deleted ({diffSummary.deleted.length})</Text>
                  {diffSummary.deleted.slice(0, 20).map((p, i) => (
                    <Text key={`d-${i}`} color="gray">  • {p.sectionPath}/{p.title}</Text>
                  ))}
                  {diffSummary.deleted.length > 20 && (
                    <Text color="gray">  … and {diffSummary.deleted.length - 20} more</Text>
                  )}
                </Box>
              )}

              {diffSummary.missingLocal.length > 0 && (
                <Box marginLeft={2} marginTop={1} flexDirection="column">
                  <Text bold color="magenta">Missing Local ({diffSummary.missingLocal.length})</Text>
                  {diffSummary.missingLocal.slice(0, 20).map((p, i) => (
                    <Text key={`m-${i}`} color="gray">  • {p.sectionPath}/{p.title}</Text>
                  ))}
                  {diffSummary.missingLocal.length > 20 && (
                    <Text color="gray">  … and {diffSummary.missingLocal.length - 20} more</Text>
                  )}
                </Box>
              )}

              {diffSummary.protectedSections.length > 0 && (
                <Box marginLeft={2} marginTop={1} flexDirection="column">
                  <Text bold color="yellow">Protected Sections ({diffSummary.protectedSections.length})</Text>
                  {diffSummary.protectedSections.map((s, i) => (
                    <Text key={`p-${i}`} color="gray">  • {s.path}</Text>
                  ))}
                </Box>
              )}

              {diffSummary.failedSections.length > 0 && (
                <Box marginLeft={2} marginTop={1} flexDirection="column">
                  <Text bold color="red">Failed Sections ({diffSummary.failedSections.length})</Text>
                  {diffSummary.failedSections.map((s, i) => (
                    <Text key={`f-${i}`} color="gray">  • {s.path}</Text>
                  ))}
                </Box>
              )}
            </Box>
          )}

          {resyncSummary && (
            <Box flexDirection="column" marginTop={1}>
              <Text bold color="green">
                Resync Results
                {resyncSummary.diffReused
                  ? ` (diff reused from ${resyncSummary.diffGeneratedAt
                      ? new Date(resyncSummary.diffGeneratedAt).toLocaleString()
                      : "unknown time"})`
                  : ""}
              </Text>
              <Box marginLeft={2} flexDirection="column">
                <Text>
                  Pages exported: <Text color="green">{resyncSummary.stats.pagesExported}</Text>
                  {"  "}Pages failed: <Text color="red">{resyncSummary.stats.pagesFailed}</Text>
                  {"  "}Deleted marked: <Text color="yellow">{resyncSummary.stats.deletedMarked}</Text>
                </Text>
                <Text>
                  Cleanup — detected: <Text color="yellow">{resyncSummary.cleanup.deletedDetected}</Text>
                  {"  "}cleaned: <Text color="green">{resyncSummary.cleanup.deletedCleaned}</Text>
                  {"  "}failures: <Text color="red">{resyncSummary.cleanup.cleanupFailures}</Text>
                </Text>
              </Box>
            </Box>
          )}
        </>
      )}

      <Box marginTop={2}>
        <SelectInput
          items={
            resyncSummary?.diffReused
              ? [
                  { label: "🔄 Force fresh rescan (no diff reuse)", value: "fresh" },
                  { label: "← Back", value: "back" },
                ]
              : [{ label: "← Back", value: "back" }]
          }
          onSelect={(item) => {
            if (item.value === "fresh") {
              setForceFresh(true);
              setDiffSummary(null);
              setResyncSummary(null);
              setReadError("");
              setStep("resync");
            } else {
              onBack();
            }
          }}
        />
      </Box>
    </Box>
  );
}
