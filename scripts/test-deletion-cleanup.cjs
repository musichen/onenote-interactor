#!/usr/bin/env node
/**
 * test-deletion-cleanup.cjs
 * Targeted test of the deletion cleanup logic for notebook "A".
 * Loads the manifest, finds pages no longer present remotely (from diff-summary),
 * deletes their local artifacts, updates manifest, and writes a session log.
 */

const fs = require("fs").promises;
const path = require("path");

const ROOT_DIR = path.join(process.cwd(), "exports", "graph", "A");
const MANIFEST_PATH = path.join(ROOT_DIR, "pages-manifest.json");
const LOGS_DIR = path.join(process.cwd(), "logs");

async function pathExists(p) {
  try {
    await fs.access(p);
    return true;
  } catch {
    return false;
  }
}

async function ensureDir(dir) {
  await fs.mkdir(dir, { recursive: true });
}

function nowIso() {
  return new Date().toISOString();
}

async function main() {
  console.log("=== Deletion Cleanup Test for Notebook 'A' ===\n");

  // Load manifest
  const manifestRaw = await fs.readFile(MANIFEST_PATH, "utf-8");
  const manifest = JSON.parse(manifestRaw);

  // Load latest diff summary
  const diffFiles = (await fs.readdir(LOGS_DIR))
    .filter(f => f.startsWith("diff-A-") && f.endsWith(".md"))
    .sort();
  const latestDiffFile = diffFiles[diffFiles.length - 1];
  console.log(`Using latest diff log: ${latestDiffFile}`);

  // Parse deleted pages from the markdown table
  const diffContent = await fs.readFile(path.join(LOGS_DIR, latestDiffFile), "utf-8");
  const deletedSection = diffContent.match(/## Deleted Pages \((\d+)\)\n\n([\s\S]*?)(?=\n## |$)/);
  if (!deletedSection) {
    console.log("No deleted pages section found in diff log.");
    process.exit(0);
  }

  const deletedCount = parseInt(deletedSection[1], 10);
  const tableLines = deletedSection[2].trim().split("\n").slice(2); // skip header + separator
  const deletedPages = [];
  for (const line of tableLines) {
    const cells = line.split("|").map(c => c.trim()).filter(Boolean);
    if (cells.length >= 3) {
      deletedPages.push({
        title: cells[0],
        sectionPath: cells[1],
        lastModifiedDateTime: cells[2]
      });
    }
  }

  console.log(`Found ${deletedPages.length} deleted page(s) in diff log:\n`);

  const cleaned = [];
  const failed = [];

  for (const dp of deletedPages) {
    // Find the matching manifest entry by title + sectionPath
    const entry = Object.values(manifest.pages).find(
      p => p.title === dp.title && p.sectionPath === dp.sectionPath
    );
    if (!entry) {
      console.log(`  ⚠️  Manifest entry not found for: ${dp.sectionPath} :: ${dp.title}`);
      continue;
    }

    console.log(`  Cleaning: ${dp.sectionPath} :: ${dp.title}`);

    const pathsToDelete = [];
    if (entry.htmlPath) pathsToDelete.push(path.join(ROOT_DIR, entry.htmlPath));
    if (entry.jsonPath) pathsToDelete.push(path.join(ROOT_DIR, entry.jsonPath));
    if (entry.markdownPath) pathsToDelete.push(path.join(ROOT_DIR, entry.markdownPath));
    if (entry.assetDir) pathsToDelete.push(path.join(ROOT_DIR, entry.assetDir));

    let anyDeleted = false;
    for (const p of pathsToDelete) {
      try {
        if (await pathExists(p)) {
          const stat = await fs.stat(p);
          if (stat.isDirectory()) {
            await fs.rm(p, { recursive: true, force: true });
          } else {
            await fs.unlink(p);
          }
          anyDeleted = true;
          console.log(`    ✅ Deleted: ${path.relative(ROOT_DIR, p)}`);
        } else {
          console.log(`    ℹ️  Already gone: ${path.relative(ROOT_DIR, p)}`);
        }
      } catch (err) {
        console.error(`    ❌ Failed to delete ${path.relative(ROOT_DIR, p)}: ${err.message}`);
        failed.push({ path: p, error: err.message });
      }
    }

    if (failed.length === 0 || anyDeleted) {
      delete manifest.pages[entry.id];
      cleaned.push({ ...entry, sectionPath: dp.sectionPath });
    }
  }

  // Write updated manifest
  manifest.generatedAt = nowIso();
  await fs.writeFile(MANIFEST_PATH, JSON.stringify(manifest, null, 2), "utf-8");
  console.log(`\n📝 Manifest updated. Remaining pages: ${Object.keys(manifest.pages).length}`);

  // Write session log
  await ensureDir(LOGS_DIR);
  const ts = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
  const logPath = path.join(LOGS_DIR, `session-A-${ts}.md`);

  const lines = [];
  lines.push(`# Export Session Log — A (Deletion Cleanup Test)`);
  lines.push("");
  lines.push(`- **Notebook:** \`A\``);
  lines.push(`- **Root:** \`${ROOT_DIR}\``);
  lines.push(`- **Generated:** ${nowIso()}`);
  lines.push("");
  lines.push("## Summary");
  lines.push("");
  lines.push("| Metric | Value |");
  lines.push("|--------|-------|");
  lines.push(`| Deleted pages detected | ${deletedCount} |`);
  lines.push(`| Deleted pages cleaned | ${cleaned.length} |`);
  lines.push(`| Cleanup failures | ${failed.length} |`);
  lines.push("");

  if (cleaned.length > 0) {
    lines.push(`## Deleted Pages Cleaned (${cleaned.length})`);
    lines.push("");
    lines.push("| Title | Section Path | HTML Path |");
    lines.push("|-------|--------------|-----------|");
    for (const p of cleaned) {
      lines.push(`| ${(p.title || "Untitled").replace(/\|/g, "\\|")} | ${(p.sectionPath || "-").replace(/\|/g, "\\|")} | ${(p.htmlPath || "-").replace(/\|/g, "\\|")} |`);
    }
    lines.push("");
  }

  if (failed.length > 0) {
    lines.push(`## Cleanup Failures (${failed.length})`);
    lines.push("");
    lines.push("| Path | Error |");
    lines.push("|------|-------|");
    for (const f of failed) {
      lines.push(`| ${f.path.replace(/\|/g, "\\|")} | ${(f.error || "-").replace(/\|/g, "\\|")} |`);
    }
    lines.push("");
  }

  lines.push("---");
  lines.push("*Deletion cleanup test — manual run*");

  await fs.writeFile(logPath, lines.join("\n"), "utf-8");
  console.log(`📝 Session log written: ${logPath}`);

  console.log("\n=== Test Complete ===");
  console.log(`  Cleaned: ${cleaned.length}`);
  console.log(`  Failed:  ${failed.length}`);
}

main().catch(err => {
  console.error("Fatal error:", err);
  process.exit(1);
});
