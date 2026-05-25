import { describe, it, beforeEach, afterEach } from "node:test";
import assert from "node:assert";
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import {
  cleanupDeletedPages,
  buildLocalPagesManifest,
  loadLocalManifest,
  writeLocalManifest,
} from "../src/onenote-interactor.js";

describe("cleanupDeletedPages", () => {
  let tmpDir;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "onenote-cleanup-test-"));
  });

  afterEach(async () => {
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  it("deletes HTML, JSON, MD and asset files for deleted pages", async () => {
    const sectionDir = path.join(tmpDir, "pages", "SecA");
    await fs.mkdir(sectionDir, { recursive: true });

    const htmlPath = path.join(sectionDir, "Deleted-page-1.html");
    const jsonPath = path.join(sectionDir, "Deleted-page-1.json");
    const mdPath = path.join(sectionDir, "Deleted-page-1.md");
    const assetDir = path.join(sectionDir, "Deleted-page-1.assets");

    await fs.writeFile(htmlPath, "<html></html>");
    await fs.writeFile(jsonPath, "{}");
    await fs.writeFile(mdPath, "# Deleted");
    await fs.mkdir(assetDir, { recursive: true });
    await fs.writeFile(path.join(assetDir, "img.png"), "fake-image");

    const manifest = {
      pages: {
        "page-1": {
          id: "page-1",
          title: "Deleted",
          sectionPath: "SecA",
          htmlPath: path.relative(tmpDir, htmlPath),
          jsonPath: path.relative(tmpDir, jsonPath),
          markdownPath: path.relative(tmpDir, mdPath),
          assetDir: path.relative(tmpDir, assetDir),
        },
      },
    };

    const deleted = [
      { id: "page-1", title: "Deleted", sectionPath: "SecA", htmlPath: path.relative(tmpDir, htmlPath), jsonPath: path.relative(tmpDir, jsonPath), markdownPath: path.relative(tmpDir, mdPath), assetDir: path.relative(tmpDir, assetDir) },
    ];

    const result = await cleanupDeletedPages(tmpDir, manifest, deleted);

    assert.strictEqual(result.cleaned.length, 1);
    assert.strictEqual(result.failed.length, 0);
    assert.strictEqual(manifest.pages["page-1"], undefined);

    // Verify files were deleted
    const remaining = await fs.readdir(sectionDir);
    assert.strictEqual(remaining.length, 0);
  });

  it("does nothing when no deleted pages", async () => {
    const manifest = { pages: {} };
    const result = await cleanupDeletedPages(tmpDir, manifest, []);
    assert.strictEqual(result.cleaned.length, 0);
    assert.strictEqual(result.failed.length, 0);
  });

  it("handles missing files gracefully", async () => {
    const manifest = {
      pages: {
        "page-1": {
          id: "page-1",
          title: "Ghost",
          sectionPath: "SecA",
          htmlPath: "pages/SecA/Ghost-page-1.html",
          jsonPath: "pages/SecA/Ghost-page-1.json",
        },
      },
    };

    const deleted = [{ id: "page-1", title: "Ghost", sectionPath: "SecA" }];
    const result = await cleanupDeletedPages(tmpDir, manifest, deleted);

    assert.strictEqual(result.cleaned.length, 1);
    assert.strictEqual(result.failed.length, 0);
    assert.strictEqual(manifest.pages["page-1"], undefined);
  });
});

describe("buildLocalPagesManifest", () => {
  let tmpDir;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "onenote-manifest-test-"));
  });

  afterEach(async () => {
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  it("builds manifest from existing files", async () => {
    const sectionDir = path.join(tmpDir, "pages", "SecA");
    await fs.mkdir(sectionDir, { recursive: true });

    // Write _section.json with page metadata
    await fs.writeFile(
      path.join(sectionDir, "_section.json"),
      JSON.stringify({
        pages: [
          { id: "page-1", title: "Test Page", lastModifiedDateTime: "2024-01-15T10:00:00Z", createdDateTime: "2024-01-10T08:00:00Z" },
        ],
      })
    );

    // Write corresponding HTML and JSON files (naming: {title}-{id}.html)
    await fs.writeFile(path.join(sectionDir, "Test Page-page-1.html"), "<html></html>");
    await fs.writeFile(path.join(sectionDir, "Test Page-page-1.json"), "{}");

    const manifest = await buildLocalPagesManifest({ root: tmpDir });

    assert.ok(manifest.generatedAt);
    assert.strictEqual(Object.keys(manifest.pages).length, 1);
    assert.strictEqual(manifest.pages["page-1"].title, "Test Page");
    assert.ok(manifest.pages["page-1"].htmlPath);
    assert.ok(manifest.pages["page-1"].jsonPath);
  });

  it("ignores pages missing HTML or JSON", async () => {
    const sectionDir = path.join(tmpDir, "pages", "SecA");
    await fs.mkdir(sectionDir, { recursive: true });

    await fs.writeFile(
      path.join(sectionDir, "_section.json"),
      JSON.stringify({
        pages: [
          { id: "page-1", title: "Has Both", lastModifiedDateTime: "2024-01-15T10:00:00Z", createdDateTime: "2024-01-10T08:00:00Z" },
          { id: "page-2", title: "Missing HTML", lastModifiedDateTime: "2024-01-15T10:00:00Z", createdDateTime: "2024-01-10T08:00:00Z" },
        ],
      })
    );

    await fs.writeFile(path.join(sectionDir, "Has Both-page-1.html"), "<html></html>");
    await fs.writeFile(path.join(sectionDir, "Has Both-page-1.json"), "{}");
    // Only JSON for page-2 (file naming: {title}-{id})
    await fs.writeFile(path.join(sectionDir, "Missing HTML-page-2.json"), "{}");

    const manifest = await buildLocalPagesManifest({ root: tmpDir });
    assert.strictEqual(Object.keys(manifest.pages).length, 1);
    assert.strictEqual(manifest.pages["page-1"].title, "Has Both");
    assert.strictEqual(manifest.pages["page-2"], undefined);
  });
});

describe("loadLocalManifest / writeLocalManifest", () => {
  let tmpDir;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "onenote-manifest-io-test-"));
  });

  afterEach(async () => {
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  it("returns empty manifest when file does not exist", async () => {
    const manifest = await loadLocalManifest(tmpDir);
    assert.strictEqual(manifest.generatedAt, null);
    assert.deepStrictEqual(manifest.pages, {});
    assert.strictEqual(manifest.rootDir, tmpDir);
  });

  it("round-trips manifest through write and load", async () => {
    const manifest = {
      pages: {
        "page-1": { id: "page-1", title: "Test" },
      },
    };

    await writeLocalManifest(tmpDir, manifest);
    const loaded = await loadLocalManifest(tmpDir);

    assert.strictEqual(loaded.pages["page-1"].title, "Test");
    assert.ok(loaded.generatedAt);
    assert.strictEqual(loaded.rootDir, tmpDir);
  });
});
