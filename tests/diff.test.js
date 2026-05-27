import { describe, it, beforeEach, afterEach } from "node:test";
import assert from "node:assert";
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import {
  diffGraphNotebook,
  graphStatus,
  setAccessTokenProvider,
  resetAccessTokenProvider,
  scanRemotePages,
  graphFetchAllJsonItems,
  graphFetchJson,
} from "../src/onenote-interactor.js";

describe("scanRemotePages", () => {
  let originalFetch;

  beforeEach(() => {
    originalFetch = global.fetch;
  });

  afterEach(() => {
    global.fetch = originalFetch;
  });

  it("returns all pages from sections", async () => {
    global.fetch = async (url) => {
      return {
        ok: true,
        status: 200,
        headers: { get: () => null },
        json: async () => ({
          value: [
            { id: "page-1", title: "Page One", lastModifiedDateTime: "2024-01-15T10:00:00Z", createdDateTime: "2024-01-10T08:00:00Z" },
          ],
        }),
      };
    };

    const structure = {
      sections: [{ name: "SecA", id: "sec-1", type: "section" }],
      sectionGroups: [],
    };

    const result = await scanRemotePages("fake-token", structure);
    assert.strictEqual(Object.keys(result.pages).length, 1);
    assert.strictEqual(result.pages["page-1"].title, "Page One");
    assert.strictEqual(result.pages["page-1"].sectionPath, "SecA");
    assert.strictEqual(result.protectedSections.length, 0);
    assert.strictEqual(result.failedSections.length, 0);
    assert.strictEqual(result.isQuick, false);
  });

  it("skips unchanged sections in quick mode", async () => {
    const callLog = [];
    global.fetch = async (url) => {
      callLog.push(url);
      return {
        ok: true,
        status: 200,
        headers: { get: () => null },
        json: async () => ({
          value: [
            { id: "page-1", title: "Page One", lastModifiedDateTime: "2024-01-10T08:00:00Z", createdDateTime: "2024-01-10T08:00:00Z" },
          ],
        }),
      };
    };

    const structure = {
      sections: [{ name: "SecA", id: "sec-1", type: "section" }],
      sectionGroups: [],
    };

    const cutoff = new Date("2024-01-15T00:00:00Z").getTime();
    const result = await scanRemotePages("fake-token", structure, { cutoffTimestamp: cutoff });
    assert.strictEqual(Object.keys(result.pages).length, 0);
    assert.strictEqual(result.isQuick, true);
    // Should only make one $top=1 call, not a full scan
    assert.strictEqual(callLog.length, 1);
    assert.ok(callLog[0].includes("$top=1"));
  });

  it("scans changed sections in quick mode", async () => {
    const callLog = [];
    global.fetch = async (url) => {
      callLog.push(url);
      if (url.includes("$top=1")) {
        return {
          ok: true,
          status: 200,
          headers: { get: () => null },
          json: async () => ({
            value: [
              { id: "page-1", title: "Page One", lastModifiedDateTime: "2024-01-20T08:00:00Z", createdDateTime: "2024-01-10T08:00:00Z" },
            ],
          }),
        };
      }
      return {
        ok: true,
        status: 200,
        headers: { get: () => null },
        json: async () => ({
          value: [
            { id: "page-1", title: "Page One", lastModifiedDateTime: "2024-01-20T08:00:00Z", createdDateTime: "2024-01-10T08:00:00Z" },
            { id: "page-2", title: "Page Two", lastModifiedDateTime: "2024-01-18T08:00:00Z", createdDateTime: "2024-01-10T08:00:00Z" },
          ],
        }),
      };
    };

    const structure = {
      sections: [{ name: "SecA", id: "sec-1", type: "section" }],
      sectionGroups: [],
    };

    const cutoff = new Date("2024-01-15T00:00:00Z").getTime();
    const result = await scanRemotePages("fake-token", structure, { cutoffTimestamp: cutoff });
    assert.ok(Object.keys(result.pages).length >= 1);
    assert.strictEqual(result.isQuick, true);
    // Should make at least 2 calls: $top=1 check + full scan
    assert.ok(callLog.length >= 2);
  });

  it("tracks protected sections on 403", async () => {
    global.fetch = async () => {
      const err = new Error("403 Forbidden");
      err.message = "Graph request failed: 403 Forbidden";
      throw err;
    };

    const structure = {
      sections: [{ name: "SecA", id: "sec-1", type: "section" }],
      sectionGroups: [],
    };

    const result = await scanRemotePages("fake-token", structure);
    assert.strictEqual(Object.keys(result.pages).length, 0);
    assert.strictEqual(result.protectedSections.length, 1);
    assert.strictEqual(result.protectedSections[0].path, "SecA");
  });

  it("tracks failed sections on other errors", async () => {
    global.fetch = async () => {
      throw new Error("Network error");
    };

    const structure = {
      sections: [{ name: "SecA", id: "sec-1", type: "section" }],
      sectionGroups: [],
    };

    const result = await scanRemotePages("fake-token", structure);
    assert.strictEqual(Object.keys(result.pages).length, 0);
    assert.strictEqual(result.failedSections.length, 1);
    assert.strictEqual(result.failedSections[0].path, "SecA");
  });
});

describe("diffGraphNotebook", () => {
  // These tests require mocking the MSAL auth flow (getAccessToken).
  // For now we test diff logic indirectly via scanRemotePages tests above.
  beforeEach(() => {
    setAccessTokenProvider(async () => "fake-token");
  });

  afterEach(() => {
    resetAccessTokenProvider();
    global.fetch = undefined;
  });

  it("detects moved pages and reports previous section path", async () => {
    const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "onenote-diff-test-"));
    try {
      const notebook = { id: "notebook-1", displayName: "A" };
      const structure = {
        sections: [
          { name: "NewSec", id: "sec-1", type: "section" }
        ],
        sectionGroups: []
      };
      await fs.mkdir(path.join(tmpDir, "pages", "OldSec"), { recursive: true });
      await fs.writeFile(path.join(tmpDir, "pages", "OldSec", "moved-page-1.html"), "<html></html>");
      await fs.writeFile(path.join(tmpDir, "pages", "OldSec", "moved-page-1.json"), JSON.stringify({ id: "page-1", title: "Moved Page", createdDateTime: "2024-01-10T08:00:00Z", lastModifiedDateTime: "2024-01-10T08:00:00Z" }));
      await fs.writeFile(path.join(tmpDir, "notebook.json"), JSON.stringify(notebook));
      await fs.writeFile(path.join(tmpDir, "structure.json"), JSON.stringify(structure));
      await fs.writeFile(path.join(tmpDir, "pages-manifest.json"), JSON.stringify({
        generatedAt: new Date().toISOString(),
        rootDir: tmpDir,
        pages: {
          "page-1": {
            id: "page-1",
            title: "Moved Page",
            sectionPath: "OldSec",
            htmlPath: path.join("pages", "OldSec", "moved-page-1.html"),
            jsonPath: path.join("pages", "OldSec", "moved-page-1.json"),
            createdDateTime: "2024-01-10T08:00:00Z",
            lastModifiedDateTime: "2024-01-10T08:00:00Z"
          }
        }
      }));

      global.fetch = async (url) => {
        return {
          ok: true,
          status: 200,
          headers: { get: () => null },
          json: async () => ({
            value: [
              {
                id: "page-1",
                title: "Moved Page",
                lastModifiedDateTime: "2024-01-20T08:00:00Z",
                createdDateTime: "2024-01-10T08:00:00Z"
              }
            ]
          })
        };
      };

      const summary = await diffGraphNotebook({ notebook: "A", root: tmpDir, full: true });
      assert.strictEqual(summary.totals.moved, 1);
      assert.strictEqual(summary.moved[0].previousSectionPath, "OldSec");
      assert.strictEqual(summary.moved[0].sectionPath, "NewSec");
    } finally {
      await fs.rm(tmpDir, { recursive: true, force: true });
    }
  });
});
