import { describe, it } from "node:test";
import assert from "node:assert";
import {
  parseArgs,
  sanitizeSegment,
  base64UrlEncode,
  truncateUtf8,
  normalizeTimestamp,
  pageModifiedAt,
  flattenStructureSections,
  shouldRetryGraphStatus,
  retryDelayMs,
  graphApiRoot,
  graphOneNoteRoot,
  graphPagesForSection,
  graphPagesForSectionRecent,
} from "../src/onenote-interactor.js";

describe("parseArgs", () => {
  it("parses simple command", () => {
    const { command, options } = parseArgs(["graph-diff"]);
    assert.strictEqual(command, "graph-diff");
    assert.deepStrictEqual(options, {});
  });

  it("parses flags with values", () => {
    const { command, options } = parseArgs(["graph-diff", "--notebook", "A", "--root", "/tmp/export"]);
    assert.strictEqual(command, "graph-diff");
    assert.strictEqual(options.notebook, "A");
    assert.strictEqual(options.root, "/tmp/export");
  });

  it("parses boolean flags", () => {
    const { options } = parseArgs(["graph-diff", "--full"]);
    assert.strictEqual(options.full, true);
  });

  it("ignores non-flag tokens", () => {
    const { options } = parseArgs(["graph-diff", "something", "--notebook", "A", "else"]);
    assert.strictEqual(options.notebook, "A");
    assert.strictEqual(options.something, undefined);
  });
});

describe("sanitizeSegment", () => {
  it("leaves clean strings alone", () => {
    assert.strictEqual(sanitizeSegment("hello-world"), "hello-world");
  });

  it("replaces invalid characters with underscores", () => {
    const result = sanitizeSegment('foo<bar>:"/\\|?*');
    assert.ok(!result.includes('<'));
    assert.ok(!result.includes('>'));
    assert.ok(!result.includes(':'));
    assert.ok(!result.includes('"'));
    assert.ok(!result.includes('/'));
    assert.ok(!result.includes('|'));
    assert.ok(!result.includes('?'));
    assert.ok(!result.includes('*'));
    assert.ok(result.startsWith('foo_bar'));
  });

  it("trims whitespace", () => {
    assert.strictEqual(sanitizeSegment("  hello  "), "hello");
  });

  it("returns 'untitled' for empty string", () => {
    assert.strictEqual(sanitizeSegment(""), "untitled");
  });

  it("replaces control characters", () => {
    assert.strictEqual(sanitizeSegment("test\x00\x1f"), "test__");
  });
});

describe("base64UrlEncode", () => {
  it("encodes simple string", () => {
    assert.strictEqual(base64UrlEncode("hello"), "aGVsbG8");
  });

  it("removes padding", () => {
    assert.ok(!base64UrlEncode("any").includes("="));
  });

  it("replaces + with - and / with _", () => {
    const encoded = base64UrlEncode(">\x7f");
    assert.ok(!encoded.includes("+"));
    assert.ok(!encoded.includes("/"));
  });
});

describe("truncateUtf8", () => {
  it("returns short string unchanged", () => {
    assert.strictEqual(truncateUtf8("hello", 100), "hello");
  });

  it("truncates to byte limit", () => {
    const result = truncateUtf8("hello world", 8);
    assert.strictEqual(result, "hello wo");
  });

  it("handles multi-byte UTF-8 correctly", () => {
    const result = truncateUtf8("日本語テキスト", 9);
    assert.strictEqual(result, "日本語");
  });

  it("returns 'untitled' for empty result", () => {
    assert.strictEqual(truncateUtf8("", 10), "untitled");
  });
});

describe("normalizeTimestamp", () => {
  it("returns ISO string for valid date", () => {
    const ts = normalizeTimestamp("2024-01-15T10:30:00Z");
    assert.strictEqual(ts, "2024-01-15T10:30:00.000Z");
  });

  it("returns null for undefined", () => {
    assert.strictEqual(normalizeTimestamp(undefined), null);
  });

  it("returns null for null", () => {
    assert.strictEqual(normalizeTimestamp(null), null);
  });
});

describe("pageModifiedAt", () => {
  it("prefers lastModifiedDateTime", () => {
    const page = {
      lastModifiedDateTime: "2024-01-15T10:00:00Z",
      lastModifiedTime: "2024-01-14T09:00:00Z",
      createdDateTime: "2024-01-10T08:00:00Z",
    };
    assert.strictEqual(pageModifiedAt(page), "2024-01-15T10:00:00.000Z");
  });

  it("falls back to lastModifiedTime", () => {
    const page = {
      lastModifiedTime: "2024-01-14T09:00:00Z",
      createdDateTime: "2024-01-10T08:00:00Z",
    };
    assert.strictEqual(pageModifiedAt(page), "2024-01-14T09:00:00.000Z");
  });

  it("falls back to createdDateTime", () => {
    const page = { createdDateTime: "2024-01-10T08:00:00Z" };
    assert.strictEqual(pageModifiedAt(page), "2024-01-10T08:00:00.000Z");
  });
});

describe("flattenStructureSections", () => {
  it("flattens top-level sections", () => {
    const structure = {
      sections: [{ name: "A", id: "1" }, { name: "B", id: "2" }],
      sectionGroups: [],
    };
    const flat = flattenStructureSections(structure);
    assert.strictEqual(flat.length, 2);
    assert.strictEqual(flat[0].path, "A");
    assert.strictEqual(flat[1].path, "B");
  });

  it("flattens nested section groups", () => {
    const structure = {
      sections: [{ name: "Top", id: "1" }],
      sectionGroups: [
        {
          name: "Group1",
          sections: [{ name: "S1", id: "2" }],
          sectionGroups: [
            {
              name: "Group2",
              sections: [{ name: "S2", id: "3" }],
              sectionGroups: [],
            },
          ],
        },
      ],
    };
    const flat = flattenStructureSections(structure);
    assert.strictEqual(flat.length, 3);
    assert.strictEqual(flat[0].path, "Top");
    assert.strictEqual(flat[1].path, "Group1/S1");
    assert.strictEqual(flat[2].path, "Group1/Group2/S2");
  });

  it("handles empty structure", () => {
    const flat = flattenStructureSections({});
    assert.strictEqual(flat.length, 0);
  });
});

describe("shouldRetryGraphStatus", () => {
  it("retries 429", () => assert.strictEqual(shouldRetryGraphStatus(429), true));
  it("retries 408", () => assert.strictEqual(shouldRetryGraphStatus(408), true));
  it("retries 423", () => assert.strictEqual(shouldRetryGraphStatus(423), true));
  it("retries 500", () => assert.strictEqual(shouldRetryGraphStatus(500), true));
  it("retries 502", () => assert.strictEqual(shouldRetryGraphStatus(502), true));
  it("retries 503", () => assert.strictEqual(shouldRetryGraphStatus(503), true));
  it("does not retry 400", () => assert.strictEqual(shouldRetryGraphStatus(400), false));
  it("does not retry 404", () => assert.strictEqual(shouldRetryGraphStatus(404), false));
  it("does not retry 200", () => assert.strictEqual(shouldRetryGraphStatus(200), false));
});

describe("graphApiRoot", () => {
  it("returns default root", () => {
    assert.strictEqual(graphApiRoot(), "https://graph.microsoft.com/v1.0/me");
  });

  it("returns user-specific root", () => {
    assert.strictEqual(graphApiRoot({ "user-id": "user123" }), "https://graph.microsoft.com/v1.0/users/user123");
  });
});

describe("graphOneNoteRoot", () => {
  it("returns default me endpoint", () => {
    assert.strictEqual(graphOneNoteRoot(), "https://graph.microsoft.com/v1.0/me/onenote");
  });

  it("returns user-specific endpoint", () => {
    assert.strictEqual(
      graphOneNoteRoot({ "user-id": "user123" }),
      "https://graph.microsoft.com/v1.0/users/user123/onenote"
    );
  });
});

describe("graphPagesForSectionRecent", () => {
  it("builds correct URL with $top=1", async () => {
    const originalFetch = global.fetch;
    let capturedUrl = null;
    global.fetch = async (url) => {
      capturedUrl = url;
      return {
        ok: true,
        status: 200,
        headers: { get: () => null },
        json: async () => ({ value: [] }),
      };
    };

    try {
      await graphPagesForSectionRecent("fake-token", "section-123");
      assert.ok(capturedUrl.includes("/sections/section-123/pages"));
      assert.ok(capturedUrl.includes("$top=1"));
    } finally {
      global.fetch = originalFetch;
    }
  });
});

describe("graphPagesForSection", () => {
  it("builds correct URL with $top=100", async () => {
    const originalFetch = global.fetch;
    let capturedUrl = null;
    global.fetch = async (url) => {
      capturedUrl = url;
      return {
        ok: true,
        status: 200,
        headers: { get: () => null },
        json: async () => ({ value: [] }),
      };
    };

    try {
      await graphPagesForSection("fake-token", "section-456");
      assert.ok(capturedUrl.includes("/sections/section-456/pages"));
      assert.ok(capturedUrl.includes("$top=100"));
    } finally {
      global.fetch = originalFetch;
    }
  });
});
