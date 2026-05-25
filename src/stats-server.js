import http from "node:http";
import os from "node:os";
import path from "node:path";
import { promises as fs } from "node:fs";
import { JSDOM } from "jsdom";

const ATTACHMENT_SCAN_TTL_MS = 30_000;
let attachmentScanCache = null;

function parseArgs(argv) {
  const options = {};
  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i];
    if (!token.startsWith("--")) {
      continue;
    }
    const key = token.slice(2);
    const next = argv[i + 1];
    if (!next || next.startsWith("--")) {
      options[key] = true;
      continue;
    }
    options[key] = next;
    i += 1;
  }
  return options;
}

async function pathExists(targetPath) {
  try {
    await fs.access(targetPath);
    return true;
  } catch {
    return false;
  }
}

async function readJsonIfExists(targetPath) {
  if (!(await pathExists(targetPath))) {
    return null;
  }
  return JSON.parse(await fs.readFile(targetPath, "utf8"));
}

async function tailFile(targetPath, maxLines = 250) {
  if (!(await pathExists(targetPath))) {
    return [];
  }
  const content = await fs.readFile(targetPath, "utf8");
  return content.split("\n").filter(Boolean).slice(-maxLines);
}

async function collectFiles(rootDir, extension) {
  const results = [];

  async function walk(currentDir) {
    const entries = await fs.readdir(currentDir, { withFileTypes: true });
    for (const entry of entries) {
      const fullPath = path.join(currentDir, entry.name);
      if (entry.isDirectory()) {
        await walk(fullPath);
      } else if (entry.isFile() && fullPath.endsWith(extension)) {
        results.push(fullPath);
      }
    }
  }

  if (await pathExists(rootDir)) {
    await walk(rootDir);
  }

  results.sort((a, b) => a.localeCompare(b));
  return results;
}

async function countFiles(rootDir) {
  let count = 0;
  let bytes = 0;

  async function walk(currentDir) {
    const entries = await fs.readdir(currentDir, { withFileTypes: true });
    for (const entry of entries) {
      const fullPath = path.join(currentDir, entry.name);
      if (entry.isDirectory()) {
        await walk(fullPath);
      } else if (entry.isFile()) {
        const stat = await fs.stat(fullPath);
        count += 1;
        bytes += stat.size;
      }
    }
  }

  if (await pathExists(rootDir)) {
    await walk(rootDir);
  }

  return { count, bytes };
}

function isGraphOneNoteResourceUrl(value) {
  return typeof value === "string" &&
    value.startsWith("https://graph.microsoft.com/") &&
    value.includes("/onenote/resources/");
}

function normalizeOneNoteHtml(html) {
  return html.replace(/<(iframe|object|audio|video)\b([^>]*)\/>/gi, "<$1$2></$1>");
}

function isLikelyImageReference(tag, value) {
  return tag === "img" || /\.(apng|avif|gif|jpe?g|png|svg|webp)(?:[?#].*)?$/i.test(value);
}

function decodeHtmlAttribute(value) {
  return value
    .replaceAll("&amp;", "&")
    .replaceAll("&quot;", "\"")
    .replaceAll("&#39;", "'")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">");
}

function extractResourceReferences(html) {
  const references = [];
  const dom = new JSDOM(normalizeOneNoteHtml(html));
  const candidates = [
    ["img", "src"],
    ["img", "data-fullres-src"],
    ["object", "data"],
    ["embed", "src"],
    ["audio", "src"],
    ["video", "src"],
    ["source", "src"],
    ["a", "href"]
  ];

  for (const [selector, attr] of candidates) {
    for (const node of dom.window.document.querySelectorAll(`${selector}[${attr}]`)) {
      const value = node.getAttribute(attr);
      if (value) {
        references.push({
          tag: selector,
          attr,
          value: decodeHtmlAttribute(value)
        });
      }
    }
  }

  return references;
}

async function scanAttachmentMetrics(rootDir) {
  const pagesDir = path.join(rootDir, "pages");
  const htmlFiles = await collectFiles(pagesDir, ".html");
  const uniqueRemote = new Set();
  const uniqueLocal = new Set();
  const pagesWithResources = new Set();
  const pagesWithRemoteResources = new Set();
  const missingLocalFiles = [];
  let totalReferences = 0;
  let remoteReferences = 0;
  let localizedReferences = 0;
  let imageReferences = 0;
  let fileAttachmentReferences = 0;

  for (const htmlFile of htmlFiles) {
    const html = await fs.readFile(htmlFile, "utf8");
    const references = extractResourceReferences(html);
    for (const ref of references) {
      const value = ref.value;
      const isRemote = isGraphOneNoteResourceUrl(value);
      const isLocalAsset = value.includes(".assets/");
      if (!isRemote && !isLocalAsset) {
        continue;
      }

      totalReferences += 1;
      pagesWithResources.add(htmlFile);
      if (isLikelyImageReference(ref.tag, value)) {
        imageReferences += 1;
      } else {
        fileAttachmentReferences += 1;
      }

      if (isRemote) {
        remoteReferences += 1;
        uniqueRemote.add(value);
        pagesWithRemoteResources.add(htmlFile);
        continue;
      }

      localizedReferences += 1;
      const localPath = path.resolve(path.dirname(htmlFile), value);
      uniqueLocal.add(localPath);
      if (!(await pathExists(localPath))) {
        missingLocalFiles.push(path.relative(rootDir, localPath));
      }
    }
  }

  let localizedFiles = 0;
  let localizedBytes = 0;
  const assetDirs = await collectAssetDirs(pagesDir);
  for (const assetDir of assetDirs) {
    const stats = await countFiles(assetDir);
    localizedFiles += stats.count;
    localizedBytes += stats.bytes;
  }
  const totalUniqueResources = uniqueLocal.size + uniqueRemote.size + missingLocalFiles.length;
  const localizedPercent = totalUniqueResources > 0
    ? Number(((uniqueLocal.size / totalUniqueResources) * 100).toFixed(2))
    : null;

  return {
    scannedAt: new Date().toISOString(),
    htmlFilesScanned: htmlFiles.length,
    pagesWithResources: pagesWithResources.size,
    pagesWithRemoteResources: pagesWithRemoteResources.size,
    totalReferences,
    imageReferences,
    fileAttachmentReferences,
    remoteReferences,
    uniqueRemoteResources: uniqueRemote.size,
    localizedReferences,
    uniqueLocalizedResources: uniqueLocal.size,
    localizedFiles,
    localizedBytes,
    totalUniqueResources,
    localizedPercent,
    orphanLocalFiles: Math.max(0, localizedFiles - uniqueLocal.size),
    missingLocalFiles: missingLocalFiles.length,
    missingLocalFileSamples: missingLocalFiles.slice(0, 20)
  };
}

async function collectAssetDirs(rootDir) {
  const results = [];

  async function walk(currentDir) {
    const entries = await fs.readdir(currentDir, { withFileTypes: true });
    for (const entry of entries) {
      const fullPath = path.join(currentDir, entry.name);
      if (!entry.isDirectory()) {
        continue;
      }
      if (entry.name.endsWith(".assets")) {
        results.push(fullPath);
      }
      await walk(fullPath);
    }
  }

  if (await pathExists(rootDir)) {
    await walk(rootDir);
  }

  return results;
}

async function getAttachmentMetrics(rootDir) {
  const now = Date.now();
  if (
    attachmentScanCache &&
    attachmentScanCache.rootDir === rootDir &&
    now - attachmentScanCache.createdAt < ATTACHMENT_SCAN_TTL_MS
  ) {
    return attachmentScanCache.metrics;
  }

  const metrics = await scanAttachmentMetrics(rootDir);
  attachmentScanCache = {
    rootDir,
    createdAt: now,
    metrics
  };
  return metrics;
}

function processAlive(pid) {
  if (!pid) {
    return false;
  }
  try {
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
}

function defaultRoot() {
  return path.join(process.cwd(), "exports", "graph", "A");
}

function statusLabel(progressState, alive) {
  if (!progressState) {
    return "idle";
  }
  if (progressState.status === "completed") {
    return "completed";
  }
  if (progressState.status === "failed") {
    return "failed";
  }
  if (progressState.status === "running" && alive) {
    return "running";
  }
  if (progressState.status === "running" && !alive) {
    return "stopped";
  }
  return progressState.status || "unknown";
}

function htmlPage() {
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>OneNote Interactor Progress</title>
  <style>
    :root {
      --bg: #0e1116;
      --panel: #171b22;
      --muted: #94a3b8;
      --text: #e5edf7;
      --accent: #61dafb;
      --good: #34d399;
      --warn: #fbbf24;
      --bad: #f87171;
      --border: #243041;
    }
    body {
      margin: 0;
      background: linear-gradient(180deg, #0b0f14 0%, #101723 100%);
      color: var(--text);
      font: 14px/1.45 ui-monospace, SFMono-Regular, Menlo, monospace;
    }
    .wrap {
      max-width: 1200px;
      margin: 0 auto;
      padding: 24px;
    }
    h1 { margin: 0 0 16px; font-size: 24px; }
    .grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 12px;
      margin-bottom: 16px;
    }
    .card, .logs {
      background: rgba(23, 27, 34, 0.92);
      border: 1px solid var(--border);
      border-radius: 14px;
      padding: 14px 16px;
      box-shadow: 0 12px 40px rgba(0,0,0,0.2);
    }
    .label { color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: .08em; }
    .value { font-size: 22px; margin-top: 6px; word-break: break-word; }
    .small { font-size: 12px; color: var(--muted); }
    .pill {
      display: inline-block;
      border-radius: 999px;
      padding: 4px 10px;
      font-size: 12px;
      font-weight: 700;
      border: 1px solid var(--border);
    }
    .running { color: var(--good); }
    .completed { color: var(--good); }
    .failed, .stopped { color: var(--bad); }
    .idle { color: var(--warn); }
    .row { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; margin: 10px 0 18px; }
    progress { width: 100%; height: 14px; }
    .logs pre {
      margin: 0;
      white-space: pre-wrap;
      word-break: break-word;
      max-height: 60vh;
      overflow: auto;
    }
    a { color: var(--accent); }
  </style>
</head>
<body>
  <div class="wrap">
    <h1>OneNote Interactor Progress</h1>
    <div class="row">
      <span id="status-pill" class="pill">loading</span>
      <span class="small" id="root"></span>
      <a href="/api/status" target="_blank">status json</a>
      <a href="/api/logs" target="_blank">logs json</a>
    </div>
    <div class="grid" id="cards"></div>
    <div class="card">
      <div class="label">Export Progress</div>
      <div class="value" id="progress-text">-</div>
      <progress id="progress-bar" value="0" max="100"></progress>
      <div class="small" id="attachment-progress-text" style="margin-top: 12px;">-</div>
      <progress id="attachment-progress-bar" value="0" max="100"></progress>
    </div>
    <div class="logs" style="margin-top: 16px;">
      <div class="label" style="margin-bottom: 10px;">Recent Logs</div>
      <pre id="logs">loading...</pre>
    </div>
  </div>
  <script>
    function card(label, value, small = "") {
      return '<div class="card"><div class="label">' + label + '</div><div class="value">' + value + '</div><div class="small">' + small + '</div></div>';
    }
    async function refresh() {
      const [statusRes, logsRes] = await Promise.all([
        fetch('/api/status'),
        fetch('/api/logs')
      ]);
      const status = await statusRes.json();
      const logs = await logsRes.json();
      document.getElementById('root').textContent = status.rootDir || '';
      const pill = document.getElementById('status-pill');
      pill.textContent = status.runtime.status;
      pill.className = 'pill ' + status.runtime.status;
      const cards = [];
      cards.push(card('Command', status.runtime.command || '-', status.runtime.phase || ''));
      cards.push(card('PID', status.runtime.pid || '-', status.runtime.alive ? 'alive' : 'not running'));
      cards.push(card('Current Section', status.runtime.currentSection || '-', ''));
      cards.push(card('Current Page', status.runtime.currentPage || '-', ''));
      cards.push(card('Accessible Pages', status.graph.accessiblePages ?? '-', 'from Graph count'));
      cards.push(card('Protected Sections', status.graph.protectedSections ?? '-', (status.graph.protectedSectionPaths || []).join(', ')));
      cards.push(card('HTML Pages', status.local.htmlFiles ?? '-', 'exported'));
      cards.push(card('Markdown Pages', status.local.markdownFiles ?? '-', 'generated'));
      cards.push(card('Asset Dirs', status.local.assetDirs ?? '-', 'localized resources'));
      cards.push(card('Resource Refs', status.attachments.totalReferences ?? '-', 'images/files referenced in HTML'));
      cards.push(card('Local Assets', status.attachments.localizedFiles ?? '-', (status.attachments.uniqueLocalizedResources ?? '-') + ' referenced unique, ' + ((status.attachments.localizedBytes || 0) / 1024 / 1024).toFixed(1) + ' MB'));
      cards.push(card('Asset Progress', (status.attachments.uniqueLocalizedResources ?? 0) + ' / ' + (status.attachments.totalUniqueResources ?? 0), 'orphan local files: ' + (status.attachments.orphanLocalFiles ?? 0)));
      cards.push(card('Missing Assets', (status.attachments.uniqueRemoteResources ?? 0) + (status.attachments.missingLocalFiles ?? 0), 'remote left + broken local links'));
      cards.push(card('Pages With Remote Assets', status.attachments.pagesWithRemoteResources ?? '-', 'need later asset retry'));
      cards.push(card('Section Summaries', status.local.sectionSummaries ?? '-', 'completed sections'));
      cards.push(card('Heartbeat', status.runtime.lastHeartbeatAt || '-', 'stale seconds: ' + (status.runtime.heartbeatAgeSeconds ?? '-')));
      cards.push(card('Last Log', status.runtime.lastLogAt || '-', ''));
      document.getElementById('cards').innerHTML = cards.join('');
      const htmlPercent = status.local.htmlPercent ?? 0;
      const mdPercent = status.local.markdownPercent ?? 0;
      const assetPercent = status.attachments.localizedPercent ?? 0;
      const assetLocalized = status.attachments.uniqueLocalizedResources ?? 0;
      const assetTotal = status.attachments.totalUniqueResources ?? 0;
      const assetMissing = (status.attachments.uniqueRemoteResources ?? 0) + (status.attachments.missingLocalFiles ?? 0);
      document.getElementById('progress-text').textContent = 'HTML ' + htmlPercent + '% | Markdown ' + mdPercent + '%';
      document.getElementById('progress-bar').value = htmlPercent;
      document.getElementById('attachment-progress-text').textContent = 'Files / Attachments ' + assetPercent + '% (' + assetLocalized + ' / ' + assetTotal + ', missing ' + assetMissing + ')';
      document.getElementById('attachment-progress-bar').value = assetPercent;
      document.getElementById('logs').textContent = (logs.lines || []).join('\\n');
    }
    refresh();
    setInterval(refresh, 5000);
  </script>
</body>
</html>`;
}

async function buildStatus(rootDir) {
  const runtimeDir = path.join(rootDir, ".runtime");
  const progressState = await readJsonIfExists(path.join(runtimeDir, "progress-state.json"));
  const countSummary = await readJsonIfExists(path.join(rootDir, "count-summary.json"));
  const alive = processAlive(progressState?.pid);
  const attachments = await getAttachmentMetrics(rootDir);
  const now = Date.now();
  const heartbeatAgeSeconds = progressState?.lastHeartbeatAt
    ? Math.max(0, Math.round((now - Date.parse(progressState.lastHeartbeatAt)) / 1000))
    : null;

  return {
    rootDir,
    runtime: {
      status: statusLabel(progressState, alive),
      command: progressState?.command || null,
      phase: progressState?.phase || null,
      pid: progressState?.pid || null,
      alive,
      currentSection: progressState?.currentSection || null,
      currentPage: progressState?.currentPage || null,
      startedAt: progressState?.startedAt || null,
      completedAt: progressState?.completedAt || null,
      lastHeartbeatAt: progressState?.lastHeartbeatAt || null,
      heartbeatAgeSeconds,
      lastLogAt: progressState?.lastLogAt || null,
      error: progressState?.error || null
    },
    graph: {
      accessiblePages: countSummary?.stats?.totalPages ?? null,
      accessibleSections: countSummary?.stats?.sectionsCounted ?? null,
      protectedSections: countSummary?.stats?.sectionsProtected ?? null,
      protectedSectionPaths: (countSummary?.stats?.protectedSections || []).map((item) => item.path)
    },
    local: {
      ...(progressState?.counters || {})
    },
    attachments
  };
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const rootDir = path.resolve(options.root || defaultRoot());
  const port = Number.parseInt(options.port || "9876", 10);

  const server = http.createServer(async (req, res) => {
    try {
      if (!req.url || req.url === "/") {
        res.writeHead(200, { "content-type": "text/html; charset=utf-8" });
        res.end(htmlPage());
        return;
      }

      if (req.url.startsWith("/api/status")) {
        const payload = await buildStatus(rootDir);
        res.writeHead(200, { "content-type": "application/json; charset=utf-8" });
        res.end(JSON.stringify(payload, null, 2));
        return;
      }

      if (req.url.startsWith("/api/logs")) {
        const lines = await tailFile(path.join(rootDir, ".runtime", "progress.log"));
        res.writeHead(200, { "content-type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ lines }, null, 2));
        return;
      }

      res.writeHead(404, { "content-type": "text/plain; charset=utf-8" });
      res.end("Not found");
    } catch (error) {
      res.writeHead(500, { "content-type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ error: error.message }, null, 2));
    }
  });

  server.listen(port, "127.0.0.1", () => {
    console.log(`Progress server running at http://127.0.0.1:${port}`);
    console.log(`Watching root: ${rootDir}`);
    console.log(`Home directory: ${os.homedir()}`);
  });
}

main().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
