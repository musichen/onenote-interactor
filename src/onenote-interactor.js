import { promises as fs } from "node:fs";
import path from "node:path";
import os from "node:os";
import { execFile as execFileCb } from "node:child_process";
import { promisify } from "node:util";
import { pathToFileURL } from "node:url";
import { JSDOM } from "jsdom";
import TurndownService from "turndown";
import { gfm } from "turndown-plugin-gfm";

const execFile = promisify(execFileCb);

const DEFAULT_BACKUP_ROOT = path.join(
  os.homedir(),
  "Library/Containers/com.microsoft.onenote.mac/Data/Library/Application Support/Microsoft User Data/OneNote/15.0/Backup"
);
const DEFAULT_CACHE_DIR = path.join(os.homedir(), ".config", "onenote-interactor");
const DEFAULT_CACHE_PATH = path.join(DEFAULT_CACHE_DIR, "msal-cache.json");
const DEFAULT_GRAPH_SCOPES = [
  "User.Read",
  "Notes.Read",
  "Notes.Read.All",
  "Files.Read",
  "Files.Read.All",
  "offline_access"
];
const DEFAULT_GRAPH_RETRY_COUNT = 8;
const DEFAULT_GRAPH_RESOURCE_RETRY_COUNT = 8;
const DEFAULT_GRAPH_BASE_DELAY_MS = 1500;
const DEFAULT_GRAPH_MAX_DELAY_MS = 120000;
const DEFAULT_GRAPH_MIN_INTERVAL_MS = 1000;
const DEFAULT_GRAPH_TIMEOUT_MS = 45000;
const DEFAULT_GRAPH_RESOURCE_TIMEOUT_MS = 30000;
const DEFAULT_PROGRESS_INTERVAL_MS = 15000;
let lastGraphRequestAt = 0;
let activeTracker = null;

function parseArgs(argv) {
  const [command, ...rest] = argv;
  const options = {};

  for (let i = 0; i < rest.length; i += 1) {
    const token = rest[i];
    if (!token.startsWith("--")) {
      continue;
    }

    const key = token.slice(2);
    const next = rest[i + 1];
    if (!next || next.startsWith("--")) {
      options[key] = true;
      continue;
    }

    options[key] = next;
    i += 1;
  }

  return { command, options };
}

function sanitizeSegment(value) {
  return value.replace(/[<>:"/\\|?*\u0000-\u001f]/g, "_").trim() || "untitled";
}

function base64UrlEncode(value) {
  return Buffer.from(value, "utf8")
    .toString("base64")
    .replaceAll("+", "-")
    .replaceAll("/", "_")
    .replace(/=+$/g, "");
}

async function loadDotEnvFile(filePath) {
  if (!(await pathExists(filePath))) {
    return;
  }
  const content = await fs.readFile(filePath, "utf8");
  for (const line of content.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) {
      continue;
    }
    const equalsIndex = trimmed.indexOf("=");
    if (equalsIndex === -1) {
      continue;
    }
    const key = trimmed.slice(0, equalsIndex).trim();
    let value = trimmed.slice(equalsIndex + 1).trim();
    if (
      (value.startsWith("\"") && value.endsWith("\"")) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    if (key && process.env[key] === undefined) {
      process.env[key] = value;
    }
  }
}

async function loadLocalEnv() {
  await loadDotEnvFile(path.join(process.cwd(), ".env"));
  await loadDotEnvFile(path.join(process.cwd(), ".env.local"));
}

function truncateUtf8(value, maxBytes) {
  let result = "";
  let bytes = 0;
  for (const char of value) {
    const charBytes = Buffer.byteLength(char);
    if (bytes + charBytes > maxBytes) {
      break;
    }
    result += char;
    bytes += charBytes;
  }
  return result.trim() || "untitled";
}

async function ensureDir(dirPath) {
  await fs.mkdir(dirPath, { recursive: true });
}

async function pathExists(targetPath) {
  try {
    await fs.access(targetPath);
    return true;
  } catch {
    return false;
  }
}

async function statSafe(targetPath) {
  try {
    return await fs.stat(targetPath);
  } catch {
    return null;
  }
}

async function scanTree(rootPath) {
  const stats = await fs.stat(rootPath);
  const node = {
    name: path.basename(rootPath),
    path: rootPath,
    type: stats.isDirectory() ? "directory" : "file",
    size: stats.size,
    mtime: stats.mtime.toISOString()
  };

  if (!stats.isDirectory()) {
    return node;
  }

  const entries = await fs.readdir(rootPath, { withFileTypes: true });
  entries.sort((a, b) => a.name.localeCompare(b.name));

  node.children = [];
  for (const entry of entries) {
    if (entry.name === ".DS_Store") {
      continue;
    }
    node.children.push(await scanTree(path.join(rootPath, entry.name)));
  }

  return node;
}

async function writeJson(targetPath, data) {
  await ensureDir(path.dirname(targetPath));
  await fs.writeFile(targetPath, JSON.stringify(data, null, 2));
}

async function readJson(targetPath) {
  return JSON.parse(await fs.readFile(targetPath, "utf8"));
}

async function readJsonIfExists(targetPath) {
  if (!(await pathExists(targetPath))) {
    return null;
  }
  return readJson(targetPath);
}

function nowIso() {
  return new Date().toISOString();
}

async function buildLocalIndex(options) {
  const backupRoot = options["backup-root"] || DEFAULT_BACKUP_ROOT;
  const notebookName = options.notebook || "A";
  const notebookPath = path.join(backupRoot, notebookName);

  if (!(await pathExists(notebookPath))) {
    throw new Error(`Notebook backup path does not exist: ${notebookPath}`);
  }

  const tree = await scanTree(notebookPath);
  const outputPath =
    options.out || path.join(process.cwd(), "exports", `local-index-${sanitizeSegment(notebookName)}.json`);

  await writeJson(outputPath, {
    generatedAt: new Date().toISOString(),
    backupRoot,
    notebookName,
    notebookPath,
    tree
  });

  console.log(`Wrote local index for notebook "${notebookName}" to ${outputPath}`);
  console.log(`Notebook path: ${notebookPath}`);
}

async function previewLocalFile(options) {
  const inputPath = options.input;
  if (!inputPath) {
    throw new Error("--input is required for local-preview");
  }

  const outputPath =
    options.out ||
    path.join(process.cwd(), "exports", `${sanitizeSegment(path.basename(inputPath))}.strings.txt`);

  const { stdout } = await execFile("strings", [inputPath], {
    maxBuffer: 50 * 1024 * 1024
  });

  await ensureDir(path.dirname(outputPath));
  await fs.writeFile(outputPath, stdout);

  console.log(`Wrote strings preview to ${outputPath}`);
}

function graphConfig() {
  const clientId = process.env.ONENOTE_CLIENT_ID;
  if (!clientId) {
    throw new Error("Missing ONENOTE_CLIENT_ID environment variable.");
  }

  return {
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${process.env.ONENOTE_TENANT_ID || "consumers"}`
    }
  };
}

async function getPublicClientApplication() {
  const { PublicClientApplication } = await import("@azure/msal-node");
  const pca = new PublicClientApplication(graphConfig());
  const tokenCache = pca.getTokenCache();
  if (await pathExists(DEFAULT_CACHE_PATH)) {
    try {
      await tokenCache.deserialize(await fs.readFile(DEFAULT_CACHE_PATH, "utf8"));
    } catch (error) {
      console.warn(`Could not load token cache: ${error.message}`);
    }
  }
  return pca;
}

async function saveTokenCache(pca) {
  const tokenCache = pca.getTokenCache();
  await ensureDir(DEFAULT_CACHE_DIR);
  await fs.writeFile(DEFAULT_CACHE_PATH, await tokenCache.serialize());
}

async function getAccessToken(scopes = DEFAULT_GRAPH_SCOPES) {
  const pca = await getPublicClientApplication();
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    try {
      const silent = await pca.acquireTokenSilent({
        account: accounts[0],
        scopes
      });
      if (silent?.accessToken) {
        await saveTokenCache(pca);
        return silent.accessToken;
      }
    } catch (error) {
      console.warn(`Silent token acquisition failed, falling back to device code: ${error.message}`);
    }
  }

  const result = await pca.acquireTokenByDeviceCode({
    scopes,
    deviceCodeCallback: (response) => {
      console.log(response.message);
    }
  });

  if (!result?.accessToken) {
    throw new Error("Failed to acquire Microsoft Graph access token.");
  }

  await saveTokenCache(pca);
  return result.accessToken;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function shouldRetryGraphStatus(status) {
  return status === 429 || status === 408 || status === 423 || status >= 500;
}

function retryDelayMs(response, attempt) {
  const retryAfter = response.headers.get("retry-after");
  if (retryAfter) {
    const seconds = Number.parseInt(retryAfter, 10);
    if (Number.isFinite(seconds) && seconds >= 0) {
      return Math.min(seconds * 1000, DEFAULT_GRAPH_MAX_DELAY_MS);
    }
  }
  return Math.min(DEFAULT_GRAPH_BASE_DELAY_MS * 2 ** attempt, DEFAULT_GRAPH_MAX_DELAY_MS);
}

async function throttleGraphRequest() {
  const now = Date.now();
  const elapsed = now - lastGraphRequestAt;
  if (elapsed < DEFAULT_GRAPH_MIN_INTERVAL_MS) {
    await sleep(DEFAULT_GRAPH_MIN_INTERVAL_MS - elapsed);
  }
  lastGraphRequestAt = Date.now();
}

async function graphFetch(
  initialToken,
  url,
  { accept, scopes = DEFAULT_GRAPH_SCOPES, retries = DEFAULT_GRAPH_RETRY_COUNT, timeoutMs = DEFAULT_GRAPH_TIMEOUT_MS } = {}
) {
  let token = initialToken;
  for (let attempt = 0; attempt <= retries; attempt += 1) {
    await throttleGraphRequest();
    let response;
    try {
      response = await fetch(url, {
        signal: AbortSignal.timeout(timeoutMs),
        headers: {
          Authorization: `Bearer ${token}`,
          ...(accept ? { Accept: accept } : {})
        }
      });
    } catch (error) {
      if (attempt < retries && (error.name === "TimeoutError" || error.name === "AbortError")) {
        const delayMs = Math.min(DEFAULT_GRAPH_BASE_DELAY_MS * 2 ** attempt, DEFAULT_GRAPH_MAX_DELAY_MS);
        console.warn(`Graph timeout for ${url}. Retrying in ${Math.round(delayMs / 1000)}s (${attempt + 1}/${retries})`);
        await sleep(delayMs);
        continue;
      }
      throw error;
    }

    if (response.ok) {
      return response;
    }

    if (response.status === 401 && attempt < retries) {
      console.warn(`Graph 401 for ${url}. Refreshing access token (${attempt + 1}/${retries})`);
      token = await getAccessToken(scopes);
      continue;
    }

    if (attempt < retries && shouldRetryGraphStatus(response.status)) {
      const delayMs = retryDelayMs(response, attempt);
      console.warn(
        `Graph ${response.status} for ${url}. Retrying in ${Math.round(delayMs / 1000)}s (${attempt + 1}/${retries})`
      );
      await sleep(delayMs);
      continue;
    }

    let bodyHint = "";
    try {
      const body = await response.clone().json();
      bodyHint = ` — ${body.error?.message || JSON.stringify(body)}`;
    } catch {
      try {
        bodyHint = ` — ${await response.clone().text()}`;
      } catch {
        // ignore
      }
    }
    throw new Error(`Graph request failed: ${response.status} ${response.statusText} for ${url}${bodyHint}`);
  }
}

async function graphFetchJson(token, url) {
  const response = await graphFetch(token, url, { accept: "application/json" });
  return response.json();
}

async function graphFetchAllJsonItems(token, url) {
  const items = [];
  let nextUrl = url;

  while (nextUrl) {
    const data = await graphFetchJson(token, nextUrl);
    items.push(...(data.value || []));
    nextUrl = data["@odata.nextLink"] || null;
  }

  return items;
}

async function graphFetchText(token, url) {
  const response = await graphFetch(token, url, { accept: "text/html" });
  return response.text();
}

async function graphFetchBuffer(token, url, options = {}) {
  const response = await graphFetch(token, url, {
    retries: options.retries ?? DEFAULT_GRAPH_RESOURCE_RETRY_COUNT,
    timeoutMs: options.timeoutMs ?? DEFAULT_GRAPH_RESOURCE_TIMEOUT_MS
  });
  return {
    buffer: Buffer.from(await response.arrayBuffer()),
    contentType: response.headers.get("content-type") || "application/octet-stream"
  };
}

function graphApiRoot(options = {}) {
  const userId = options["user-id"] || options.user;
  if (userId) {
    return `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userId)}`;
  }
  return "https://graph.microsoft.com/v1.0/me";
}

function graphOneNoteRoot(options = {}) {
  return `${graphApiRoot(options)}/onenote`;
}

async function graphLogin() {
  const token = await getAccessToken();
  const payload = JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString("utf8"));
  console.log(
    JSON.stringify(
      {
        message: "Authentication successful",
        cachePath: DEFAULT_CACHE_PATH,
        account: payload.preferred_username || payload.upn || payload.email || null,
        expiresAt: payload.exp ? new Date(payload.exp * 1000).toISOString() : null
      },
      null,
      2
    )
  );
}

function extensionFromContentType(contentType) {
  const normalized = contentType.split(";")[0].trim().toLowerCase();
  const map = new Map([
    ["image/png", ".png"],
    ["image/jpeg", ".jpg"],
    ["image/gif", ".gif"],
    ["image/webp", ".webp"],
    ["image/svg+xml", ".svg"],
    ["application/pdf", ".pdf"],
    ["text/plain", ".txt"],
    ["audio/mpeg", ".mp3"],
    ["audio/wav", ".wav"],
    ["video/mp4", ".mp4"],
    ["application/zip", ".zip"]
  ]);
  return map.get(normalized) || "";
}

function isGraphOneNoteResourceUrl(value) {
  return typeof value === "string" &&
    value.startsWith("https://graph.microsoft.com/") &&
    value.includes("/onenote/resources/");
}

function normalizeOneNoteHtml(html) {
  // OneNote sometimes emits XML-style self-closing tags for non-void HTML elements.
  // JSDOM/Turndown can then treat the rest of the document as iframe/object fallback HTML.
  return html
    .replace(/<(iframe|object|audio|video)\b([^>]*)\/>/gi, "<$1$2></$1>")
    .replace(/<\/iframe>/gi, "")
    .replace(/<iframe\b([^>]*)>/gi, "<iframe$1></iframe>");
}

function relativePath(fromFile, toFile) {
  return path.relative(path.dirname(fromFile), toFile) || path.basename(toFile);
}

function createTurndownService() {
  const service = new TurndownService({
    codeBlockStyle: "fenced",
    headingStyle: "atx",
    bulletListMarker: "-",
    emDelimiter: "_"
  });

  service.use(gfm);

  service.addRule("preserveCheckboxes", {
    filter: (node) => node.nodeName === "SPAN" && node.getAttribute?.("data-tag") === "to-do",
    replacement: (content) => `- [ ] ${content.trim()}\n`
  });

  service.addRule("preserveCompletedCheckboxes", {
    filter: (node) => node.nodeName === "P" && node.getAttribute?.("data-tag") === "to-do:completed",
    replacement: (content) => `- [x] ${content.trim()}\n`
  });

  service.addRule("attachmentObjects", {
    filter: (node) => node.nodeName === "OBJECT",
    replacement: (_content, node) => {
      const href = node.getAttribute?.("data") || "";
      const label = node.getAttribute?.("data-attachment") || href || "OneNote attachment";
      return href ? `\n[${label}](${href})\n` : `\n${label}\n`;
    }
  });

  service.addRule("embeddedFrames", {
    filter: (node) => node.nodeName === "IFRAME",
    replacement: (_content, node) => {
      const href = node.getAttribute?.("data-original-src") || node.getAttribute?.("src") || "";
      return href ? `\n[Embedded content](${href})\n` : "";
    }
  });

  service.addRule("plainTextTables", {
    filter: (node) => node.nodeName === "TABLE",
    replacement: (content) => {
      return `\n${content.replace(/\n{3,}/g, "\n\n").trim()}\n`;
    }
  });

  return service;
}

function normalizeInlineText(value) {
  return value.replace(/\s+/g, " ");
}

function markdownEscape(value) {
  return value
    .replaceAll("\\", "\\\\")
    .replaceAll("[", "\\[")
    .replaceAll("]", "\\]")
    .replaceAll("*", "\\*")
    .replaceAll("_", "\\_");
}

function nodeTextAsMarkdown(node) {
  if (node.nodeType === 3) {
    return normalizeInlineText(node.nodeValue || "");
  }

  if (node.nodeType !== 1) {
    return "";
  }

  const name = node.nodeName.toLowerCase();
  if (name === "br") {
    return "\n";
  }
  if (name === "a") {
    const href = node.getAttribute("href") || "";
    const label = markdownFromChildren(node).trim() || href;
    return href ? `[${label}](${href})` : label;
  }
  if (name === "img") {
    const src = node.getAttribute("src") || "";
    const alt = markdownEscape(node.getAttribute("alt") || "image");
    return src ? `![${alt}](${src})` : alt;
  }
  if (name === "object") {
    const href = node.getAttribute("data") || "";
    const label = markdownEscape(node.getAttribute("data-attachment") || href || "OneNote attachment");
    const childContent = markdownFromChildren(node).trim();
    const attachmentLink = href ? `[${label}](${href})` : label;
    return [attachmentLink, childContent].filter(Boolean).join("\n");
  }
  if (name === "iframe") {
    const href = node.getAttribute("data-original-src") || node.getAttribute("src") || "";
    return href ? `[Embedded content](${href})` : "";
  }
  if (name === "table") {
    return tableToMarkdownText(node);
  }

  return markdownFromChildren(node);
}

function markdownFromChildren(node) {
  return [...node.childNodes].map(nodeTextAsMarkdown).join("");
}

function blockToMarkdown(node) {
  const name = node.nodeName.toLowerCase();
  if (name === "table") {
    return tableToMarkdownText(node);
  }
  const text = markdownFromChildren(node)
    .split("\n")
    .map((line) => line.replace(/[ \t]+/g, " ").trim())
    .filter(Boolean)
    .join("\n");
  return text;
}

function tableToMarkdownText(node) {
  const rows = [...node.querySelectorAll("tr")].map((row) => {
    const cells = [...row.querySelectorAll("th,td")].map((cell) =>
      markdownFromChildren(cell).replace(/\s+/g, " ").trim()
    );
    return cells.filter(Boolean).join(" | ");
  });
  return rows.filter(Boolean).join("\n");
}

function topLeftFromStyle(node) {
  const style = node.getAttribute?.("style") || "";
  const top = Number.parseFloat(style.match(/top:\s*(-?\d+(?:\.\d+)?)px/i)?.[1] || "0");
  const left = Number.parseFloat(style.match(/left:\s*(-?\d+(?:\.\d+)?)px/i)?.[1] || "0");
  return { top, left };
}

function createMarkdownFromOneNoteDocument(document) {
  const body = document.body;
  if (!body) {
    return "";
  }

  const positionedBlocks = [...body.children]
    .filter((node) => node.nodeName.toLowerCase() !== "script")
    .map((node, index) => ({
      node,
      index,
      ...topLeftFromStyle(node)
    }))
    .sort((a, b) => a.top - b.top || a.left - b.left || a.index - b.index);

  const blocks = positionedBlocks
    .map(({ node }) => blockToMarkdown(node))
    .map((block) => block.trim())
    .filter(Boolean);

  return blocks.join("\n\n").replace(/\n{4,}/g, "\n\n\n").trim();
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

async function countDirectoriesBySuffix(rootDir, suffix) {
  let count = 0;

  async function walk(currentDir) {
    const entries = await fs.readdir(currentDir, { withFileTypes: true });
    for (const entry of entries) {
      const fullPath = path.join(currentDir, entry.name);
      if (!entry.isDirectory()) {
        continue;
      }
      if (entry.name.endsWith(suffix)) {
        count += 1;
      }
      await walk(fullPath);
    }
  }

  if (await pathExists(rootDir)) {
    await walk(rootDir);
  }

  return count;
}

async function collectDirectoriesBySuffix(rootDir, suffix) {
  const results = [];

  async function walk(currentDir) {
    const entries = await fs.readdir(currentDir, { withFileTypes: true });
    for (const entry of entries) {
      const fullPath = path.join(currentDir, entry.name);
      if (!entry.isDirectory()) {
        continue;
      }
      if (entry.name.endsWith(suffix)) {
        results.push(fullPath);
      }
      await walk(fullPath);
    }
  }

  if (await pathExists(rootDir)) {
    await walk(rootDir);
  }

  results.sort((a, b) => a.localeCompare(b));
  return results;
}

async function countFilesAndBytes(rootDir) {
  let count = 0;
  let bytes = 0;

  async function walk(currentDir) {
    const entries = await fs.readdir(currentDir, { withFileTypes: true });
    for (const entry of entries) {
      const fullPath = path.join(currentDir, entry.name);
      if (entry.isDirectory()) {
        await walk(fullPath);
      } else if (entry.isFile()) {
        const stats = await fs.stat(fullPath);
        count += 1;
        bytes += stats.size;
      }
    }
  }

  if (await pathExists(rootDir)) {
    await walk(rootDir);
  }

  return { count, bytes };
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

async function getExportResourceCounts(rootDir) {
  const pagesDir = path.join(rootDir, "pages");
  const htmlFiles = await collectFiles(pagesDir, ".html");
  const uniqueRemote = new Set();
  const uniqueLocal = new Set();
  const pagesWithResources = new Set();
  const pagesWithRemoteResources = new Set();
  const missingLocalFiles = [];
  let totalReferences = 0;
  let imageReferences = 0;
  let fileAttachmentReferences = 0;
  let remoteReferences = 0;
  let localizedReferences = 0;

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
  for (const assetDir of await collectDirectoriesBySuffix(pagesDir, ".assets")) {
    const stats = await countFilesAndBytes(assetDir);
    localizedFiles += stats.count;
    localizedBytes += stats.bytes;
  }
  const totalUniqueResources = uniqueLocal.size + uniqueRemote.size + missingLocalFiles.length;
  const localizedPercent = totalUniqueResources > 0
    ? Number(((uniqueLocal.size / totalUniqueResources) * 100).toFixed(2))
    : null;

  return {
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

async function getExportFilesystemCounts(rootDir) {
  const pagesDir = path.join(rootDir, "pages");
  return {
    htmlFiles: (await collectFiles(pagesDir, ".html")).length,
    markdownFiles: (await collectFiles(pagesDir, ".md")).length,
    jsonFiles: (await collectFiles(pagesDir, ".json")).length,
    assetDirs: await countDirectoriesBySuffix(pagesDir, ".assets"),
    sectionSummaries: (await collectFiles(pagesDir, "_section.json")).length
  };
}

async function createProgressTracker({ command, rootDir, notebookName = null }) {
  const runtimeDir = path.join(rootDir, ".runtime");
  const statePath = path.join(runtimeDir, "progress-state.json");
  const logPath = path.join(runtimeDir, "progress.log");
  await ensureDir(runtimeDir);

  let queue = Promise.resolve();
  const enqueue = (task) => {
    queue = queue.then(task).catch((error) => {
      process.stderr.write(`[progress-tracker] ${error.message}\n`);
    });
    return queue;
  };

  const state = {
    command,
    notebookName,
    rootDir,
    pid: process.pid,
    status: "running",
    phase: "starting",
    currentSection: null,
    currentPage: null,
    startedAt: nowIso(),
    lastHeartbeatAt: nowIso(),
    lastLogAt: null,
    completedAt: null,
    error: null,
    counters: {}
  };

  const originalConsole = {
    log: console.log.bind(console),
    warn: console.warn.bind(console),
    error: console.error.bind(console)
  };

  const persistState = async (withCounts = false) => {
    state.lastHeartbeatAt = nowIso();
    if (withCounts) {
      const countSummary = await readJsonIfExists(path.join(rootDir, "count-summary.json"));
      state.counters = {
        ...(await getExportFilesystemCounts(rootDir)),
        countedTotalPages: countSummary?.stats?.totalPages ?? null,
        countedSections: countSummary?.stats?.sectionsCounted ?? null,
        protectedSections: countSummary?.stats?.sectionsProtected ?? null
      };
    }
    await writeJson(statePath, state);
  };

  const persistSnapshot = async () => {
    const countSummary = await readJsonIfExists(path.join(rootDir, "count-summary.json"));
    state.lastHeartbeatAt = nowIso();
    state.counters = {
      ...(await getExportFilesystemCounts(rootDir)),
      countedTotalPages: countSummary?.stats?.totalPages ?? null,
      countedSections: countSummary?.stats?.sectionsCounted ?? null,
      protectedSections: countSummary?.stats?.sectionsProtected ?? null
    };
    if (state.counters.countedTotalPages) {
      state.counters.htmlPercent = Number(
        ((state.counters.htmlFiles / state.counters.countedTotalPages) * 100).toFixed(2)
      );
      state.counters.markdownPercent = Number(
        ((state.counters.markdownFiles / state.counters.countedTotalPages) * 100).toFixed(2)
      );
    } else {
      state.counters.htmlPercent = null;
      state.counters.markdownPercent = null;
    }
    await writeJson(statePath, state);
  };

  const appendLog = (level, args) => {
    const rendered = args
      .map((value) => (typeof value === "string" ? value : JSON.stringify(value)))
      .join(" ");
    const line = `[${nowIso()}] [${level}] ${rendered}\n`;
    state.lastLogAt = nowIso();
    enqueue(() => fs.appendFile(logPath, line));
  };

  console.log = (...args) => {
    originalConsole.log(...args);
    appendLog("info", args);
  };
  console.warn = (...args) => {
    originalConsole.warn(...args);
    appendLog("warn", args);
  };
  console.error = (...args) => {
    originalConsole.error(...args);
    appendLog("error", args);
  };

  await fs.writeFile(logPath, "");
  await persistSnapshot();

  const interval = setInterval(() => {
    enqueue(() => persistSnapshot());
  }, DEFAULT_PROGRESS_INTERVAL_MS);
  interval.unref();

  return {
    statePath,
    logPath,
    async setPhase(phase) {
      state.phase = phase;
      await enqueue(() => persistState());
    },
    async setCurrentSection(section) {
      state.currentSection = section;
      await enqueue(() => persistState());
    },
    async setCurrentPage(page) {
      state.currentPage = page;
      await enqueue(() => persistState());
    },
    async finalize(status, error = null) {
      clearInterval(interval);
      state.status = status;
      state.error = error ? String(error.message || error) : null;
      state.completedAt = nowIso();
      state.currentPage = null;
      await enqueue(() => persistSnapshot());
      console.log = originalConsole.log;
      console.warn = originalConsole.warn;
      console.error = originalConsole.error;
    }
  };
}

function createRunStats() {
  return {
    sectionsCompleted: 0,
    sectionsSkipped: 0,
    sectionsFailed: 0,
    sectionsProtected: 0,
    pagesExported: 0,
    pagesSkipped: 0,
    pagesFailed: 0,
    failedSections: [],
    protectedSections: [],
    failedPages: []
  };
}

function pageBasePath(sectionDir, page) {
  const rawTitle = sanitizeSegment(page.title || page.id);
  const id = String(page.id);
  const suffix = `-${id}`;
  const maxFilenameBytes = 220;
  const title = Buffer.byteLength(`${rawTitle}${suffix}`) <= maxFilenameBytes
    ? rawTitle
    : truncateUtf8(rawTitle, Math.max(40, maxFilenameBytes - Buffer.byteLength(suffix)));
  return path.join(sectionDir, `${title}-${page.id}`);
}

function pageFileSet(sectionDir, page) {
  const basePath = pageBasePath(sectionDir, page);
  return {
    basePath,
    htmlPath: `${basePath}.html`,
    jsonPath: `${basePath}.json`,
    markdownPath: `${basePath}.md`,
    assetDir: `${basePath}.assets`
  };
}

function normalizeTimestamp(value) {
  return value ? new Date(value).toISOString() : null;
}

function pageModifiedAt(page) {
  return normalizeTimestamp(page.lastModifiedDateTime || page.lastModifiedTime || page.createdDateTime);
}

async function loadLocalManifest(rootDir) {
  return (await readJsonIfExists(path.join(rootDir, "pages-manifest.json"))) || {
    generatedAt: null,
    rootDir,
    pages: {}
  };
}

async function writeLocalManifest(rootDir, manifest) {
  await writeJson(path.join(rootDir, "pages-manifest.json"), {
    ...manifest,
    generatedAt: nowIso(),
    rootDir
  });
}

async function updateManifestPage(rootDir, manifest, sectionPath, sectionDir, page) {
  const files = pageFileSet(sectionDir, page);
  const rel = (targetPath) => path.relative(rootDir, targetPath);
  manifest.pages[page.id] = {
    id: page.id,
    title: page.title || page.id,
    sectionPath,
    lastModifiedDateTime: pageModifiedAt(page),
    createdDateTime: normalizeTimestamp(page.createdDateTime),
    htmlPath: rel(files.htmlPath),
    jsonPath: rel(files.jsonPath),
    markdownPath: rel(files.markdownPath),
    assetDir: rel(files.assetDir),
    webUrl: page.links?.oneNoteWebUrl?.href || page.webUrl || null,
    exportedAt: nowIso()
  };
}

async function postprocessHtmlFile(token, htmlPath, turndownService, options = {}) {
  const html = normalizeOneNoteHtml(await fs.readFile(htmlPath, "utf8"));
  const dom = new JSDOM(html);
  const { document } = dom.window;
  const assetDir = `${htmlPath.slice(0, -".html".length)}.assets`;
  const resourceMap = new Map();
  const failedResources = [];
  let assetCounter = 0;

  async function localizeUrl(url) {
    if (!isGraphOneNoteResourceUrl(url)) {
      return url;
    }

    if (options.skipAssets) {
      return url;
    }

    if (resourceMap.has(url)) {
      return relativePath(htmlPath, resourceMap.get(url));
    }

    await ensureDir(assetDir);
    let response;
    try {
      response = await graphFetchBuffer(token, url);
    } catch (error) {
      failedResources.push({ url, error: error.message });
      console.warn(`  Resource skipped after retries: ${url} :: ${error.message}`);
      return url;
    }
    const { buffer, contentType } = response;
    const urlObject = new URL(url);
    const resourceId = urlObject.pathname.split("/").at(-2) || `resource-${assetCounter + 1}`;
    const assetFilename = `${sanitizeSegment(resourceId)}${extensionFromContentType(contentType)}`;
    const targetPath = path.join(assetDir, assetFilename);
    await fs.writeFile(targetPath, buffer);
    resourceMap.set(url, targetPath);
    assetCounter += 1;
    return relativePath(htmlPath, targetPath);
  }

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
    const nodes = [...document.querySelectorAll(`${selector}[${attr}]`)];
    for (const node of nodes) {
      const current = node.getAttribute(attr);
      if (!current) {
        continue;
      }
      if (isGraphOneNoteResourceUrl(current)) {
        node.setAttribute(attr, await localizeUrl(current));
      }
    }
  }

  const processedHtml = dom.serialize();
  await fs.writeFile(htmlPath, processedHtml);

  const markdown = createMarkdownFromOneNoteDocument(document) || turndownService.turndown(document.body.innerHTML);
  const markdownPath = `${htmlPath.slice(0, -".html".length)}.md`;
  await fs.writeFile(markdownPath, `${markdown.trim()}\n`);

  return {
    htmlPath,
    markdownPath,
    downloadedAssets: assetCounter,
    failedResources
  };
}

async function graphFindNotebook(token, notebookName, options = {}) {
  if (options["notebook-id"]) {
    const notebook = await graphFetchJson(
      token,
      `${graphOneNoteRoot(options)}/notebooks/${options["notebook-id"]}`
    );
    return notebook;
  }

  const data = await graphFetchJson(
    token,
    `${graphOneNoteRoot(options)}/notebooks?$top=200`
  );

  const match = (data.value || []).find((item) => item.displayName === notebookName);
  if (!match) {
    const names = (data.value || []).map((item) => item.displayName).join(", ");
    throw new Error(`Notebook "${notebookName}" not found. Available notebooks: ${names}`);
  }

  return match;
}

async function graphResolveLink(options) {
  const link = options.url || options.link || options._;
  if (!link || typeof link !== "string") {
    throw new Error("Provide a OneDrive/OneNote link with --url.");
  }

  const token = await getAccessToken();
  const shareId = `u!${base64UrlEncode(link)}`;
  const url = `https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem`;
  const driveItem = await graphFetchJson(token, url);
  const result = {
    inputUrl: link,
    shareId,
    driveItem: {
      id: driveItem.id,
      name: driveItem.name,
      webUrl: driveItem.webUrl,
      size: driveItem.size,
      package: driveItem.package || null,
      folder: driveItem.folder || null,
      file: driveItem.file || null,
      parentReference: driveItem.parentReference || null,
      remoteItem: driveItem.remoteItem || null,
      sharepointIds: driveItem.sharepointIds || null
    }
  };
  console.log(JSON.stringify(result, null, 2));
}

async function graphGet(options) {
  const url = options.url;
  if (!url || typeof url !== "string") {
    throw new Error("Provide a Graph URL with --url.");
  }
  const token = await getAccessToken();
  const data = await graphFetchJson(token, url);
  console.log(JSON.stringify(data, null, 2));
}

async function safeChildPath(parentDir, originalName, usedNames) {
  const parsed = path.parse(sanitizeSegment(originalName));
  let candidate = `${parsed.name}${parsed.ext}`;
  let counter = 2;
  while (usedNames.has(candidate.toLowerCase()) || (await pathExists(path.join(parentDir, candidate)))) {
    candidate = `${parsed.name}-${counter}${parsed.ext}`;
    counter += 1;
  }
  usedNames.add(candidate.toLowerCase());
  return path.join(parentDir, candidate);
}

async function exportDrivePackageItem(token, driveId, itemId, targetDir, stats, itemPath = []) {
  await ensureDir(targetDir);
  const childrenUrl =
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/children` +
    "?$select=id,name,size,folder,file,package,webUrl,lastModifiedDateTime,createdDateTime,parentReference";
  const children = await graphFetchAllJsonItems(token, childrenUrl);
  const usedNames = new Set();

  for (const child of children) {
    const childPath = [...itemPath, child.name];
    const localPath = await safeChildPath(targetDir, child.name, usedNames);
    stats.itemsSeen += 1;
    await activeTracker?.setCurrentSection(path.dirname(childPath.join("/")) || "/");
    await activeTracker?.setCurrentPage(childPath.join("/"));

    if (child.folder || child.package) {
      stats.foldersCreated += 1;
      await activeTracker?.setPhase("drive-folder");
      await ensureDir(localPath);
      await writeJson(path.join(localPath, ".driveItem.json"), child);
      await exportDrivePackageItem(token, driveId, child.id, localPath, stats, childPath);
      continue;
    }

    await activeTracker?.setPhase("drive-file");
    console.log(`Drive file: ${childPath.join("/")}`);
    try {
      const { buffer } = await graphFetchBuffer(
        token,
        `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${child.id}/content`,
        { retries: DEFAULT_GRAPH_RETRY_COUNT, timeoutMs: 120000 }
      );
      await fs.writeFile(localPath, buffer);
      await writeJson(`${localPath}.driveItem.json`, child);
      stats.filesDownloaded += 1;
      stats.bytesDownloaded += buffer.byteLength;
      stats.files.push({
        id: child.id,
        name: child.name,
        path: childPath.join("/"),
        localPath: path.relative(stats.rootDir, localPath),
        graphSize: child.size ?? null,
        downloadedBytes: buffer.byteLength
      });
    } catch (error) {
      stats.filesFailed += 1;
      stats.failedFiles.push({
        id: child.id,
        name: child.name,
        path: childPath.join("/"),
        error: error.message
      });
      console.error(`Failed Drive file ${childPath.join("/")}: ${error.message}`);
    }
  }
}

async function exportDrivePackage(options) {
  const driveId = options["drive-id"];
  const itemId = options["item-id"];
  if (!driveId || !itemId) {
    throw new Error("Provide --drive-id and --item-id for graph-drive-export.");
  }

  const outDir = options.out || path.join(process.cwd(), "exports", "drive", sanitizeSegment(itemId));
  const token = await getAccessToken();
  await ensureDir(outDir);
  await activeTracker?.setPhase("drive-export-start");

  const rootItem = await graphFetchJson(
    token,
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}?$select=id,name,size,folder,file,package,webUrl,lastModifiedDateTime,createdDateTime,parentReference`
  );
  await writeJson(path.join(outDir, "drive-package-root.json"), rootItem);

  const stats = {
    generatedAt: nowIso(),
    rootDir: outDir,
    driveId,
    itemId,
    packageName: rootItem.name,
    packageSize: rootItem.size ?? null,
    itemsSeen: 0,
    foldersCreated: 0,
    filesDownloaded: 0,
    filesFailed: 0,
    bytesDownloaded: 0,
    files: [],
    failedFiles: []
  };

  console.log(`Starting Drive package export: ${rootItem.name} (${driveId}/${itemId})`);
  await exportDrivePackageItem(token, driveId, itemId, outDir, stats, [rootItem.name]);
  await writeJson(path.join(outDir, "drive-package-summary.json"), stats);
  console.log(`Drive package export complete: ${stats.filesDownloaded} files, ${stats.filesFailed} failed.`);
}

async function graphSectionGroups(token, notebookId, options = {}) {
  const items = await graphFetchAllJsonItems(
    token,
    `${graphOneNoteRoot(options)}/notebooks/${notebookId}/sectionGroups?$top=200`
  );
  return items;
}

async function graphSectionsForNotebook(token, notebookId, options = {}) {
  const items = await graphFetchAllJsonItems(
    token,
    `${graphOneNoteRoot(options)}/notebooks/${notebookId}/sections?$top=200`
  );
  return items;
}

async function graphSectionsForGroup(token, sectionGroupId, options = {}) {
  const items = await graphFetchAllJsonItems(
    token,
    `${graphOneNoteRoot(options)}/sectionGroups/${sectionGroupId}/sections?$top=200`
  );
  return items;
}

async function graphChildGroups(token, sectionGroupId, options = {}) {
  const items = await graphFetchAllJsonItems(
    token,
    `${graphOneNoteRoot(options)}/sectionGroups/${sectionGroupId}/sectionGroups?$top=200`
  );
  return items;
}

async function graphPagesForSection(token, sectionId, options = {}) {
  const items = await graphFetchAllJsonItems(
    token,
    `${graphOneNoteRoot(options)}/sections/${sectionId}/pages?$top=100`
  );
  return items;
}

async function graphPagesForSectionRecent(token, sectionId, options = {}) {
  // Fetch just the most recently modified page (default order is lastModifiedDateTime desc)
  const items = await graphFetchAllJsonItems(
    token,
    `${graphOneNoteRoot(options)}/sections/${sectionId}/pages?$top=1`
  );
  return items;
}

async function buildGraphStructure(token, notebook, options = {}) {
  async function walkGroup(sectionGroup) {
    const sections = await graphSectionsForGroup(token, sectionGroup.id, options);
    const childGroups = await graphChildGroups(token, sectionGroup.id, options);

    return {
      id: sectionGroup.id,
      name: sectionGroup.displayName,
      type: "sectionGroup",
      sections: sections.map((section) => ({
        id: section.id,
        name: section.displayName,
        type: "section"
      })),
      sectionGroups: await Promise.all(childGroups.map(walkGroup))
    };
  }

  const [sections, sectionGroups] = await Promise.all([
    graphSectionsForNotebook(token, notebook.id, options),
    graphSectionGroups(token, notebook.id, options)
  ]);

  return {
    id: notebook.id,
    name: notebook.displayName,
    webUrl: notebook.links?.oneNoteWebUrl?.href || null,
    sections: sections.map((section) => ({
      id: section.id,
      name: section.displayName,
      type: "section"
    })),
    sectionGroups: await Promise.all(sectionGroups.map(walkGroup))
  };
}

async function listGraphNotebook(options) {
  const notebookName = options.notebook || "A";
  const token = await getAccessToken();
  const notebook = await graphFindNotebook(token, notebookName, options);
  const structure = await buildGraphStructure(token, notebook, options);

  console.log(JSON.stringify({ notebook, structure }, null, 2));
}

async function listAllGraphNotebooks(options) {
  const token = await getAccessToken();
  // Use lighter retry settings for this quick list call (2 retries, 15s timeout)
  const response = await graphFetch(token, `${graphOneNoteRoot(options)}/notebooks?$top=200`, {
    accept: "application/json",
    retries: 2,
    timeoutMs: 15000,
  });
  const data = await response.json();
  const notebooks = (data.value || []).map((n) => ({
    id: n.id,
    displayName: n.displayName,
    createdDateTime: n.createdDateTime,
    lastModifiedDateTime: n.lastModifiedDateTime,
  }));
  console.log(JSON.stringify({ notebooks }, null, 2));
}

function flattenStructureSections(structure) {
  const flattened = [];

  for (const section of structure.sections || []) {
    flattened.push({
      path: section.name,
      ...section
    });
  }

  function walkGroup(group, prefix = []) {
    const groupPath = [...prefix, group.name];
    for (const section of group.sections || []) {
      flattened.push({
        path: [...groupPath, section.name].join("/"),
        ...section
      });
    }
    for (const child of group.sectionGroups || []) {
      walkGroup(child, groupPath);
    }
  }

  for (const group of structure.sectionGroups || []) {
    walkGroup(group, []);
  }

  return flattened;
}

async function countGraphNotebookPages(options) {
  const notebookName = options.notebook || "A";
  const outPath = options.out || path.join(process.cwd(), "exports", "graph", sanitizeSegment(notebookName), "count-summary.json");
  const token = await getAccessToken();
  const notebook = await graphFindNotebook(token, notebookName, options);
  const structure = await buildGraphStructure(token, notebook, options);
  const sections = flattenStructureSections(structure);

  const stats = {
    notebook: notebook.displayName,
    sectionCount: sections.length,
    totalPages: 0,
    sectionsCounted: 0,
    sectionsProtected: 0,
    sectionsFailed: 0,
    protectedSections: [],
    failedSections: [],
    perSection: []
  };

  for (const section of sections) {
    try {
      const pages = await graphPagesForSection(token, section.id, options);
      stats.totalPages += pages.length;
      stats.sectionsCounted += 1;
      stats.perSection.push({
        path: section.path,
        id: section.id,
        pages: pages.length
      });
      console.log(`Counted ${pages.length} pages: ${section.path}`);
    } catch (error) {
      if (error.message.includes("403 Forbidden")) {
        stats.sectionsProtected += 1;
        stats.protectedSections.push({
          path: section.path,
          id: section.id,
          error: error.message
        });
        console.error(`Protected section while counting: ${section.path}`);
      } else {
        stats.sectionsFailed += 1;
        stats.failedSections.push({
          path: section.path,
          id: section.id,
          error: error.message
        });
        console.error(`Failed section while counting: ${section.path} :: ${error.message}`);
      }
    }
  }

  await writeJson(outPath, {
    generatedAt: new Date().toISOString(),
    stats
  });

  console.log(`Count summary written to ${outPath}`);
}

function sectionDirFromPath(rootDir, sectionPath) {
  const segments = sectionPath.split("/").map(sanitizeSegment);
  return path.join(rootDir, "pages", ...segments);
}

async function buildLocalPagesManifest(options) {
  const rootDir = options.root || path.join(process.cwd(), "exports", "graph", "A");
  const pagesRoot = path.join(rootDir, "pages");
  const sectionSummaries = await collectFiles(pagesRoot, "_section.json");
  const manifest = {
    generatedAt: nowIso(),
    rootDir,
    pages: {}
  };

  for (const summaryPath of sectionSummaries) {
    const sectionDir = path.dirname(summaryPath);
    const sectionPath = path.relative(pagesRoot, sectionDir);
    const summary = await readJson(summaryPath);
    for (const page of summary.pages || []) {
      const files = pageFileSet(sectionDir, page);
      if ((await pathExists(files.htmlPath)) && (await pathExists(files.jsonPath))) {
        await updateManifestPage(rootDir, manifest, sectionPath, sectionDir, page);
      }
    }
  }

  await writeLocalManifest(rootDir, manifest);
  console.log(`Wrote manifest with ${Object.keys(manifest.pages).length} exported pages to ${path.join(rootDir, "pages-manifest.json")}`);
  return manifest;
}

async function getNotebookAndStructureForRoot(token, notebookName, rootDir, options = {}) {
  const notebookPath = path.join(rootDir, "notebook.json");
  const structurePath = path.join(rootDir, "structure.json");
  if (options["refresh-structure"] !== true && (await pathExists(notebookPath)) && (await pathExists(structurePath))) {
    return {
      notebook: await readJson(notebookPath),
      structure: await readJson(structurePath)
    };
  }

  const notebook = await graphFindNotebook(token, notebookName, options);
  const structure = await buildGraphStructure(token, notebook, options);
  await ensureDir(rootDir);
  await writeJson(notebookPath, notebook);
  await writeJson(structurePath, structure);
  return { notebook, structure };
}

async function scanRemotePages(token, structure, options = {}) {
  const sections = flattenStructureSections(structure);
  const pages = {};
  const protectedSections = [];
  const failedSections = [];
  const isQuick = !!options.cutoffTimestamp;

  let scannedSections = 0;
  for (const section of sections) {
    await activeTracker?.setPhase("scan-section");
    await activeTracker?.setCurrentSection(section.path);
    try {
      let sectionPages;
      if (isQuick) {
        // Quick mode: check the single most recently modified page first.
        // Default ordering is by lastModifiedDateTime, so $top=1 gives us
        // the most recent. If it's older than the cutoff, skip the section.
        const recent = await graphPagesForSectionRecent(token, section.id, options);
        if (recent.length === 0) {
          console.log(`Skipped (empty): ${section.path}`);
          continue;
        }
        const mostRecentTime = pageModifiedAt(recent[0]);
        if (!mostRecentTime || new Date(mostRecentTime).getTime() <= options.cutoffTimestamp) {
          console.log(`Skipped (unchanged): ${section.path}`);
          continue;
        }
        // Section has recent changes — do a full scan of just this section
        sectionPages = await graphPagesForSection(token, section.id, options);
      } else {
        sectionPages = await graphPagesForSection(token, section.id, options);
      }
      for (const page of sectionPages) {
        pages[page.id] = {
          id: page.id,
          title: page.title || page.id,
          sectionPath: section.path,
          sectionId: section.id,
          lastModifiedDateTime: pageModifiedAt(page),
          createdDateTime: normalizeTimestamp(page.createdDateTime),
          page
        };
      }
      scannedSections += 1;
      console.log(`Scanned ${sectionPages.length} pages: ${section.path}${isQuick ? " (quick)" : ""}`);
    } catch (error) {
      if (error.message.includes("403 Forbidden")) {
        protectedSections.push({ path: section.path, id: section.id, error: error.message });
        console.error(`Protected section while scanning: ${section.path}`);
      } else {
        failedSections.push({ path: section.path, id: section.id, error: error.message });
        console.error(`Failed section while scanning: ${section.path} :: ${error.message}`);
      }
    }
  }

  console.log(`Scan complete: ${Object.keys(pages).length} total pages from ${scannedSections}/${sections.length} sections (${protectedSections.length} protected, ${failedSections.length} failed)`);
  return { pages, protectedSections, failedSections, isQuick };
}

async function writeDiffLogMarkdown(summary, rootDir) {
  const projectRoot = process.cwd();
  const logsDir = path.join(projectRoot, "logs");
  await ensureDir(logsDir);
  const ts = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
  const logPath = path.join(logsDir, `diff-${sanitizeSegment(summary.notebook)}-${ts}.md`);

  const fsCounts = await getExportFilesystemCounts(rootDir);
  const resCounts = await getExportResourceCounts(rootDir);

  const t = summary.totals;
  const lines = [];
  lines.push(`# OneNote Diff Report — ${summary.notebook}`);
  lines.push("");
  lines.push(`- **Notebook:** \`${summary.notebook}\``);
  lines.push(`- **Export root:** \`${rootDir}\``);
  lines.push(`- **Generated:** ${summary.generatedAt}`);
  lines.push("");

  lines.push("## Summary");
  lines.push("");
  lines.push("| Metric | Value |");
  lines.push("|--------|-------|");
  lines.push(`| Remote pages (Graph) | ${t.remotePages} |`);
  lines.push(`| Local manifest pages | ${t.localManifestPages} |`);
  lines.push(`| Added | ${t.added} |`);
  lines.push(`| Updated | ${t.updated} |`);
  lines.push(`| Deleted | ${t.deleted} |`);
  lines.push(`| Missing local files | ${t.missingLocal} |`);
  lines.push(`| Unchanged | ${t.unchanged} |`);
  lines.push(`| Protected sections | ${t.protectedSections} |`);
  lines.push(`| Failed sections | ${t.failedSections} |`);
  lines.push("");

  lines.push("## Filesystem Stats");
  lines.push("");
  lines.push("| Metric | Value |");
  lines.push("|--------|-------|");
  lines.push(`| HTML files | ${fsCounts.htmlFiles} |`);
  lines.push(`| Markdown files | ${fsCounts.markdownFiles} |`);
  lines.push(`| JSON files | ${fsCounts.jsonFiles} |`);
  lines.push(`| Asset directories | ${fsCounts.assetDirs} |`);
  lines.push(`| Section summaries | ${fsCounts.sectionSummaries} |`);
  lines.push("");

  lines.push("## Resource Stats");
  lines.push("");
  lines.push("| Metric | Value |");
  lines.push("|--------|-------|");
  lines.push(`| Total resource references | ${resCounts.totalReferences} |`);
  lines.push(`| Image references | ${resCounts.imageReferences} |`);
  lines.push(`| File/attachment references | ${resCounts.fileAttachmentReferences} |`);
  lines.push(`| Remote Graph references | ${resCounts.remoteReferences} |`);
  lines.push(`| Localized references | ${resCounts.localizedReferences} |`);
  lines.push(`| Unique local asset files | ${resCounts.uniqueLocalizedResources} |`);
  lines.push(`| Unique remote resources | ${resCounts.uniqueRemoteResources} |`);
  lines.push(`| Pages with resources | ${resCounts.pagesWithResources} |`);
  lines.push(`| Pages with remote resources | ${resCounts.pagesWithRemoteResources} |`);
  lines.push(`| Missing local asset files | ${resCounts.missingLocalFiles} |`);
  lines.push("");

  function pageTable(title, items, extraCols = []) {
    if (!items || items.length === 0) return;
    lines.push(`## ${title} (${items.length})`);
    lines.push("");
    const headers = ["Title", "Section Path", ...extraCols.map(c => c.header)];
    lines.push("| " + headers.join(" | ") + " |");
    lines.push("| " + headers.map(() => "---").join(" | ") + " |");
    for (const p of items) {
      const cols = [
        (p.title || "Untitled").replace(/\|/g, "\\|"),
        (p.sectionPath || "-").replace(/\|/g, "\\|")
      ];
      for (const col of extraCols) {
        cols.push(String(col.value(p) ?? "-").replace(/\|/g, "\\|"));
      }
      lines.push("| " + cols.join(" | ") + " |");
    }
    lines.push("");
  }

  pageTable("Added Pages", summary.added, [
    { header: "Created", value: p => p.createdDateTime?.slice(0, 19) || "-" },
    { header: "Modified", value: p => p.lastModifiedDateTime?.slice(0, 19) || "-" }
  ]);

  pageTable("Updated Pages", summary.updated, [
    { header: "Previous Modified", value: p => p.previousLastModifiedDateTime?.slice(0, 19) || "-" },
    { header: "New Modified", value: p => p.lastModifiedDateTime?.slice(0, 19) || "-" }
  ]);

  pageTable("Deleted Pages", summary.deleted, [
    { header: "Last Modified (local)", value: p => p.lastModifiedDateTime?.slice(0, 19) || "-" },
    { header: "HTML Path", value: p => p.htmlPath || "-" }
  ]);

  pageTable("Missing Local Files", summary.missingLocal, [
    { header: "Modified", value: p => p.lastModifiedDateTime?.slice(0, 19) || "-" }
  ]);

  if (summary.protectedSections && summary.protectedSections.length > 0) {
    lines.push(`## Protected Sections (${summary.protectedSections.length})`);
    lines.push("");
    lines.push("| Path | Error |");
    lines.push("|------|-------|");
    for (const s of summary.protectedSections) {
      lines.push(`| ${(s.path || "-").replace(/\|/g, "\\|")} | ${(s.error || "-").replace(/\|/g, "\\|")} |`);
    }
    lines.push("");
  }

  if (summary.failedSections && summary.failedSections.length > 0) {
    lines.push(`## Failed Sections (${summary.failedSections.length})`);
    lines.push("");
    lines.push("| Path | Error |");
    lines.push("|------|-------|");
    for (const s of summary.failedSections) {
      lines.push(`| ${(s.path || "-").replace(/\|/g, "\\|")} | ${(s.error || "-").replace(/\|/g, "\\|")} |`);
    }
    lines.push("");
  }

  lines.push("---");
  lines.push("*Generated by onenote-interactor*");

  await fs.writeFile(logPath, lines.join("\n"), "utf8");
  console.log(`Wrote diff log to ${logPath}`);
}

async function diffGraphNotebook(options) {
  const rootDir = options.root || options.out || path.join(process.cwd(), "exports", "graph", sanitizeSegment(options.notebook || "A"));
  const notebookJson = await readJsonIfExists(path.join(rootDir, "notebook.json"));
  const notebookName = options.notebook || notebookJson?.displayName || "A";
  const token = await getAccessToken();
  const { structure } = await getNotebookAndStructureForRoot(token, notebookName, rootDir, options);
  const manifest = await loadLocalManifest(rootDir);
  if (!manifest.generatedAt) {
    console.log("Local manifest not found; building it from existing exported files first.");
    Object.assign(manifest, await buildLocalPagesManifest({ root: rootDir }));
  }

  // Determine quick vs full scan mode
  const isFull = options.full === true;
  let cutoffTimestamp = null;
  if (!isFull && manifest.pages && Object.keys(manifest.pages).length > 0) {
    const timestamps = Object.values(manifest.pages)
      .map((p) => p.lastModifiedDateTime)
      .filter(Boolean)
      .map((t) => new Date(t).getTime());
    if (timestamps.length > 0) {
      const maxTimestamp = Math.max(...timestamps);
      // 1-hour buffer to account for Graph sync lag
      cutoffTimestamp = maxTimestamp - 60 * 60 * 1000;
      console.log(`Quick diff: cutoff ${new Date(cutoffTimestamp).toISOString()} (max manifest timestamp minus 1h)`);
    }
  }

  if (!cutoffTimestamp) {
    console.log("Full diff: scanning all pages (no manifest, --full flag, or empty manifest).");
  }

  const scanOptions = cutoffTimestamp ? { ...options, cutoffTimestamp } : options;
  const remote = await scanRemotePages(token, structure, scanOptions);
  const added = [];
  const updated = [];
  const missingLocal = [];
  const unchanged = [];
  const deleted = [];

  const verbose = options.verbose === true;
  const suspicious = [];
  for (const remotePage of Object.values(remote.pages)) {
    const localPage = manifest.pages[remotePage.id];
    if (!localPage) {
      if (verbose) console.log(`  ADDED: ${remotePage.sectionPath}/${remotePage.title}`);
      added.push(remotePage);
      continue;
    }

    const localHtmlPath = path.join(rootDir, localPage.htmlPath || "");
    const localJsonPath = path.join(rootDir, localPage.jsonPath || "");
    if (!(await pathExists(localHtmlPath)) || !(await pathExists(localJsonPath))) {
      if (verbose) console.log(`  MISSING: ${remotePage.sectionPath}/${remotePage.title}`);
      missingLocal.push(remotePage);
      continue;
    }

    const remoteTime = remotePage.lastModifiedDateTime ? new Date(remotePage.lastModifiedDateTime).getTime() : 0;
    const localTime = localPage.lastModifiedDateTime ? new Date(localPage.lastModifiedDateTime).getTime() : 0;

    if (remoteTime > localTime) {
      if (verbose) {
        console.log(
          `  UPDATED: ${remotePage.sectionPath}/${remotePage.title} ` +
          `(local: ${localPage.lastModifiedDateTime} → remote: ${remotePage.lastModifiedDateTime})`
        );
      }
      updated.push({
        ...remotePage,
        previousLastModifiedDateTime: localPage.lastModifiedDateTime
      });
      continue;
    }

    if (localTime > remoteTime) {
      // Manifest claims page is newer than Graph — suspicious, usually means
      // Graph hasn't synced the latest changes yet, OR the manifest was
      // updated with a future timestamp somehow.
      suspicious.push({
        ...remotePage,
        localLastModifiedDateTime: localPage.lastModifiedDateTime,
        reason: "manifest newer than Graph"
      });
      if (verbose) {
        console.log(
          `  SUSPICIOUS: ${remotePage.sectionPath}/${remotePage.title} ` +
          `(local: ${localPage.lastModifiedDateTime} > remote: ${remotePage.lastModifiedDateTime})`
        );
      }
    } else if (verbose) {
      console.log(
        `  unchanged: ${remotePage.sectionPath}/${remotePage.title} ` +
        `(local: ${localPage.lastModifiedDateTime}, remote: ${remotePage.lastModifiedDateTime})`
      );
    }
    unchanged.push(remotePage);
  }

  if (suspicious.length > 0) {
    console.log("");
    console.log(`⚠️  ${suspicious.length} page(s) have manifest timestamps newer than Graph (possible sync lag):`);
    for (const p of suspicious.slice(0, 20)) {
      console.log(`   • ${p.sectionPath}/${p.title} (local: ${p.localLastModifiedDateTime}, remote: ${p.lastModifiedDateTime})`);
    }
    if (suspicious.length > 20) {
      console.log(`   ... and ${suspicious.length - 20} more`);
    }
    console.log("");
  }

  // In quick mode we cannot detect deletions because we didn't scan all pages.
  // Only run deletion detection in full mode.
  if (!remote.isQuick) {
    for (const localPage of Object.values(manifest.pages || {})) {
      if (!remote.pages[localPage.id]) {
        deleted.push(localPage);
      }
    }
  }

  const summary = {
    generatedAt: nowIso(),
    notebook: notebookName,
    rootDir,
    isQuick: remote.isQuick || false,
    totals: {
      remotePages: Object.keys(remote.pages).length,
      localManifestPages: Object.keys(manifest.pages || {}).length,
      added: added.length,
      updated: updated.length,
      missingLocal: missingLocal.length,
      deleted: deleted.length,
      unchanged: unchanged.length,
      protectedSections: remote.protectedSections.length,
      failedSections: remote.failedSections.length
    },
    added,
    updated,
    missingLocal,
    deleted,
    protectedSections: remote.protectedSections,
    failedSections: remote.failedSections
  };

  const outPath = options.out || path.join(rootDir, "diff-summary.json");
  await writeJson(outPath, summary);
  await writeDiffLogMarkdown(summary, rootDir);
  console.log(`Wrote diff summary to ${outPath}`);
  console.log(JSON.stringify(summary.totals, null, 2));

  if (summary.totals.added === 0 && summary.totals.updated === 0 && summary.totals.deleted === 0 && summary.totals.missingLocal === 0) {
    console.log("");
    console.log("⚠️  No changes detected. Possible reasons:");
    console.log("   • OneNote changes haven't synced to Microsoft Graph yet (can take a few minutes)");
    console.log("   • The modified page's timestamp hasn't updated in Graph API");
    console.log("   • Pages were added to a new section that wasn't discovered (try --refresh-structure)");
    console.log("   • The local manifest is missing or out of date (run graph-manifest first)");
    console.log("");
  }

  return summary;
}

async function exportSinglePage(token, rootDir, manifest, remotePage, options = {}) {
  const sectionDir = sectionDirFromPath(rootDir, remotePage.sectionPath);
  await ensureDir(sectionDir);
  const files = pageFileSet(sectionDir, remotePage.page);
  await activeTracker?.setPhase("resync-page");
  await activeTracker?.setCurrentSection(remotePage.sectionPath);
  await activeTracker?.setCurrentPage(`${remotePage.sectionPath} :: ${remotePage.title}`);
  console.log(`Resync page: ${remotePage.sectionPath} :: ${remotePage.title}`);
  const html = await graphFetchText(token, `${graphOneNoteRoot(options)}/pages/${remotePage.id}/content`);
  await fs.writeFile(files.htmlPath, html);
  await writeJson(files.jsonPath, remotePage.page);
  await updateManifestPage(rootDir, manifest, remotePage.sectionPath, sectionDir, remotePage.page);
}

async function cleanupDeletedPages(rootDir, manifest, deletedPages) {
  const cleaned = [];
  const failed = [];
  for (const page of deletedPages) {
    const pathsToDelete = [];
    if (page.htmlPath) pathsToDelete.push(path.join(rootDir, page.htmlPath));
    if (page.jsonPath) pathsToDelete.push(path.join(rootDir, page.jsonPath));
    if (page.markdownPath) pathsToDelete.push(path.join(rootDir, page.markdownPath));
    if (page.assetDir) pathsToDelete.push(path.join(rootDir, page.assetDir));

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
          console.log(`Deleted: ${path.relative(rootDir, p)}`);
        }
      } catch (err) {
        console.error(`Failed to delete ${p}: ${err.message}`);
        failed.push({ path: p, error: err.message });
      }
    }

    if (anyDeleted || failed.length === 0) {
      cleaned.push(page);
      delete manifest.pages[page.id];
    }
  }
  return { cleaned, failed };
}

async function writeSessionLog(notebookName, rootDir, diff, exportStats, cleanupResult) {
  const logsDir = path.join(process.cwd(), "logs");
  await ensureDir(logsDir);
  const ts = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
  const logPath = path.join(logsDir, `session-${sanitizeSegment(notebookName)}-${ts}.md`);

  const lines = [];
  lines.push(`# Export Session Log — ${notebookName}`);
  lines.push("");
  lines.push(`- **Notebook:** \`${notebookName}\``);
  lines.push(`- **Root:** \`${rootDir}\``);
  lines.push(`- **Generated:** ${nowIso()}`);
  lines.push("");

  lines.push("## Summary");
  lines.push("");
  lines.push("| Metric | Value |");
  lines.push("|--------|-------|");
  lines.push(`| Added pages exported | ${exportStats.pagesExported} |`);
  lines.push(`| Export failures | ${exportStats.pagesFailed} |`);
  lines.push(`| Deleted pages detected | ${diff.totals.deleted} |`);
  lines.push(`| Deleted pages cleaned | ${cleanupResult.cleaned.length} |`);
  lines.push(`| Cleanup failures | ${cleanupResult.failed.length} |`);
  lines.push("");

  if (diff.added.length > 0) {
    lines.push(`## Added Pages (${diff.added.length})`);
    lines.push("");
    lines.push("| Title | Section Path |");
    lines.push("|-------|--------------|");
    for (const p of diff.added) {
      lines.push(`| ${(p.title || "Untitled").replace(/\|/g, "\\|")} | ${(p.sectionPath || "-").replace(/\|/g, "\\|")} |`);
    }
    lines.push("");
  }

  if (diff.updated.length > 0) {
    lines.push(`## Updated Pages (${diff.updated.length})`);
    lines.push("");
    lines.push("| Title | Section Path |");
    lines.push("|-------|--------------|");
    for (const p of diff.updated) {
      lines.push(`| ${(p.title || "Untitled").replace(/\|/g, "\\|")} | ${(p.sectionPath || "-").replace(/\|/g, "\\|")} |`);
    }
    lines.push("");
  }

  if (diff.missingLocal.length > 0) {
    lines.push(`## Missing Local Files (${diff.missingLocal.length})`);
    lines.push("");
    lines.push("| Title | Section Path |");
    lines.push("|-------|--------------|");
    for (const p of diff.missingLocal) {
      lines.push(`| ${(p.title || "Untitled").replace(/\|/g, "\\|")} | ${(p.sectionPath || "-").replace(/\|/g, "\\|")} |`);
    }
    lines.push("");
  }

  if (cleanupResult.cleaned.length > 0) {
    lines.push(`## Deleted Pages Cleaned (${cleanupResult.cleaned.length})`);
    lines.push("");
    lines.push("| Title | Section Path |");
    lines.push("|-------|--------------|");
    for (const p of cleanupResult.cleaned) {
      lines.push(`| ${(p.title || "Untitled").replace(/\|/g, "\\|")} | ${(p.sectionPath || "-").replace(/\|/g, "\\|")} |`);
    }
    lines.push("");
  }

  if (exportStats.failedPages.length > 0) {
    lines.push(`## Export Failures (${exportStats.failedPages.length})`);
    lines.push("");
    lines.push("| Title | Section Path | Error |");
    lines.push("|-------|--------------|-------|");
    for (const p of exportStats.failedPages) {
      lines.push(`| ${(p.title || "Untitled").replace(/\|/g, "\\|")} | ${(p.sectionPath || "-").replace(/\|/g, "\\|")} | ${(p.error || "-").replace(/\|/g, "\\|")} |`);
    }
    lines.push("");
  }

  if (cleanupResult.failed.length > 0) {
    lines.push(`## Cleanup Failures (${cleanupResult.failed.length})`);
    lines.push("");
    lines.push("| Path | Error |");
    lines.push("|------|-------|");
    for (const f of cleanupResult.failed) {
      lines.push(`| ${f.path.replace(/\|/g, "\\|")} | ${(f.error || "-").replace(/\|/g, "\\|")} |`);
    }
    lines.push("");
  }

  lines.push("---");
  lines.push("*Auto-generated by graph-resync*");

  await fs.writeFile(logPath, lines.join("\n"), "utf-8");
  console.log(`Wrote session log to ${logPath}`);
  return logPath;
}

async function loadExistingDiffSummary(rootDir, notebookName, maxAgeMs = 24 * 60 * 60 * 1000) {
  try {
    const diffPath = path.join(rootDir, "diff-summary.json");
    const data = await readJsonIfExists(diffPath);
    if (!data || !data.totals) return null;
    if (data.notebook !== notebookName) {
      console.log(`Existing diff is for notebook "${data.notebook}", not "${notebookName}". Recomputing.`);
      return null;
    }
    const generatedAt = data.generatedAt ? new Date(data.generatedAt).getTime() : 0;
    const ageMs = Date.now() - generatedAt;
    if (ageMs > maxAgeMs) {
      console.log(`Existing diff is ${Math.round(ageMs / 60000)}min old (max ${Math.round(maxAgeMs / 60000)}min). Recomputing.`);
      return null;
    }
    console.log(`Reusing existing diff from ${data.generatedAt} (${Math.round(ageMs / 60000)}min old).`);
    return data;
  } catch (error) {
    console.log(`Could not load existing diff: ${error.message}. Recomputing.`);
    return null;
  }
}

async function resyncGraphNotebook(options) {
  const notebookName = options.notebook || "A";
  const rootDir = options.root || options.out || path.join(process.cwd(), "exports", "graph", sanitizeSegment(notebookName));
  const token = await getAccessToken();

  let diff;
  let diffReused = false;
  if (options["use-diff"] === true) {
    diff = await loadExistingDiffSummary(rootDir, notebookName);
    if (diff) {
      diffReused = true;
    }
  }
  if (!diff) {
    diff = await diffGraphNotebook({ ...options, notebook: notebookName, root: rootDir });
  }
  const manifest = await loadLocalManifest(rootDir);
  const toExport = [...diff.added, ...diff.updated, ...diff.missingLocal];
  const forceTitles = (options["force-titles"] || "")
    .split(",")
    .map((s) => s.trim().toLowerCase())
    .filter(Boolean);

  if (forceTitles.length > 0) {
    // Also scan unchanged pages — force-export any whose title matches,
    // bypassing timestamp comparison (useful when Graph sync lag hides edits).
    for (const remotePage of diff.unchanged || []) {
      const titleLower = (remotePage.title || "").toLowerCase();
      const matches = forceTitles.some((pattern) => titleLower.includes(pattern));
      if (matches && !toExport.find((p) => p.id === remotePage.id)) {
        console.log(`Force export (title match): ${remotePage.sectionPath}/${remotePage.title}`);
        toExport.push(remotePage);
      }
    }
  }

  const stats = {
    pagesExported: 0,
    pagesFailed: 0,
    failedPages: [],
    deletedMarked: diff.deleted.length
  };

  for (const remotePage of toExport) {
    try {
      await exportSinglePage(token, rootDir, manifest, remotePage, options);
      stats.pagesExported += 1;
    } catch (error) {
      stats.pagesFailed += 1;
      stats.failedPages.push({
        id: remotePage.id,
        title: remotePage.title,
        sectionPath: remotePage.sectionPath,
        error: error.message
      });
      console.error(`Failed resync page ${remotePage.title}: ${error.message}`);
    }
  }

  // Clean up deleted pages
  const cleanupResult = await cleanupDeletedPages(rootDir, manifest, diff.deleted);
  if (cleanupResult.cleaned.length > 0) {
    console.log(`Cleaned up ${cleanupResult.cleaned.length} deleted page(s).`);
  }

  await writeLocalManifest(rootDir, manifest);
  const resyncSummary = {
    generatedAt: nowIso(),
    notebook: notebookName,
    rootDir,
    diffReused,
    diffGeneratedAt: diffReused ? diff.generatedAt : undefined,
    stats,
    cleanup: {
      deletedDetected: diff.deleted.length,
      deletedCleaned: cleanupResult.cleaned.length,
      cleanupFailures: cleanupResult.failed.length
    }
  };
  await writeJson(path.join(rootDir, "resync-summary.json"), resyncSummary);
  await writeSessionLog(notebookName, rootDir, diff, stats, cleanupResult);
  console.log(`Resync complete. Exported ${stats.pagesExported} pages, failed ${stats.pagesFailed}. Cleaned ${cleanupResult.cleaned.length} deleted pages.`);
}

async function graphStatus(options) {
  const rootDir = options.root || path.join(process.cwd(), "exports", "graph", "A");
  const structure = await readJsonIfExists(path.join(rootDir, "structure.json"));
  const countSummary = await readJsonIfExists(path.join(rootDir, "count-summary.json"));
  const exportSummary = await readJsonIfExists(path.join(rootDir, "export-summary.json"));
  const postprocessSummary = await readJsonIfExists(path.join(rootDir, "postprocess-summary.json"));
  const filesystemCounts = await getExportFilesystemCounts(rootDir);
  const resourceCounts = await getExportResourceCounts(rootDir);

  let totalSections = null;
  if (structure) {
    totalSections = flattenStructureSections(structure).length;
  }

  const status = {
    rootDir,
    cachePath: DEFAULT_CACHE_PATH,
    totalSections,
    countedTotalPages: countSummary?.stats?.totalPages ?? null,
    countedSections: countSummary?.stats?.sectionsCounted ?? null,
    countedProtectedSections: countSummary?.stats?.sectionsProtected ?? null,
    exportedSectionsCompleted: exportSummary?.stats?.sectionsCompleted ?? null,
    exportedSectionsProtected: exportSummary?.stats?.sectionsProtected ?? null,
    exportedPages: exportSummary?.stats?.pagesExported ?? null,
    htmlFiles: filesystemCounts.htmlFiles,
    markdownFiles: filesystemCounts.markdownFiles,
    jsonFiles: filesystemCounts.jsonFiles,
    assetDirs: filesystemCounts.assetDirs,
    resources: resourceCounts,
    sectionSummaries: filesystemCounts.sectionSummaries,
    postprocessedHtmlFiles: postprocessSummary?.htmlFilesProcessed ?? null,
    downloadedAssets: postprocessSummary?.downloadedAssets ?? null
  };

  console.log(JSON.stringify(status, null, 2));
}

function countMatches(value, regex) {
  return (value.match(regex) || []).length;
}

function markdownRawHtmlScore(markdown) {
  return {
    htmlTagCount: countMatches(markdown, /<(?:div|span|p|table|tbody|thead|tr|td|object|iframe|html|body)\b/gi),
    htmlCloseTagCount: countMatches(markdown, /<\/(?:div|span|p|table|tbody|thead|tr|td|object|iframe|html|body)>/gi),
    graphResourceLinks: countMatches(markdown, /https:\/\/graph\.microsoft\.com\/[^)\s"']*\/onenote\/resources\/[^)\s"']*/gi)
  };
}

function readableTextFromHtml(html) {
  const dom = new JSDOM(normalizeOneNoteHtml(html));
  return (dom.window.document.body?.textContent || "").replace(/\s+/g, " ").trim();
}

function readableTextFromMarkdown(markdown) {
  return markdown
    .replace(/!\[[^\]]*]\([^)]+\)/g, " ")
    .replace(/\[([^\]]*)]\([^)]+\)/g, "$1")
    .replace(/[`*_#>|-]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

async function auditGraphExport(options) {
  const rootDir = options.root || path.join(process.cwd(), "exports", "graph", "A");
  const pagesDir = path.join(rootDir, "pages");
  const htmlFiles = await collectFiles(pagesDir, ".html");
  const findings = [];
  const totals = {
    htmlFiles: htmlFiles.length,
    markdownFiles: 0,
    missingMarkdown: 0,
    malformedHtml: 0,
    markdownWithRawHtml: 0,
    markdownWithGraphResources: 0,
    markdownTextSuspiciouslySmall: 0,
    htmlWithGraphResources: 0,
    suspiciousPages: 0
  };

  for (const htmlPath of htmlFiles) {
    const markdownPath = `${htmlPath.slice(0, -".html".length)}.md`;
    const html = await fs.readFile(htmlPath, "utf8");
    const normalizedHtml = normalizeOneNoteHtml(html);
    const htmlTextChars = readableTextFromHtml(normalizedHtml).length;
    const htmlResourceRefs = extractResourceReferences(normalizedHtml);
    const htmlGraphResources = htmlResourceRefs.filter((ref) => isGraphOneNoteResourceUrl(ref.value)).length;
    const hasBodyClose = /<\/body>/i.test(normalizedHtml);
    const hasHtmlClose = /<\/html>/i.test(normalizedHtml);
    const normalizedChanged = normalizedHtml !== html;

    const finding = {
      htmlPath: path.relative(rootDir, htmlPath),
      markdownPath: path.relative(rootDir, markdownPath),
      htmlBytes: Buffer.byteLength(html),
      htmlLines: html.split(/\r?\n/).length,
      normalizedChanged,
      hasBodyClose,
      hasHtmlClose,
      htmlGraphResources,
      markdownExists: await pathExists(markdownPath),
      markdownBytes: null,
      markdownRawHtml: null,
      htmlTextChars,
      markdownTextChars: null,
      markdownToHtmlTextRatio: null,
      issues: []
    };

    if (!hasBodyClose || !hasHtmlClose) {
      finding.issues.push("html_missing_closing_body_or_html");
      totals.malformedHtml += 1;
    }
    if (normalizedChanged) {
      finding.issues.push("html_has_self_closing_non_void_tags");
    }
    if (htmlGraphResources > 0) {
      finding.issues.push("html_has_remote_graph_resources");
      totals.htmlWithGraphResources += 1;
    }

    if (!finding.markdownExists) {
      finding.issues.push("missing_markdown");
      totals.missingMarkdown += 1;
    } else {
      totals.markdownFiles += 1;
      const markdown = await fs.readFile(markdownPath, "utf8");
      const markdownTextChars = readableTextFromMarkdown(markdown).length;
      finding.markdownBytes = Buffer.byteLength(markdown);
      finding.markdownTextChars = markdownTextChars;
      finding.markdownToHtmlTextRatio = htmlTextChars > 0
        ? Number((markdownTextChars / htmlTextChars).toFixed(4))
        : null;
      finding.markdownRawHtml = markdownRawHtmlScore(markdown);
      if (finding.markdownRawHtml.htmlTagCount > 20 || finding.markdownRawHtml.htmlCloseTagCount > 20) {
        finding.issues.push("markdown_raw_html_leak");
        totals.markdownWithRawHtml += 1;
      }
      if (finding.markdownRawHtml.graphResourceLinks > 0) {
        finding.issues.push("markdown_has_remote_graph_resources");
        totals.markdownWithGraphResources += 1;
      }
      if (htmlTextChars >= 2000 && markdownTextChars / htmlTextChars < 0.5) {
        finding.issues.push("markdown_text_suspiciously_small");
        totals.markdownTextSuspiciouslySmall += 1;
      }
    }

    if (finding.issues.length > 0) {
      findings.push(finding);
    }
  }

  totals.suspiciousPages = findings.length;
  const report = {
    generatedAt: nowIso(),
    rootDir,
    totals,
    findings
  };
  const outPath = options.out || path.join(rootDir, "quality-report.json");
  await writeJson(outPath, report);
  console.log(`Quality report written to ${outPath}`);
  console.log(JSON.stringify(totals, null, 2));
}

async function loadKnownProtectedSectionPaths(rootDir) {
  const countSummary = await readJsonIfExists(path.join(rootDir, "count-summary.json"));
  return new Set((countSummary?.stats?.protectedSections || []).map((section) => section.path).filter(Boolean));
}

async function exportSectionPages(token, section, targetDir, stats, sectionPath = section.name, options = {}) {
  if (options.knownProtectedSectionPaths?.has(sectionPath)) {
    stats.sectionsProtected += 1;
    stats.protectedSections.push({
      sectionGroup: path.dirname(sectionPath) === "." ? null : path.dirname(sectionPath),
      section: section.name,
      reason: "known_protected_from_count_summary",
      error: null
    });
    console.log(`Skipping known protected section: ${sectionPath}`);
    return;
  }

  await activeTracker?.setCurrentSection(sectionPath);
  await activeTracker?.setPhase("export-section");
  console.log(`Exporting section: ${section.name}`);
  const safeSectionName = sanitizeSegment(section.name);
  const sectionDir = path.join(targetDir, safeSectionName);
  const sectionSummaryPath = path.join(sectionDir, "_section.json");

  await ensureDir(sectionDir);
  let pages;
  if (await pathExists(sectionSummaryPath)) {
    const summary = await readJson(sectionSummaryPath);
    pages = summary.pages || [];
    console.log(`  Resuming section from summary with ${pages.length} pages`);
  } else {
    pages = await graphPagesForSection(token, section.id, options);
    console.log(`  Found ${pages.length} pages`);

    await writeJson(sectionSummaryPath, {
      section,
      pages
    });
  }

  for (const page of pages) {
    const { htmlPath, jsonPath } = pageFileSet(sectionDir, page);
    if ((await pathExists(htmlPath)) && (await pathExists(jsonPath))) {
      stats.pagesSkipped += 1;
      continue;
    }

    console.log(`  Page: ${page.title || page.id}`);
    await activeTracker?.setCurrentPage(`${sectionPath} :: ${page.title || page.id}`);
    try {
      const html = await graphFetchText(token, `${graphOneNoteRoot(options)}/pages/${page.id}/content`);
      await fs.writeFile(htmlPath, html);
      await writeJson(jsonPath, page);
      stats.pagesExported += 1;
    } catch (error) {
      stats.pagesFailed += 1;
      stats.failedPages.push({
        section: section.name,
        pageId: page.id,
        pageTitle: page.title || page.id,
        error: error.message
      });
      console.error(`  Failed page: ${page.title || page.id} :: ${error.message}`);
    }
  }

  stats.sectionsCompleted += 1;
}

async function exportGroupRecursive(token, group, targetDir, stats, groupPath = [], options = {}) {
  const safeName = sanitizeSegment(group.name);
  const groupDir = path.join(targetDir, safeName);
  const logicalGroupPath = [...groupPath, group.name];
  await activeTracker?.setPhase("export-group");
  await activeTracker?.setCurrentSection(logicalGroupPath.join("/"));
  console.log(`Entering section group: ${group.name}`);
  await ensureDir(groupDir);

  for (const section of group.sections || []) {
    try {
      await exportSectionPages(token, section, groupDir, stats, [...logicalGroupPath, section.name].join("/"), options);
    } catch (error) {
      if (error.message.includes("403 Forbidden")) {
        stats.sectionsProtected += 1;
        stats.protectedSections.push({
          sectionGroup: group.name,
          section: section.name,
          reason: "protected_or_forbidden",
          error: error.message
        });
        console.error(`Protected section ${section.name} in group ${group.name}: ${error.message}`);
        continue;
      }
      stats.sectionsFailed += 1;
      stats.failedSections.push({
        sectionGroup: group.name,
        section: section.name,
        error: error.message
      });
      console.error(`Failed section ${section.name} in group ${group.name}: ${error.message}`);
    }
  }

  for (const child of group.sectionGroups || []) {
    await exportGroupRecursive(token, child, groupDir, stats, logicalGroupPath, options);
  }
}

async function exportGraphNotebook(options) {
  const notebookName = options.notebook || "A";
  const outDir = options.out || path.join(process.cwd(), "exports", "graph", sanitizeSegment(notebookName));
  console.log(`Starting export for notebook: ${notebookName}`);
  const token = await getAccessToken();
  console.log("Access token acquired");
  await ensureDir(outDir);
  const notebookPath = path.join(outDir, "notebook.json");
  const structurePath = path.join(outDir, "structure.json");
  const useCachedStructure = options["refresh-structure"] !== true && (await pathExists(notebookPath)) && (await pathExists(structurePath));

  let notebook;
  let structure;
  if (useCachedStructure) {
    notebook = await readJson(notebookPath);
    structure = await readJson(structurePath);
    console.log(`Using cached notebook metadata from ${outDir}`);
  } else {
    notebook = await graphFindNotebook(token, notebookName, options);
    console.log(`Notebook found: ${notebook.displayName}`);
    structure = await buildGraphStructure(token, notebook, options);
    console.log("Notebook structure fetched");
    await writeJson(notebookPath, notebook);
    await writeJson(structurePath, structure);
    console.log(`Wrote notebook metadata to ${outDir}`);
  }

  const pagesRoot = path.join(outDir, "pages");
  await ensureDir(pagesRoot);
  const stats = createRunStats();
  const knownProtectedSectionPaths = await loadKnownProtectedSectionPaths(outDir);
  await activeTracker?.setPhase("export-start");

  for (const section of structure.sections || []) {
    try {
      await exportSectionPages(token, section, pagesRoot, stats, section.name, { ...options, knownProtectedSectionPaths });
    } catch (error) {
      if (error.message.includes("403 Forbidden")) {
        stats.sectionsProtected += 1;
        stats.protectedSections.push({
          sectionGroup: null,
          section: section.name,
          reason: "protected_or_forbidden",
          error: error.message
        });
        console.error(`Protected top-level section ${section.name}: ${error.message}`);
        continue;
      }
      stats.sectionsFailed += 1;
      stats.failedSections.push({
        sectionGroup: null,
        section: section.name,
        error: error.message
      });
      console.error(`Failed top-level section ${section.name}: ${error.message}`);
    }
  }

  for (const group of structure.sectionGroups || []) {
    await exportGroupRecursive(token, group, pagesRoot, stats, [], { knownProtectedSectionPaths });
  }

  await writeJson(path.join(outDir, "export-summary.json"), {
    generatedAt: new Date().toISOString(),
    notebook: notebook.displayName,
    stats
  });
  console.log(`Exported notebook "${notebookName}" HTML to ${outDir}`);
}

async function postprocessGraphExport(options) {
  const rootDir = options.root || path.join(process.cwd(), "exports", "graph", "A");
  const pagesDir = path.join(rootDir, "pages");
  const force = options.force === true;
  const onlyReport = options["only-report"] === true;

  if (!(await pathExists(pagesDir))) {
    throw new Error(`Pages directory does not exist: ${pagesDir}`);
  }

  console.log(`Post-processing exported pages in ${pagesDir}`);
  const token = await getAccessToken();
  console.log("Access token acquired");
  const turndownService = createTurndownService();
  let htmlFiles = await collectFiles(pagesDir, ".html");
  if (onlyReport) {
    const reportPath = options.report || path.join(rootDir, "quality-report.json");
    const report = await readJson(reportPath);
    const reportFiles = new Set((report.findings || []).map((finding) => path.join(rootDir, finding.htmlPath)));
    htmlFiles = htmlFiles.filter((htmlFile) => reportFiles.has(htmlFile));
    console.log(`Using quality report ${reportPath}; selected ${htmlFiles.length} suspicious HTML files`);
  }
  await activeTracker?.setPhase("postprocess-start");
  console.log(`Found ${htmlFiles.length} HTML files`);

  const results = [];
  for (const htmlFile of htmlFiles) {
    const markdownPath = `${htmlFile.slice(0, -".html".length)}.md`;
    if (!force && (await pathExists(markdownPath))) {
      continue;
    }
    await activeTracker?.setCurrentSection(path.dirname(path.relative(rootDir, htmlFile)));
    await activeTracker?.setCurrentPage(path.relative(rootDir, htmlFile));
    await activeTracker?.setPhase("postprocess-file");
    console.log(`Post-processing: ${path.relative(rootDir, htmlFile)}`);
    results.push(await postprocessHtmlFile(token, htmlFile, turndownService, {
      skipAssets: options["skip-assets"] === true
    }));
  }

  const summary = {
    generatedAt: new Date().toISOString(),
    rootDir,
    pagesDir,
    htmlFilesProcessed: results.length,
    downloadedAssets: results.reduce((sum, item) => sum + item.downloadedAssets, 0),
    failedResources: results.flatMap((item) => item.failedResources || []),
    results
  };

  await writeJson(path.join(rootDir, "postprocess-summary.json"), summary);
  console.log(`Post-processing complete. Summary written to ${path.join(rootDir, "postprocess-summary.json")}`);
}

async function graphSync(options) {
  const notebookName = options.notebook || "A";
  const outDir = options.out || path.join(process.cwd(), "exports", "graph", sanitizeSegment(notebookName));
  await exportGraphNotebook({ ...options, notebook: notebookName, out: outDir });
  await postprocessGraphExport({ root: outDir });
}

async function generateHtmlViewer(options) {
  const rootDir = options.root || path.join(process.cwd(), "exports", "graph", "A");
  const structure = await readJsonIfExists(path.join(rootDir, "structure.json"));
  const manifest = await loadLocalManifest(rootDir);

  if (!structure) {
    throw new Error(`No structure.json found in ${rootDir}. Run graph-export first.`);
  }

  // Build flat list of pages with their file paths from manifest
  const pages = Object.values(manifest.pages || {})
    .filter((p) => p.htmlPath)
    .map((p) => ({
      id: p.id,
      title: p.title || "Untitled",
      sectionPath: p.sectionPath || "",
      htmlPath: p.htmlPath
    }));

  const notebookName = structure.name || "Notebook";

  const viewerHtml = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(notebookName)} — OneNote Viewer</title>
<style>
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; height: 100vh; overflow: hidden; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
  #layout { display: flex; height: 100vh; }
  #sidebar { width: 300px; min-width: 200px; max-width: 450px; background: #faf9f8; border-right: 1px solid #e1dfdd; display: flex; flex-direction: column; resize: horizontal; overflow: auto; }
  #sidebar-header { padding: 12px 16px; border-bottom: 1px solid #e1dfdd; background: #f3f2f1; }
  #sidebar-header h2 { margin: 0 0 8px 0; font-size: 14px; font-weight: 600; color: #323130; }
  #search { width: 100%; padding: 6px 10px; border: 1px solid #8a8886; border-radius: 4px; font-size: 13px; }
  #search:focus { outline: none; border-color: #0078d4; }
  #tree { flex: 1; overflow-y: auto; padding: 8px 0; font-size: 13px; }
  .group { margin-bottom: 2px; }
  .group-header, .section-header, .page-item { padding: 5px 16px; cursor: pointer; user-select: none; display: flex; align-items: center; gap: 6px; }
  .group-header:hover, .section-header:hover, .page-item:hover { background: #edebe9; }
  .group-header { font-weight: 600; color: #323130; }
  .section-header { padding-left: 28px; color: #605e5c; font-weight: 500; }
  .page-item { padding-left: 40px; color: #323130; text-decoration: none; }
  .page-item.active { background: #0078d4; color: #fff; }
  .page-item.active:hover { background: #106ebe; }
  .toggle { display: inline-block; width: 14px; text-align: center; font-size: 10px; color: #605e5c; transition: transform 0.15s; }
  .toggle.collapsed { transform: rotate(-90deg); }
  .children { overflow: hidden; transition: max-height 0.2s ease; }
  .children.collapsed { max-height: 0; }
  .page-count { margin-left: auto; font-size: 11px; color: #a19f9d; font-weight: 400; }
  .active .page-count { color: rgba(255,255,255,0.8); }
  #content { flex: 1; display: flex; flex-direction: column; min-width: 0; }
  #breadcrumbs { padding: 8px 16px; background: #f3f2f1; border-bottom: 1px solid #e1dfdd; font-size: 12px; color: #605e5c; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  #breadcrumbs span { color: #0078d4; cursor: pointer; }
  #breadcrumbs span:hover { text-decoration: underline; }
  iframe { flex: 1; border: none; width: 100%; background: #fff; }
  .hidden { display: none !important; }
</style>
</head>
<body>
<div id="layout">
  <div id="sidebar">
    <div id="sidebar-header">
      <h2>${escapeHtml(notebookName)}</h2>
      <input type="text" id="search" placeholder="Search pages...">
    </div>
    <div id="tree"></div>
  </div>
  <div id="content">
    <div id="breadcrumbs">Select a page from the sidebar</div>
    <iframe id="frame" src="about:blank"></iframe>
  </div>
</div>
<script>
const NOTEBOOK = ${JSON.stringify(notebookName)};
const STRUCTURE = ${JSON.stringify(structure)};
const PAGES = ${JSON.stringify(pages)};

function buildTree() {
  const tree = document.getElementById('tree');
  const flatSections = [];

  // Flatten structure sections with their paths
  (STRUCTURE.sections || []).forEach(s => flatSections.push({ ...s, path: s.name, depth: 0 }));
  function walkGroup(group, prefix) {
    const gp = [...prefix, group.name];
    (group.sections || []).forEach(s => flatSections.push({ ...s, path: gp.concat(s.name).join('/'), depth: gp.length }));
    (group.sectionGroups || []).forEach(g => walkGroup(g, gp));
  }
  (STRUCTURE.sectionGroups || []).forEach(g => walkGroup(g, []));

  // Attach pages to sections
  const sectionPages = new Map();
  PAGES.forEach(p => {
    const list = sectionPages.get(p.sectionPath) || [];
    list.push(p);
    sectionPages.set(p.sectionPath, list);
  });

  // Build DOM
  const container = document.createElement('div');

  // Top-level sections
  (STRUCTURE.sections || []).forEach(sec => {
    const pages = sectionPages.get(sec.name) || [];
    container.appendChild(createSection(sec.name, sec.name, pages, 0));
  });

  // Section groups
  (STRUCTURE.sectionGroups || []).forEach(g => {
    container.appendChild(createGroup(g, [], sectionPages));
  });

  tree.innerHTML = '';
  tree.appendChild(container);
}

function createGroup(group, prefix, sectionPages) {
  const gp = [...prefix, group.name];
  const el = document.createElement('div');
  el.className = 'group';

  const header = document.createElement('div');
  header.className = 'group-header';
  const toggle = document.createElement('span');
  toggle.className = 'toggle';
  toggle.textContent = '▼';
  header.appendChild(toggle);
  header.appendChild(document.createTextNode(group.name));
  el.appendChild(header);

  const children = document.createElement('div');
  children.className = 'children';

  let totalPages = 0;
  (group.sections || []).forEach(s => {
    const pages = sectionPages.get(gp.concat(s.name).join('/')) || [];
    totalPages += pages.length;
    children.appendChild(createSection(s.name, gp.concat(s.name).join('/'), pages, gp.length));
  });
  (group.sectionGroups || []).forEach(g => {
    const childGroup = createGroup(g, gp, sectionPages);
    children.appendChild(childGroup);
    totalPages += parseInt(childGroup.dataset.pageCount || 0);
  });

  el.dataset.pageCount = totalPages;
  const countBadge = document.createElement('span');
  countBadge.className = 'page-count';
  countBadge.textContent = totalPages;
  header.appendChild(countBadge);

  el.appendChild(children);

  header.addEventListener('click', () => {
    children.classList.toggle('collapsed');
    toggle.classList.toggle('collapsed');
  });

  return el;
}

function createSection(name, path, pages, depth) {
  const el = document.createElement('div');
  el.className = 'group';

  const header = document.createElement('div');
  header.className = 'section-header';
  const toggle = document.createElement('span');
  toggle.className = 'toggle';
  toggle.textContent = pages.length > 0 ? '▼' : '';
  header.appendChild(toggle);
  header.appendChild(document.createTextNode(name));
  el.appendChild(header);

  const children = document.createElement('div');
  children.className = 'children';

  pages.sort((a, b) => a.title.localeCompare(b.title));
  pages.forEach(p => {
    const pageEl = document.createElement('a');
    pageEl.className = 'page-item';
    pageEl.href = '#';
    pageEl.textContent = p.title;
    pageEl.dataset.path = p.sectionPath;
    pageEl.dataset.html = p.htmlPath;
    pageEl.dataset.title = p.title;
    pageEl.addEventListener('click', (e) => {
      e.preventDefault();
      loadPage(p);
    });
    children.appendChild(pageEl);
  });

  el.appendChild(children);

  header.addEventListener('click', () => {
    if (pages.length === 0) return;
    children.classList.toggle('collapsed');
    toggle.classList.toggle('collapsed');
  });

  return el;
}

function loadPage(page) {
  document.getElementById('frame').src = page.htmlPath;
  document.getElementById('breadcrumbs').innerHTML = escapeHtml(page.sectionPath) + ' / <span>' + escapeHtml(page.title) + '</span>';
  document.querySelectorAll('.page-item.active').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.page-item').forEach(el => {
    if (el.dataset.html === page.htmlPath) el.classList.add('active');
  });
  // Update URL hash for shareability
  history.replaceState(null, '', '#' + encodeURIComponent(page.htmlPath));
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

// Search
let searchDebounce;
document.getElementById('search').addEventListener('input', (e) => {
  clearTimeout(searchDebounce);
  searchDebounce = setTimeout(() => filterTree(e.target.value), 150);
});

function filterTree(query) {
  const q = query.trim().toLowerCase();
  document.querySelectorAll('.page-item').forEach(el => {
    const match = !q || el.dataset.title.toLowerCase().includes(q);
    el.classList.toggle('hidden', !match);
  });
  document.querySelectorAll('.group').forEach(g => {
    const hasVisible = g.querySelector('.page-item:not(.hidden)') !== null;
    g.classList.toggle('hidden', !hasVisible && q.length > 0);
  });
}

// Handle hash on load
window.addEventListener('DOMContentLoaded', () => {
  buildTree();
  if (location.hash) {
    const htmlPath = decodeURIComponent(location.hash.slice(1));
    const page = PAGES.find(p => p.htmlPath === htmlPath);
    if (page) loadPage(page);
  }
});
</script>
</body>
</html>`;

  const viewerPath = path.join(rootDir, "viewer.html");
  await fs.writeFile(viewerPath, viewerHtml, "utf8");
  console.log(`Generated viewer: ${viewerPath}`);
  console.log(`Open it in a browser: file://${viewerPath}`);
}

function printHelp() {
  console.log(`Usage:
  node src/onenote-interactor.js graph-login
  node src/onenote-interactor.js login
  node src/onenote-interactor.js auth
  node src/onenote-interactor.js local-index [--notebook A] [--backup-root PATH] [--out FILE]
  node src/onenote-interactor.js local-preview --input FILE.one [--out FILE]
  node src/onenote-interactor.js graph-list [--notebook A]
  node src/onenote-interactor.js graph-resolve-link --url URL
  node src/onenote-interactor.js graph-get --url GRAPH_URL
  node src/onenote-interactor.js graph-drive-export --drive-id DRIVE_ID --item-id ITEM_ID [--out DIR]
  node src/onenote-interactor.js graph-count [--notebook A] [--notebook-id ID] [--user-id USER_ID] [--out FILE]
  node src/onenote-interactor.js graph-manifest [--root DIR]
  node src/onenote-interactor.js graph-diff [--notebook A] [--notebook-id ID] [--user-id USER_ID] [--root DIR] [--full]
  node src/onenote-interactor.js graph-resync [--notebook A] [--notebook-id ID] [--user-id USER_ID] [--root DIR] [--full] [--use-diff]
  node src/onenote-interactor.js graph-export [--notebook A] [--notebook-id ID] [--user-id USER_ID] [--out DIR]
  node src/onenote-interactor.js graph-postprocess [--root DIR] [--force] [--only-report] [--skip-assets]
  node src/onenote-interactor.js graph-status [--root DIR]
  node src/onenote-interactor.js graph-audit [--root DIR] [--out FILE]
  node src/onenote-interactor.js graph-sync [--notebook A] [--out DIR]
  node src/onenote-interactor.js graph-viewer [--root DIR]
`);
}

async function main() {
  await loadLocalEnv();
  const { command, options } = parseArgs(process.argv.slice(2));
  let trackerRoot = null;
  if (["graph-count", "graph-export", "graph-postprocess", "graph-sync", "graph-diff", "graph-resync", "graph-drive-export"].includes(command)) {
    trackerRoot =
      command === "graph-postprocess" || command === "graph-diff" || command === "graph-resync"
        ? (options.root || path.join(process.cwd(), "exports", "graph", "A"))
        : command === "graph-drive-export"
          ? (options.out || path.join(process.cwd(), "exports", "drive", sanitizeSegment(options["item-id"] || "package")))
        : command === "graph-count" && options.out
          ? path.dirname(options.out)
          : (options.out || path.join(process.cwd(), "exports", "graph", sanitizeSegment(options.notebook || "A")));
    activeTracker = await createProgressTracker({
      command,
      rootDir: trackerRoot,
      notebookName: options.notebook || "A"
    });
  }

  switch (command) {
    case "graph-login":
    case "login":
    case "auth":
      await graphLogin();
      break;
    case "local-index":
      await buildLocalIndex(options);
      break;
    case "local-preview":
      await previewLocalFile(options);
      break;
    case "graph-list":
      await listGraphNotebook(options);
      break;
    case "graph-list-all":
      await listAllGraphNotebooks(options);
      break;
    case "graph-resolve-link":
      await graphResolveLink(options);
      break;
    case "graph-get":
      await graphGet(options);
      break;
    case "graph-drive-export":
      await exportDrivePackage(options);
      break;
    case "graph-count":
      await countGraphNotebookPages(options);
      break;
    case "graph-manifest":
      await buildLocalPagesManifest(options);
      break;
    case "graph-diff":
      await diffGraphNotebook(options);
      break;
    case "graph-resync":
      await resyncGraphNotebook(options);
      break;
    case "graph-export":
      await exportGraphNotebook(options);
      break;
    case "graph-postprocess":
      await postprocessGraphExport(options);
      break;
    case "graph-status":
      await graphStatus(options);
      break;
    case "graph-audit":
      await auditGraphExport(options);
      break;
    case "graph-sync":
      await graphSync(options);
      break;
    case "graph-viewer":
      await generateHtmlViewer(options);
      break;
    case "--help":
    case "-h":
    case undefined:
      printHelp();
      break;
    default:
      throw new Error(`Unknown command: ${command}`);
  }

  if (activeTracker) {
    await activeTracker.finalize("completed");
    activeTracker = null;
  }
}

const isMainModule = process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href;
if (isMainModule) {
  main().catch((error) => {
    console.error(error.message);
    const tracker = activeTracker;
    activeTracker = null;
    if (tracker) {
      tracker.finalize("failed", error).catch(() => {});
    }
    process.exitCode = 1;
  });
}

// Exported for testing
export {
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
  loadLocalManifest,
  writeLocalManifest,
  buildLocalPagesManifest,
  diffGraphNotebook,
  resyncGraphNotebook,
  cleanupDeletedPages,
  scanRemotePages,
  graphFetchAllJsonItems,
  graphFetchJson
};
