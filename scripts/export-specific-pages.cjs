const fs = require("fs").promises;
const path = require("path");
const { PublicClientApplication } = require("@azure/msal-node");

const CLIENT_ID = process.env.ONENOTE_CLIENT_ID;
const TENANT_ID = process.env.ONENOTE_TENANT_ID || "consumers";
const CACHE_PATH = path.join(require("os").homedir(), ".config", "onenote-interactor", "msal-cache.json");

const pagesToExport = [
  { id: "0-0106ca5f69494e0629004c2401c970a0!1-BAAFF22D74E6C2C7!2585", title: "addedPage1001", sectionPath: "inbox" },
  { id: "0-d282a605501926080ffd0a459111ac85!1-BAAFF22D74E6C2C7!2585", title: "SomeNewPage1234", sectionPath: "inbox" },
  { id: "0-2ea2c99964e6fd0c36fc4079e1ab7357!1-BAAFF22D74E6C2C7!9354", title: "Build Own Coding Agent", sectionPath: "REFERENCE/A/AI Artificial Intelligence" },
  { id: "0-56db739c6e21c9478756e90a2bc9d174!1-BAAFF22D74E6C2C7!9086", title: "Drafts Ideas Random", sectionPath: "REFERENCE/A/Arts" },
  { id: "0-4c6c91af39e72d0806be7e6f390d11a1!1-BAAFF22D74E6C2C7!2789", title: "DAW Digital Audio Workstation", sectionPath: "REFERENCE/D/D д" },
  { id: "0-3d92bc4d3f4d4e0000775334684d64e5!1-BAAFF22D74E6C2C7!sef0feae1c77f4b00be6b1790419a83f8", title: "Interslavic / Medžuslovjansky jezyk", sectionPath: "REFERENCE/L/Language and Linguistics" },
  { id: "0-592441012b81560b10f58990e9f146a7!1-BAAFF22D74E6C2C7!9245", title: "addedpage123", sectionPath: "REFERENCE/P/People" }
];

const rootDir = path.join(process.cwd(), "exports", "graph", "A");

async function getAccessToken() {
  const pca = new PublicClientApplication({
    auth: { clientId: CLIENT_ID, authority: `https://login.microsoftonline.com/${TENANT_ID}` },
    cache: { cachePlugin: {
      beforeCacheAccess: async (ctx) => {
        try { ctx.tokenCache.deserialize(await fs.readFile(CACHE_PATH, "utf8")); } catch {}
      },
      afterCacheAccess: async (ctx) => {
        if (ctx.cacheHasChanged) {
          await fs.mkdir(path.dirname(CACHE_PATH), { recursive: true });
          await fs.writeFile(CACHE_PATH, ctx.tokenCache.serialize());
        }
      }
    }}
  });
  const accounts = await pca.getTokenCache().getAllAccounts();
  const silentRequest = { account: accounts[0], scopes: ["Notes.Read", "Notes.Read.All"] };
  const response = await pca.acquireTokenSilent(silentRequest).catch(() => null);
  if (response) return response.accessToken;
  const deviceCodeRequest = { scopes: ["Notes.Read", "Notes.Read.All"], deviceCodeCallback: (resp) => console.log(resp.message) };
  const deviceResponse = await pca.acquireTokenByDeviceCode(deviceCodeRequest);
  return deviceResponse.accessToken;
}

async function graphFetchText(token, url) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Graph ${res.status} for ${url}`);
  return res.text();
}

async function graphFetchJson(token, url) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Graph ${res.status} for ${url}`);
  return res.json();
}

function sanitizeSegment(str) {
  return String(str || "").replace(/[\\/:*?"<>|]/g, "_").trim();
}

async function ensureDir(dir) {
  await fs.mkdir(dir, { recursive: true });
}

async function main() {
  console.log("Acquiring token...");
  const token = await getAccessToken();
  console.log("Token acquired. Exporting 7 pages...\n");

  for (const pageInfo of pagesToExport) {
    try {
      const sectionDir = path.join(rootDir, "pages", ...pageInfo.sectionPath.split("/"));
      await ensureDir(sectionDir);
      const safeTitle = sanitizeSegment(pageInfo.title);
      const htmlPath = path.join(sectionDir, `${safeTitle}-${pageInfo.id}.html`);
      const jsonPath = path.join(sectionDir, `${safeTitle}-${pageInfo.id}.json`);

      console.log(`Exporting: ${pageInfo.sectionPath} :: ${pageInfo.title}`);
      const html = await graphFetchText(token, `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageInfo.id}/content`);
      const meta = await graphFetchJson(token, `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageInfo.id}`);

      await fs.writeFile(htmlPath, html);
      await fs.writeFile(jsonPath, JSON.stringify(meta, null, 2));
      console.log(`  -> ${htmlPath}`);
    } catch (err) {
      console.error(`FAILED: ${pageInfo.sectionPath} :: ${pageInfo.title} — ${err.message}`);
    }
  }

  console.log("\nDone.");
}

main().catch(console.error);
