# onenote-interactor

Small CLI for exploring and exporting OneNote notebooks via Microsoft Graph.

The Graph workflow is the main path for a portable export:

- notebook structure as JSON
- pages as HTML
- embedded Graph-hosted resources downloaded locally
- HTML rewritten to local asset paths
- Markdown generated next to each exported page
- resumable exports that can be rerun safely

## Install

```bash
cd onenote-interactor
pnpm install
```

## Interactive CLI

Launch the TUI for a guided, menu-driven experience:

```bash
pnpm cli
```

This provides:
- Onboarding wizard (Azure app setup, `.env.local` creation)
- Authentication with Microsoft Graph
- Notebook listing and selection
- Export, post-process, diff, resync
- Markdown-only book creation
- Stats server lifecycle management
- Sync status dashboard

## Commands

```bash
pnpm run auth:login
pnpm run graph:list
pnpm run graph:count
pnpm run graph:export
pnpm run graph:manifest
pnpm run graph:diff
pnpm run graph:resync
pnpm run graph:postprocess
pnpm run graph:status
pnpm run graph:sync
pnpm run stats:server
pnpm run stats:server:silent
pnpm run stats:server:status
pnpm run stats:server:stop
```

Equivalent direct CLI usage:

```bash
node src/onenote-interactor.js graph-login
node src/onenote-interactor.js login
node src/onenote-interactor.js auth
node src/onenote-interactor.js graph-list --notebook <NotebookName>
node src/onenote-interactor.js graph-count --notebook <NotebookName> --out exports/graph/<NotebookName>/count-summary.json
node src/onenote-interactor.js graph-export --notebook <NotebookName> --out exports/graph/<NotebookName>
node src/onenote-interactor.js graph-manifest --root exports/graph/<NotebookName>
node src/onenote-interactor.js graph-diff --notebook <NotebookName> --root exports/graph/<NotebookName>
node src/onenote-interactor.js graph-resync --notebook <NotebookName> --root exports/graph/<NotebookName>
node src/onenote-interactor.js graph-postprocess --root exports/graph/<NotebookName>
node src/onenote-interactor.js graph-status --root exports/graph/<NotebookName>
node src/onenote-interactor.js graph-sync --notebook <NotebookName> --out exports/graph/<NotebookName>
node src/stats-server.js --root exports/graph/<NotebookName> --port 9876
node scripts/create-markdown-only-book.cjs
```

## Local mode

Build a JSON index of the local OneNote Mac backup tree:

```bash
pnpm run local:index
```

Or run it manually:

```bash
node src/onenote-interactor.js local-index --notebook <NotebookName> --out exports/local-index-<NotebookName>.json
```

Preview readable strings from a local `.one` file:

```bash
node src/onenote-interactor.js local-preview \
  --input "$HOME/Library/Containers/com.microsoft.onenote.mac/Data/Library/Application Support/Microsoft User Data/OneNote/15.0/Backup/<NotebookName>/SectionName.one" \
  --out exports/preview.txt
```

This does not decode real OneNote structure yet, but it helps verify the local files contain recoverable text.

## Graph mode

Graph mode is the path to proper HTML export.

### 1. Create an Azure app registration

This tool uses delegated Microsoft Graph auth with device-code flow.

As of May 21, 2026, a plain personal Outlook account is not enough by itself for app registration. You need:

- a Microsoft account that owns the OneNote notebook
- an Azure account / Entra directory where you can create the app registration

The app can still sign in against a personal Microsoft account notebook after that.

### 1a. Register the app

In Azure Portal:

1. Open `App registrations`
2. Click `New registration`
3. Suggested name: `onenote-interactor`
4. Supported account types:
   - `Any Entra ID tenant + Personal Microsoft accounts`
5. Create the app

Copy the resulting:

- `Application (client) ID`

### 1b. Authentication settings

In the app registration:

1. Open `Authentication`
2. Enable `Allow public client flows`
3. Save

No client secret is needed for this CLI.

### 1c. API permissions

Add Microsoft Graph delegated permissions.

Minimum permissions for read-only OneNote export:

- `User.Read`
- `Notes.Read`
- `Notes.Read.All`
- `offline_access`

Optional if you later want more file/attachment-oriented workflows:

- `Files.Read`
- `Files.Read.All`

You do not need application permissions for this CLI.

### 2. Configure environment

Create a local env file:

```bash
cp .env.example .env.local
```

Then set:

```bash
ONENOTE_CLIENT_ID="your-app-client-id"
ONENOTE_TENANT_ID="consumers"
```

For a personal OneDrive-backed notebook, `consumers` is a good default.

To load it in your shell:

```bash
set -a
source .env.local
set +a
```

### 3. Sign in once and persist the token cache

Run:

```bash
pnpm run auth:login
```

The CLI prints a Microsoft device code. Open the shown URL, enter the code, approve, and return to the terminal.

The token cache is persisted locally at:

`~/.config/onenote-interactor/msal-cache.json`

That means users should not need to re-auth every single run.

If silent token reuse fails, the CLI falls back to device-code auth automatically.

### 4. Test notebook access

List the matching notebook and basic structure:

```bash
pnpm run graph:list
```

### 5. Count the exact notebook page total

This walks every section and counts pages through the API:

```bash
pnpm run graph:count
```

It writes:

- `exports/graph/<NotebookName>/count-summary.json`

This is useful for comparing:

- total pages in OneNote
- exported pages on disk
- protected sections
- remaining gaps

### 6. Export pages as HTML

```bash
pnpm run graph:export
```

This writes files under:

`exports/graph/<NotebookName>`

Expected output:

- `notebook.json`
- `structure.json`
- `export-summary.json`
- `pages/.../*.html`
- `pages/.../*.json`

Notes:

- exports are resumable
- rerunning `graph:export` skips sections already completed via `_section.json`
- if Microsoft Graph returns `403 Forbidden`, the section is recorded as likely protected/locked
- if Microsoft Graph returns `429 Too Many Requests`, the CLI now backs off and retries automatically

### 7. Download embedded resources and generate Markdown

This second-stage pass walks the exported HTML files, downloads embedded OneNote resources
such as images from Microsoft Graph, rewrites the HTML to local asset paths, and writes a
Markdown file next to each page.

```bash
pnpm run graph:postprocess
```

Expected extra output:

- `pages/.../*.assets/*`
- `pages/.../*.md`
- `postprocess-summary.json`

### 8. Check progress / resume state

```bash
pnpm run graph:status
```

This summarizes:

- total sections discovered
- counted total pages, if `graph:count` has run
- exported sections/pages
- local HTML / JSON / Markdown counts
- downloaded asset totals

### 8b. Live browser progress view

Run:

```bash
pnpm run stats:server
```

Or start silently in the background:

```bash
pnpm run stats:server:silent
```

Then open:

`http://127.0.0.1:9876`

The page shows:

- current command and phase
- whether the process is still alive
- current section and page
- HTML / Markdown / asset counts
- current percentage against the exact Graph page total
- recent live logs
- stale heartbeat detection if the worker stops updating

Lifecycle commands:

```bash
pnpm run stats:server:status   # check if running
pnpm run stats:server:stop     # stop the server
```

While tracked commands run, the CLI writes runtime data to:

- `exports/graph/<NotebookName>/.runtime/progress-state.json`
- `exports/graph/<NotebookName>/.runtime/progress.log`

### 9. Create a markdown-only copy

To produce a clean folder with **only** `.md` files (no HTML, no JSON, no assets), run:

```bash
node scripts/create-markdown-only-book.cjs
```

This interactively asks which book to export, then:

1. Copies only `.md` files to `exports/markdown/{book}/`
2. **Strips OneNote IDs** from filenames:
   - Before: `page title-0-abc...!1-USER!123.md`
   - After: `page title.md`
3. Adds a breadcrumb header to each file:
   ```markdown
   <!-- NotebookName/REFERENCE/P/People/page title.md -->
   ```
4. Generates `{book}_index.md` with:
   - Export date
   - Total markdown file count
   - Complete sorted list of all file paths

### 10. One-shot sync

```bash
pnpm run graph:sync
```

This runs:

1. `graph-export`
2. `graph-postprocess`

This is useful once auth is configured and you want a simple "continue export and hydrate assets" command.

### 10. Incremental diff and resync

After a full export, build a local manifest:

```bash
pnpm run graph:manifest
```

This writes:

- `exports/graph/<NotebookName>/pages-manifest.json`

The manifest is keyed by Graph page id and stores:

- title
- section path
- `createdDateTime`
- `lastModifiedDateTime`
- local HTML / JSON / Markdown / asset paths

Later, when the OneNote notebook may have changed, run:

```bash
pnpm run graph:diff
```

This compares the local manifest with fresh Microsoft Graph page metadata and writes:

- `exports/graph/<NotebookName>/diff-summary.json`
- `logs/diff-{notebook}-{timestamp}.md` — human-readable Markdown report

It classifies pages as:

- `added`
- `updated`
- `missingLocal`
- `deleted`
- `unchanged`

To fetch only pages that are new, updated, or missing locally:

```bash
pnpm run graph:resync
```

This writes:

- `exports/graph/<NotebookName>/resync-summary.json`
- `logs/session-{notebook}-{timestamp}.md` — session log with exports + deletions

`graph-resync` now automatically:
1. Exports added/updated/missing pages
2. **Deletes local artifacts** (HTML, JSON, MD, asset dirs) for pages deleted in OneNote
3. **Updates the manifest** to remove deleted entries
4. **Writes a session log** to `logs/` documenting the entire run

## Progress and resume behavior

- `graph-export` is safe to rerun
- completed sections are revisited if page HTML is missing
- completed Markdown files are skipped by `graph-postprocess`
- token cache is persisted between runs
- protected sections are recorded separately so they can be handled manually later
- incremental resync uses page ids and `lastModifiedDateTime` from Microsoft Graph page metadata

## Output structure

Typical output tree:

```text
exports/graph/<NotebookName>/
  notebook.json
  structure.json
  count-summary.json
  export-summary.json
  postprocess-summary.json
  pages-manifest.json
  diff-summary.json
  resync-summary.json
  pages/
    Section1/
    Section2/
    REFERENCE/
      ...
exports/markdown/<NotebookName>/
  <NotebookName>_index.md
  Section1/
  Section2/
  REFERENCE/
logs/
  diff-<NotebookName>-2026-05-22T21-39-18.md
  session-<NotebookName>-2026-05-23T02-17-05.md
```

For each exported page you should eventually see:

- `page-name-<id>.html`
- `page-name-<id>.json`
- `page-name-<id>.md`
- `page-name-<id>.assets/` when the page contains downloadable resources

The `logs/` directory contains timestamped Markdown reports from every `graph-diff` and `graph-resync` run.

## Notes and limits

- Graph export and post-processing each use device-code auth.
- `graph:export` fetches notebook structure and page HTML.
- `graph:postprocess` fetches page-linked Graph resources and rewrites HTML for local portability.
- If Microsoft Graph returns `403 Forbidden` for a section, the exporter records it as a likely protected/locked section so you can revisit it manually later.
- Some OneNote resources may still need future handling improvements depending on how Microsoft exposes them in page HTML.
- Password-protected sections are expected to require manual handling or temporary unlocking in OneNote before export.

## Release notes / OSS checklist

This repo now includes:

- persisted MSAL token cache
- explicit login command
- progress/status command
- resumable export behavior
- Graph page counting
- local asset hydration
- Markdown generation

Good next OSS improvements:

1. Add richer attachment typing and link rewriting.
2. Add optional concurrency controls / throttle flags.
3. Add a final export completeness report that compares API page totals to local files.
4. Add tests around retry, resume, and protected-section handling.
