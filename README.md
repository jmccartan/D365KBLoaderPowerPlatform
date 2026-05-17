# D365 KB Loader

A Power Apps **Code App** that reads HTML and Word documents from a SharePoint
folder, lets you preview / tweak each one, and pushes them into the Dynamics
365 Knowledgebase (`knowledgearticle` table). Every run produces a formatted
Excel report saved back to the source folder, so the audit trail travels with
the content — no separate list, no extra SharePoint plumbing.

A custom Fluent UI v9 blue theme, hero header, numbered stepper, and
drill-down SharePoint browsers make the experience feel like a first-class
Microsoft app, not a stock Power Apps form.

## What it does

1. **Configure** — pick a SharePoint site and folder from a **drill-down
   browser** (or paste a URL). No more copy/pasting paths.
2. **Scan & pre-process** — enumerates files, classifies (`.docx`, `.html/.htm`,
   skipped), converts DOCX to HTML with `mammoth`, sanitizes with
   `sanitize-html`, derives a title.
3. **Review** — Fluent UI list with checkboxes; per-article tabs for
   **Preview**, **Edit HTML**, and **Raw source**. Title is editable.
4. **Load** — creates `knowledgearticle` rows for selected items, streaming
   progress in the UI.
5. **Report** — a formatted `KB-Loader-Report-YYYYMMDD-hhmmss.xlsx` is
   auto-saved to the **same folder you scanned** (Summary sheet +
   color-coded, filterable Activity Log). Click **Save Excel report** any
   time to regenerate.

## Stack

- Power Apps Code App (`pac code`) — modern code-first Power App, deployed
  with `pac code push`.
- React 18 + TypeScript + Vite
- Fluent UI v9
- `mammoth` (DOCX → HTML), `sanitize-html`, `exceljs` (run report)

## Project layout

```
D365KBLoader/                Code App project
  src/
    App.tsx                  Top-level wizard, hero header, stepper
    theme.ts                 Custom blue Fluent brand ramp
    components/              ConfigPanel, ReviewPanel, ProgressPanel,
                             Stepper, BrowseSiteDialog, BrowseFolderDialog
    processing/pipeline.ts   DOCX/HTML → sanitized HTML
    reporting/report.ts      Excel run report (exceljs)
    services/                KbLoaderService interface + Mock + PowerPlatform impl
    types.ts
```

## Local dev (mock data, no env required)

```powershell
cd D365KBLoader
npm install
npm run dev
```

Browse to <http://localhost:3000>. Mock mode is on by default — you'll see four
sample files end-to-end without touching SharePoint or Dataverse.

## Wire up real connectors

1. **Authenticate** to the target environment:

   ```powershell
   pac auth create --environment <ENV_URL_OR_ID>
   ```

2. **Initialize the Code App** (only once):

   ```powershell
   cd D365KBLoader
   pac code init -n "D365 KB Loader" -d "Bulk-load SharePoint docs into D365 Knowledgebase" -b dist -f index.html
   ```

3. **Add data sources** — this generates strongly-typed clients under `src/Models/`:

   ```powershell
   pac code add-data-source -a shared_sharepointonline
   pac code add-data-source -a shared_commondataserviceforapps -t knowledgearticle
   ```

4. **Wire the generated clients** into
   `src/services/PowerPlatformKbLoaderService.ts` — replace the
   `loadSharePointClient` and `loadDataverseClient` stubs with the actual
   imports the CLI generated. The TODO comments mark the spots. The real
   service uses these SharePoint connector actions: `GetAllSites` (site
   browser), `GetFolderItemsByPath` (folder browser), `GetFolderFilesByPath`
   (scan), `GetFileContent` (download), `CreateFile` (upload the run report).

5. **Switch to real mode** — in `.env.local`:

   ```
   VITE_USE_REAL_CONNECTORS=true
   ```

6. **Build & push**:

   ```powershell
   npm run build
   pac code push
   ```

   The CLI prints the published app URL.

   > **Note:** there is no SharePoint list to create. Per-run reports are
   > written directly to the source folder as `.xlsx` files.

## Field mapping (knowledgearticle)

| App field | knowledgearticle column |
|-----------|------------------------|
| `title`   | `title`                |
| `html`    | `content`              |
| derived   | `description` ("Imported from {path}") |

Adjust `createKnowledgeArticle` in
`PowerPlatformKbLoaderService.ts` if your org uses additional required fields
(e.g. `subjectid`, `languagelocaleid`).

## Activity log

Every scan / process / load / skip action is recorded in-memory and shown in
the Progress tab. When a load finishes, a formatted Excel report
(`KB-Loader-Report-*.xlsx`) is written to the **same folder you scanned** so
the history travels with the source content — no separate list, no extra
SharePoint plumbing. You can also click **Save Excel report** at any time to
regenerate it.

## Why a Code App (not classic canvas)?

Code Apps give you a real React/TS codebase, npm packages (we need `mammoth`
for DOCX), proper version control, and Fluent UI v9 — far more polish than the
canvas designer can produce. They deploy as first-class Power Apps with the
same connectors and governance as canvas apps.
