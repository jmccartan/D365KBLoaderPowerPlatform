# D365 KB Loader

A Power Apps **Code App** that reads HTML and Word documents from a SharePoint
folder, lets you preview / tweak each one, and pushes them into the Dynamics
365 Knowledgebase (`knowledgearticle` table). Every action is written to a
SharePoint list so you have a durable audit trail.

## What it does

1. **Configure** — point the app at a SharePoint site + folder + log list.
2. **Scan & pre-process** — enumerates files, classifies (`.docx`, `.html/.htm`,
   skipped), converts DOCX to HTML with `mammoth`, sanitizes with
   `sanitize-html`, derives a title.
3. **Review** — Fluent UI list with checkboxes; per-article tabs for
   **Preview**, **Edit HTML**, and **Raw source**. Title is editable.
4. **Load** — creates `knowledgearticle` rows for selected items, streaming
   progress and writing to the log list.
5. **History** — read the SharePoint log list any time from the Progress tab.

## Stack

- Power Apps Code App (`pac code`) — modern code-first Power App, deployed
  with `pac code push`.
- React 18 + TypeScript + Vite
- Fluent UI v9
- `mammoth` (DOCX → HTML), `sanitize-html`

## Project layout

```
D365KBLoader/                Code App project
  src/
    App.tsx                  Top-level wizard
    components/              ConfigPanel, ReviewPanel, ProgressPanel
    processing/pipeline.ts   DOCX/HTML → sanitized HTML
    services/                KbLoaderService interface + Mock + PowerPlatform impl
    types.ts
sharepoint/
  KbLoaderLog.list.json      Schema for the SharePoint log list
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
   imports the CLI generated. The TODO comments mark the spots.

5. **Create the log list** in your SharePoint site using the schema in
   `sharepoint/KbLoaderLog.list.json`. Default name expected by the UI:
   **KB Loader Log**.

6. **Switch to real mode**:

   In `.env.local`:

   ```
   VITE_USE_REAL_CONNECTORS=true
   ```

7. **Build & push**:

   ```powershell
   npm run build
   pac code push
   ```

   The CLI prints the published app URL.

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

Every scan/process/load/skip action writes one row to the SharePoint list
defined in **Configure**. The Progress tab reads back from the list with
**Refresh log**, so the history is durable and visible to anyone with list
access — not just the user running the load.

## Why a Code App (not classic canvas)?

Code Apps give you a real React/TS codebase, npm packages (we need `mammoth`
for DOCX), proper version control, and Fluent UI v9 — far more polish than the
canvas designer can produce. They deploy as first-class Power Apps with the
same connectors and governance as canvas apps.
