# D365 KB Loader

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

A Power Apps **Code App** that reads HTML and Word documents from a SharePoint
folder, lets you preview / tweak each one, and pushes them into the Dynamics
365 Knowledgebase (`knowledgearticle` table). Every run produces a formatted
Excel report saved back to the source folder, so the audit trail travels with
the content — no separate list, no extra SharePoint plumbing.

A custom Fluent UI v9 blue theme, hero header, numbered stepper, and
drill-down SharePoint browsers make the experience feel like a first-class
Microsoft app, not a stock Power Apps form.

![D365 KB Loader screenshot](docs/screenshot.png)

## What it does

1. **Configure** — pick a SharePoint site and folder from a **drill-down
   browser** (or paste a URL). No more copy/pasting paths.
2. **Scan & pre-process** — enumerates files, classifies (`.docx`, `.html/.htm`,
   skipped), converts DOCX to HTML with `mammoth`, sanitizes with
   `sanitize-html`, derives a title.
3. **Review** — Fluent UI list with checkboxes; per-article tabs for
   **Preview**, **Edit HTML**, and **Raw source**. Title is editable. The
   Edit tab has a **Suggest edits with Copilot** button that proposes
   structure and clarity improvements you can Accept or Decline.
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

### Prerequisites

**Local tooling**

- **Node.js 18 LTS or newer** (Vite 5 / React 18 requirement).
- **.NET SDK 6+** — needed to install the Power Platform CLI as a global tool.
- **Power Platform CLI (`pac`)** installed locally. Install via:
  ```powershell
  dotnet tool install --global Microsoft.PowerApps.CLI.Tool
  ```
  Or use the [winget](https://learn.microsoft.com/power-platform/developer/cli/introduction#install-power-platform-cli) /
  MSI installer documented by Microsoft.

**Identity & environment**

- A signed-in account with **Power Platform admin** (or equivalent
  environment-maker) privileges in the target environment. `pac code init` and
  `pac code push` both write into Dataverse, and `pac code add-data-source`
  registers connector references — these all require admin-level permissions.
- **Code Apps preview must be enabled** on the target environment
  (Power Platform Admin Center → Environments → *your env* → Settings →
  Product → Features → **Code Apps**). Without this flag, `pac code init`
  and `pac code push` will fail.
- A Dataverse environment that includes the **Dynamics 365 Customer Service**
  (or another Knowledge-enabled) solution, so the `knowledgearticle` table
  exists.

**Licensing**

- The maker and all end-users need a **Power Apps Premium** (or per-app)
  license. Code Apps use the Dataverse and SharePoint connectors which are
  premium.

### Steps

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

### After deployment

- **Authorize connections** — first time the app runs, Power Apps will prompt
  the user to sign in to the SharePoint and Dataverse connectors. They must
  consent before the app can list sites or create articles.
- **Share the app** — from the Power Apps maker portal
  (`make.powerapps.com`), open the published app and share it with the users
  or AAD security group that should be able to run it.
- **End-user permissions** — each runner needs:
  - **Contribute** (or higher) on the target SharePoint folder, so the
    Excel report can be uploaded.
  - A Dataverse security role that grants **Create** on `knowledgearticle`
    (e.g., *Customer Service Representative* or *Knowledge Manager*).
  - A **Power Apps Premium** license (see Prerequisites).

## Troubleshooting

- **`.env.local` changes don't take effect** — Vite reads env vars at startup;
  stop and restart `npm run dev`.
- **`pac code init` fails with a feature-not-enabled error** — Code Apps
  preview isn't turned on for the environment (see Prerequisites).
- **Connector 401 / 403 after push** — open the app once in Power Apps to
  authorize each connection, or re-share the connection from the maker portal.
- **Empty folder browser** — verify the signed-in user has access to the
  SharePoint site; the connector silently returns an empty list if not.
- **Report didn't upload** — check the runner has Contribute on the folder.
  The in-app Progress tab surfaces the underlying SharePoint error.

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

## Copilot suggestions (article review)

The Edit tab on each article includes a **Suggest edits with Copilot** button.
In mock mode this runs a deterministic heuristic (adds missing `<h1>`, inserts
a Summary callout, converts "Step N" lines into ordered lists, wraps loose
text in `<p>`, tightens the title, etc.) and returns a summary + change list
in a review dialog where you Accept or Decline.

For a real LLM-backed implementation, swap the body of `suggestEdits` in
`PowerPlatformKbLoaderService.ts` to call either:

- An **Azure OpenAI custom connector** (chat completions action). Use a
  system prompt like *"You are a KB editor. Improve clarity, structure, and
  tone. Return JSON: { html, title?, summary, changes[] }"* and parse the
  response into an `ArticleSuggestion`.
- A **Dataverse AI Prompt** (Power Platform "Prompts" feature) executed via
  the Dataverse connector's `ExecuteAction`.

The dialog UI doesn't need to change — only the service method.

## Why a Code App (not classic canvas)?

Code Apps give you a real React/TS codebase, npm packages (we need `mammoth`
for DOCX), proper version control, and Fluent UI v9 — far more polish than the
canvas designer can produce. They deploy as first-class Power Apps with the
same connectors and governance as canvas apps.

## Promoting between environments (managed solutions)

`pac code push` deploys directly into one environment. To move the app across
**Dev → Test → Prod**, wrap it in a managed solution:

```powershell
# In Dev — create a solution that contains the published Code App
pac solution init --publisher-name kbloader --publisher-prefix kbl
pac solution add-reference --path .                           # add the Code App
pac solution pack --zipfile bin\KbLoader_unmanaged.zip --folder src --packagetype Unmanaged
pac solution pack --zipfile bin\KbLoader_managed.zip   --folder src --packagetype Managed

# In each downstream env — authenticate, then import
pac auth create --environment <TARGET_ENV>
pac solution import --path bin\KbLoader_managed.zip --publish-changes
```

> Tip: in Test/Prod, re-authorize the SharePoint and Dataverse connections
> after import — connection references don't carry credentials across
> environments.

## License

[MIT](LICENSE)
