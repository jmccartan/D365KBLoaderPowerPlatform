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

1. **Pick the target environment** — an Environment chip in the header lists
   the Power Platform / Dataverse environments you can access and shows a
   green check next to each one that has the Dynamics 365 Knowledgebase
   (`knowledgearticle` table) installed. Load is blocked until you pick an
   environment where the KB is available.
2. **Configure** — pick a SharePoint site and folder from a **drill-down
   browser**, or drag-and-drop local files (`.docx`, `.html`, `.pdf`, `.md`)
   directly. Set article defaults (language, subject/category, publish-on-load,
   duplicate behavior) and toggle recursive / incremental scanning. Save the
   whole configuration as a named **Profile** for re-use.
3. **Scan & pre-process** — enumerates files, classifies (`.docx`, `.html/.htm`,
   `.pdf`, `.md`, skipped), converts each to sanitized HTML, derives a title,
   and runs a PII / sensitive-content scan (emails, SSNs, credit-card-like
   patterns).
4. **Review** — Fluent UI list with checkboxes; per-article tabs for
   **Preview**, **Edit** (visual rich-text editor or raw HTML source), and
   **Raw source**. Title is editable, articles are drag-to-reorder, and
   keyboard shortcuts (`J/K`, arrows, `Space`, `Ctrl/Cmd+Enter`) make bulk
   review fast. The Edit tab has a **Suggest edits with Copilot** button —
   or apply Copilot suggestions in bulk across selected articles. Click
   **Scan for overlap** to compare candidates against your existing D365
   KB and flag likely duplicates.
5. **Load** — for each selected article, optionally skip / update / create a
   new `knowledgearticle` row based on the duplicate setting, honoring the
   default and per-article language, subject, and publish/draft choice.
   Streams progress in the UI with **Open in D365** links after success.
6. **Report** — a formatted `KB-Loader-Report-YYYYMMDD-hhmmss.xlsx` is
   auto-saved to the **same folder you scanned** (Summary sheet +
   color-coded, filterable Activity Log including the user identity). Click
   **Save Excel report** any time to regenerate, or **Email report…** to
   send it to a distribution list via Outlook.

### Polish

- **Dark mode** toggle in the header (persisted in localStorage).
- **Keyboard shortcuts** with a `?` help popover in the Review toolbar.
- **Drag-to-reorder** articles in the Review list.

## Stack

- Power Apps Code App (`pac code`) — modern code-first Power App, deployed
  with `pac code push`.
- React 18 + TypeScript + Vite
- Fluent UI v9 (custom blue brand theme, light + dark)
- `mammoth` (DOCX → HTML), `pdfjs-dist` (PDF text extraction),
  built-in Markdown converter, `sanitize-html`, `exceljs` (run report)

## Project layout

```
D365KBLoader/                Code App project
  src/
    App.tsx                  Top-level wizard, hero header, stepper
    theme.ts                 Custom blue Fluent brand ramp (light + dark)
    components/              ConfigPanel, ReviewPanel, ProgressPanel,
                             Stepper, EnvironmentPicker, KbDefaultsCard,
                             ProfilesBar, LocalFilesDropZone,
                             BrowseSiteDialog, BrowseFolderDialog,
                             BrowseSubjectDialog, SuggestEditsDialog,
                             RichTextEditor
    processing/pipeline.ts   DOCX/HTML/PDF/Markdown → sanitized HTML
    reporting/report.ts      Excel run report (exceljs)
    services/                KbLoaderService interface + Mock + PowerPlatform
                             impl, copilotSuggest.ts (heuristic editor),
                             overlapDetect.ts (KB duplicate scorer),
                             piiScan.ts (email / SSN / credit-card detector)
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
   service uses these connector actions:
   - **SharePoint** — `GetAllSites` (site browser), `GetFolderItemsByPath`
     (folder browser), `GetFolderFilesByPath` (scan), `GetFileContent`
     (download), `CreateFile` (upload the run report + extracted images).
   - **Dataverse** — `Create` and `Update` on `knowledgearticle` (load),
     `ListRecords` on `knowledgearticle` filtered to published / draft
     (overlap scan + duplicate detection), `RetrieveProvisionedLanguages`
     (language picker), `subject` table list (subject picker), `WhoAmI`
     and `systemuser` retrieve (current-user identity for audit logging),
     **Global Discovery Service** for environment enumeration,
     `EntityDefinitions(LogicalName='knowledgearticle')` for the
     knowledgebase availability probe.
   - **Outlook (optional)** — `SendEmailV2` for the **Email report** button.
   - **(Optional) Copilot suggestions** — swap `suggestEdits()` for an
     Azure OpenAI custom connector or a Dataverse AI Prompt action (see
     [Copilot suggestions](#copilot-suggestions-article-review) below).

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

## Article defaults &amp; per-article overrides

The **Article defaults** card on the Configure step controls:

| Setting | Effect |
|---------|--------|
| **Default language** | Sets `languagelocaleid` on every created article. Loaded from `RetrieveProvisionedLanguages` in real mode. |
| **Default subject / category** | Browseable subject tree (Dataverse `subject` table). |
| **Publish on load** | Off = articles land as **Draft**; on = published immediately (`statecode=3, statuscode=7`). |
| **Duplicate behavior** | **Skip** the candidate / **Update existing** / **Create new** when an article with the same title already exists. Match is by exact `title`. |
| **Include sub-folders** | Recurse into nested SharePoint folders during scan. |
| **Incremental** | Read the `KB-Loader-Report-*.xlsx` files in the source folder and skip anything previously loaded successfully. |

Each article in the Review step can override the language and subject in its
detail pane.

## Source flexibility

- **SharePoint** — pick a site → drill into a folder.
- **Local upload** — drag-and-drop `.docx`, `.html`, `.pdf`, `.md` onto the
  Configure step, or click to browse. Same review + load flow applies.
- **PDF** — text is extracted page-by-page with `pdfjs-dist` (worker loaded
  lazily, image-only pages flagged as warnings).
- **Markdown** — converted to HTML in-process (headings, lists, code fences,
  bold/italic/links).

## Rich-text editor

The Edit tab on each article toggles between:

- **Visual** — a Fluent-themed rich editor with Bold / Italic / Underline /
  H1 / H2 / Bullet list / Numbered list / Link. Output is re-sanitized
  through `sanitize-html` on every change so reviewer edits are safe.
- **Source** — the raw HTML textarea for power users.

## PII / sensitive-content scan

Every processed article is scanned for emails, US SSNs (`XXX-XX-XXXX`), and
credit-card-like number runs. Findings appear as a yellow callout in the
Review detail pane. When **Block load on PII** is enabled in the defaults,
articles with findings can't be selected for load until the content is
cleaned.

## Saved scan profiles

Click **Save profile…** in the Profiles bar above the Configure step to
store the current site + folder + defaults + environment under a name.
Profiles persist in `localStorage` (mock) or a Dataverse custom table
(real). One-click "Apply" restores the whole configuration so recurring
imports take seconds.

## Bulk Copilot edits

In the Review toolbar, **Apply Copilot to selected** runs `suggestEdits`
across every selected article, presenting each suggestion in turn for
Accept / Decline / Skip-all. Useful for normalizing a fresh batch of imports
in one sweep.

## Email report

After a run, click **Email report…** in the Progress step to send the
`.xlsx` to a comma-separated To list with a prefilled summary. Mock mode
logs the request; real mode uses the Outlook connector's `SendEmailV2`
with the file as a Base64 attachment.

## Dark mode

The sun/moon button in the hero header toggles between the custom light and
dark Fluent brand themes. Choice is persisted across sessions.

## Keyboard shortcuts (Review step)

| Key | Action |
|-----|--------|
| `J` or `↓` | Next article |
| `K` or `↑` | Previous article |
| `Space` | Toggle selection on the active article |
| `Ctrl/Cmd + Enter` | Trigger Load (when enabled) |
| `?` | Show the help popover |

Plus **drag-to-reorder** in the article list (HTML5 native — no extra
dependencies).

## Environment picker

Click the **Environment** chip in the header to choose which Power Platform
environment to target. The dialog lists every environment the signed-in user
can access and checks each one for the `knowledgearticle` table in parallel:

| Status | Meaning |
|--------|---------|
| ✅ **Knowledgebase available** | The Customer Service / Knowledge Management solution is installed — safe to load. |
| ⛔ **Knowledgebase not installed** | The environment has no `knowledgearticle` table; loading is blocked. |
| ⚠️ **Check failed** | The verification request errored (network, auth) — you can still try, but loading may fail. |

The Load and Scan-for-overlap buttons are disabled until a healthy
environment is selected, so you can't accidentally publish into a sandbox
that lacks the KB schema.

Behind the scenes the real service uses the
**Global Discovery Service** (`https://globaldisco.crm.dynamics.com/api/discovery/v2.0/Instances`)
to enumerate environments, and an `EntityDefinitions(LogicalName='knowledgearticle')`
metadata probe on each environment's Web API for the availability check.

## Overlap detection (existing KB scan)

The Review step has a **Scan for overlap** toolbar button. It compares each
candidate (title + first body chunk) against your existing
`knowledgearticle` rows and flags likely duplicates:

- Each flagged candidate gets an orange **N overlap(s)** badge in the list.
- The detail pane shows the top 3 matches with a relative score, the
  existing title, an excerpt, a link into D365, and human-readable reasons
  (e.g. "Title similarity 83%", "Shared keywords: vpn, mfa, setup").
- Articles with a top match ≥ 80% are **auto-deselected** so duplicates
  don't slip into the load by accident — you can re-check them if you
  intend to publish anyway.

Mock mode scores against a small built-in sample of existing KB articles
so the feature is fully exercisable offline. The real service pulls
candidates from Dataverse via `ListRecords` on `knowledgearticle` (filtered
to published / draft) and runs the same client-side scorer (`overlapDetect.ts`)
— swap that call for Dataverse Relevance Search if your KB is very large.

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
