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

<!-- Drop a 1600×900+ PNG at docs/screenshot.png to surface here.
     See docs/README.md for capture tips. -->
&nbsp;

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
npm run build
npm run dev
```

Browse to <http://localhost:3000>. Mock mode is on by default — you'll see four
sample files end-to-end without touching SharePoint or Dataverse.

## Setup &amp; install: GitHub → Power Platform environment

This walkthrough takes you from a fresh machine to a running Code App in your
target environment. Allow ~30 minutes the first time.

> **Where to run each command**
> - Steps 0–1 run from anywhere (or the repo root after cloning).
> - **Every other step, including all `npm` and `pac code` commands, runs
>   from `D365KBLoader\`** — that's the folder that holds `package.json` and
>   (after step 5) `power.config.json`. If `pac code …` errors with
>   `power.config.json not found`, you're in the wrong directory or you
>   skipped `pac code init`.
> - `dist\` is created by `npm run build`. `pac code init` only *records*
>   the build path; the folder must actually exist before `pac code push`.

> **Heads up — CLI is changing.** Microsoft is replacing `pac code …` with
> a new npm-based CLI shipped in `@microsoft/power-apps` (commands become
> `npx power-apps init`, `npx power-apps push`, etc.). The `pac code`
> commands below still work today but will be deprecated. See
> <https://learn.microsoft.com/en-us/power-apps/developer/code-apps/how-to/npm-quickstart>.

### 0. Get the source

```powershell
git clone https://github.com/jmccartan/D365KBLoaderPowerPlatform.git
cd D365KBLoaderPowerPlatform
```

### 1. Install local tooling

You need **Node.js 18 LTS+**, **.NET SDK 6+**, and the **Power Platform CLI**.

```powershell
# Verify what you have
node --version          # expect v18.x or v20.x
dotnet --list-sdks      # expect 6.x or 8.x

# Install pac CLI as a .NET global tool (recommended)
dotnet tool install --global Microsoft.PowerApps.CLI.Tool

# Or via winget
winget install Microsoft.PowerAppsCLI

# Verify
pac --version
```

If `pac` isn't on your PATH after the dotnet install, add
`%USERPROFILE%\.dotnet\tools` to your PATH and reopen PowerShell.

### 2. Verify the app locally (mock mode)

Before touching anything in Power Platform, run the app against mock data so
you can see the UI behave end-to-end.

```powershell
cd D365KBLoader
npm install
npm run build
npm run dev
```

Open <http://localhost:3000>. You should see the hero header with the
**Environment** chip, the Configure step with sample SharePoint sites, and
four mock files (`Reset-Password.html`, `VPN-Setup.docx`, …). Walk through
Configure → Scan → Review → Load to confirm the pipeline works on your
machine. Hit `Ctrl+C` to stop the dev server.

Run the unit tests:

```powershell
npm test          # one-shot
npm run test:watch  # re-run on file save
```

### 3. Prepare the target Power Platform environment

You need an **environment-maker** account on the target Dataverse environment
and **Power Platform admin** to flip the feature flag below.

1. Sign into the **Power Platform Admin Center**
   (<https://admin.powerplatform.microsoft.com>).
2. **Environments** → pick your target → **Settings** → **Product** →
   **Features**.
3. Turn **Code Apps** **On**. Save. Without this, `pac code init` errors
   with a feature-not-enabled message.
4. In the same env, confirm the **Dynamics 365 Customer Service** (or any
   Knowledge-enabled) solution is installed — this is what creates the
   `knowledgearticle` table. If it's missing, install it from
   **Resources → Dynamics 365 apps** before going further.
5. Make sure the user who will run the app has:
   - A **Power Apps Premium** (or per-app) license.
   - A Dataverse security role that grants **Create / Write** on
     `knowledgearticle` (e.g. *Customer Service Representative*,
     *Knowledge Manager*, or a custom role).

### 4. Authenticate the pac CLI

Grab the environment's URL from PPAC (something like
`https://contoso.crm.dynamics.com`) and run:

```powershell
pac auth create --environment https://contoso.crm.dynamics.com
```

A browser window opens — sign in with the env-maker account. Confirm:

```powershell
pac auth list      # shows the active profile
pac org who        # confirms which env you're targeting
```

### 5. Initialize the Code App

This step registers the app in the environment and writes a `.power/`
project folder with connection-reference metadata.

```powershell
cd D365KBLoader
pac code init --displayName "D365 KB Loader" --description "Bulk-load SharePoint docs into D365 Knowledgebase" --buildPath dist --fileEntryPoint index.html
```

> **Heads up:** flag names changed in pac CLI 2.6 — use `--displayName` (not
> `--name`), `--buildPath` (not `--build-folder`), and `--fileEntryPoint`
> (not `--entry-file`). Run `pac code init --help` to confirm what your
> version expects.

Commit the new `.power/` folder to git so collaborators inherit the same
app identity.

### 6. Add connector data sources

Run from `D365KBLoader\` (where `power.config.json` now lives).

Each of these registers a connection reference and generates a typed client
under `src/Models/` that the real service will import. Run **all three**
(Outlook is optional but powers the Email-report feature):

```powershell
cd D365KBLoader       # if you aren't already there
pac code add-data-source -a shared_sharepointonline
pac code add-data-source -a shared_commondataserviceforapps -t knowledgearticle
pac code add-data-source -a shared_office365              # for SendEmailV2
```

For the **subject** picker and language list, also expose the `subject` and
`languagelocale` tables:

```powershell
pac code add-data-source -a shared_commondataserviceforapps -t subject
pac code add-data-source -a shared_commondataserviceforapps -t languagelocale
```

After running these, check that `src/Models/` (or whatever folder the CLI
chose for your `pac` version) contains the generated client classes.

### 7. Wire the generated clients into the real service

Open `src/services/PowerPlatformKbLoaderService.ts` and find the two helpers
at the bottom of the file:

```ts
async function loadSharePointClient(): Promise<any> { throw new Error(...) }
async function loadDataverseClient(): Promise<any>  { throw new Error(...) }
```

Replace each `throw` with a real import. The exact module names depend on
your `pac` version — read the generated files to confirm. Example:

```ts
async function loadSharePointClient() {
  const mod = await import('../Models/SharePointOnlineService');
  return new mod.SharePointOnlineService();
}

async function loadDataverseClient() {
  const ka  = await import('../Models/knowledgearticleService');
  const sub = await import('../Models/subjectService');
  const lang= await import('../Models/languagelocaleService');
  return {
    knowledgearticle: ka.knowledgearticleService,
    subject:          { list: sub.subjectService.list },
    org:              { retrieveProvisionedLanguages: lang.languagelocaleService.retrieveProvisionedLanguages },
    whoAmI:           ka.knowledgearticleService.whoAmI,
    systemuser:       ka.knowledgearticleService.systemuser,
    discovery:        ka.knowledgearticleService.discovery,
    fetch:            (url: string) => fetch(url, { headers: { Authorization: `Bearer ${await ka.knowledgearticleService.getToken()}` } }),
  };
}

async function loadEnvBaseUrl() {
  // Most generated clients expose the env URL — example field name; check yours.
  return (await import('../Models/knowledgearticleService')).knowledgearticleService.environmentUrl;
}
```

Tip: the connector-action names referenced in the methods above
(`GetAllSites`, `GetFolderItemsByPath`, `CreateFile`, etc.) come straight
from the connector schema. If your generated client wraps them under
slightly different names (camelCased, etc.), update the method bodies in
`PowerPlatformKbLoaderService.ts` accordingly.

### 8. Switch the app to real mode

Create `D365KBLoader/.env.local`:

```
VITE_USE_REAL_CONNECTORS=true
```

Restart `npm run dev` if it's running — Vite only reads env vars at startup.

### 9. Build &amp; push

Run from `D365KBLoader\`. `npm run build` produces the `dist\` folder that
`pac code push` uploads; without it, push has nothing to publish.

```powershell
cd D365KBLoader       # if you aren't already there
npm run build
pac code push
```

The CLI prints the **published app URL**. Open it in a browser — you'll be
prompted to authorize the SharePoint and Dataverse connections the first
time. Allow them.

### 10. Share with end-users

1. Go to <https://make.powerapps.com> and switch to the target environment
   (top-right env picker).
2. **Apps** → find **D365 KB Loader** → click the **⋯** menu → **Share**.
3. Add the AAD user(s) or security group, set **Co-owner** or **User**, and
   save.

Each end-user, on first run, will be prompted to authorize the connectors
under their own identity.

### 11. (Optional) Promote between environments with a managed solution

`pac code push` deploys directly into one environment. To move the app from
**Dev → Test → Prod**:

```powershell
# In Dev
pac solution init --publisher-name kbloader --publisher-prefix kbl
pac solution add-reference --path .                # add the Code App
pac solution pack --zipfile bin\KbLoader_managed.zip --folder src --packagetype Managed

# In each downstream env
pac auth create --environment <TARGET_ENV>
pac solution import --path bin\KbLoader_managed.zip --publish-changes
```

After import in Test/Prod, re-authorize the SharePoint and Dataverse
connections — connection references don't carry credentials across
environments.

### Quick sanity checklist

- [ ] `pac org who` returns the right environment
- [ ] `src/Models/` contains generated SharePoint + Dataverse clients
- [ ] `.env.local` has `VITE_USE_REAL_CONNECTORS=true`
- [ ] `npm run build` exits cleanly
- [ ] `pac code push` printed a URL
- [ ] App opens, authorizes connectors, the **Environment** chip shows a
      green check for the chosen env
- [ ] Scan returns real SharePoint files
- [ ] Load creates a `knowledgearticle` row (verify in
      `https://<env>.crm.dynamics.com` → **Knowledge Articles**)

---

## Troubleshooting

| Symptom | Likely cause &amp; fix |
|---------|------------------|
| `pac auth create` opens browser then errors | Pop-up blocker or stale token. Try `pac auth clear` then re-run. |
| `pac code init` says "Feature not enabled" | Flip **Code Apps** to On in PPAC → Settings → Product → Features. |
| `pac code add-data-source` fails with `power.config.json not found` | You're not in `D365KBLoader\`, or you skipped `pac code init`. `cd D365KBLoader` and confirm `power.config.json` exists; if not, re-run step 5. |
| `pac code add-data-source` fails with "User does not have permission" | Your account isn't an environment-maker. Have an admin grant the *Environment Maker* role in PPAC. |
| TypeScript build errors after running `add-data-source` | Generated client names don't match the imports in `PowerPlatformKbLoaderService.ts`. Open `src/Models/` and reconcile the class names. |
| `.env.local` changes ignored | Vite reads env vars at startup — restart `npm run dev` / re-run `npm run build`. |
| `pac code push` succeeds, but app shows blank screen | Check the browser console; the most common cause is a generated-client import path mismatch. |
| Connector 401 / 403 after push | Open the app once in <https://make.powerapps.com> to authorize each connection under the runner's identity. |
| Environment chip shows ⚠ "Check failed" for every env | The signed-in user can't reach the Global Discovery Service. Confirm the Dataverse connection in the maker portal has the user consented. |
| Environment shows ⛔ "Knowledgebase not installed" | The selected env doesn't have the Customer Service / Knowledge Management solution. Install it from PPAC → Resources → Dynamics 365 apps, or pick a different env. |
| Empty SharePoint folder browser | Either the user lacks access, or the connection is targeting the wrong tenant. Check the SharePoint connection in `make.powerapps.com → Connections`. |
| Excel report didn't upload | Runner needs **Contribute** on the SharePoint target folder. The in-app Progress tab surfaces the underlying error. |
| Load fails with "Required field missing: subjectid / languagelocaleid" | Your environment marks these as required. Set defaults in the **Article defaults** card before scanning. |
| PDF text comes back garbled or empty | Image-only PDF — the app emits a per-page warning. Run OCR on the source first. |
| Email-report button does nothing | Outlook data source wasn't added. Run `pac code add-data-source -a shared_office365` then re-push. |

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

See **Step 11** of the [Setup &amp; install](#setup--install-github--power-platform-environment)
guide above.

## Quick start (TL;DR)

For an experienced Power Platform dev who's already done the full walkthrough
once and just wants to redeploy / spin up a fresh env. Skips explanation —
see [Setup &amp; install](#setup--install-github--power-platform-environment)
for the full version.

**Prereqs (one-time per machine):** Node 18+, .NET SDK 6+, `pac` CLI,
**Code Apps** toggled On in PPAC for the target env, signed-in user has
Environment Maker + write access to `knowledgearticle`.

Each command in its own copy block — click the copy button on whichever
step you need.

**0. Clone**

```powershell
git clone https://github.com/jmccartan/D365KBLoaderPowerPlatform.git
```

```powershell
cd D365KBLoaderPowerPlatform\D365KBLoader
```

**1. Auth to the target env** (replace `<your-env>` with your env subdomain)

```powershell
pac auth create --environment https://<your-env>.crm.dynamics.com
```

**2. Install + build**

```powershell
npm install
```

```powershell
npm run build
```

**3. Register the app**

```powershell
pac code init --displayName "D365 KB Loader" --description "Bulk-load SharePoint docs into D365 Knowledgebase" --buildPath dist --fileEntryPoint index.html
```

**4. Add connectors**

```powershell
pac code add-data-source -a shared_sharepointonline
```

```powershell
pac code add-data-source -a shared_commondataserviceforapps -t knowledgearticle
```

```powershell
pac code add-data-source -a shared_commondataserviceforapps -t subject
```

```powershell
pac code add-data-source -a shared_commondataserviceforapps -t languagelocale
```

```powershell
pac code add-data-source -a shared_office365
```

**5. Flip to real connectors and push**

```powershell
"VITE_USE_REAL_CONNECTORS=true" | Out-File -Encoding ascii .env.local
```

```powershell
npm run build
```

```powershell
pac code push
```

`pac code push` prints the published app URL. Open it once in
<https://make.powerapps.com> to authorize each connection, then share the
app with end-users from **Apps → ⋯ → Share**.

> Run **every** command above from `D365KBLoader\`. If `pac code …` says
> `power.config.json not found`, you're in the wrong directory or skipped
> step 3.

## License

[MIT](LICENSE)
