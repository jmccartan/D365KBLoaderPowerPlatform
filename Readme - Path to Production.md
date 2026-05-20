# Readme — Path to Production

The current deploy is **mock mode** (`VITE_USE_REAL_CONNECTORS=false`). The app is
in Power Apps and renders, but Configure / Scan / Load all use fake data.

This doc is the detailed playbook to flip it to **real connectors** end-to-end.
Three independent workstreams; do them in this order.

---

## Overview of work remaining

| # | Workstream | Owner | Blocking? |
|---|---|---|---|
| 1 | Create SharePoint Online + Office 365 connections in the env | You (manual, maker portal) | Yes — CLI cannot create OAuth connections |
| 2 | Register those connections as data sources | CLI | Depends on #1 |
| 3 | Refactor `PowerPlatformKbLoaderService.ts` to the new generated client shape | Code change | Depends on #2 for typings |
| 4 | Flip flag, rebuild, re-push, smoke test | CLI + browser | Depends on #3 |

Target env: **Dev-Demo-Early** (`https://orge8583733.crm.dynamics.com/`),
env id `9f0cea31-d244-eb65-a1ea-7175d8464b01`, tenant `ABSx29592600.onmicrosoft.com`.

---

## Workstream 1 — Create SharePoint + Office 365 connections

The new `power-apps` CLI does **not** create OAuth connections. You have to do
this once, by hand, signed in as a user with **Maker** role on the target env.

### 1.1 Sign in to the maker portal

Open Edge in the Playwright-managed profile (or any browser) and navigate to:

```
https://make.powerapps.com/?tenantId=ABSx29592600.onmicrosoft.com
```

Sign in as `admin@ABSx29592600.onmicrosoft.com` (the env owner — same identity
`pac auth list` shows). The current Playwright Edge profile is signed in as
"John McCartan" who does **not** have maker role on this env, so this step needs
either:

- A password for `admin@ABSx29592600` (preferred), **or**
- Grant `John McCartan` a maker role on `Dev-Demo-Early` via the Power Platform
  Admin Center, then continue with the existing browser session.

### 1.2 Pick the right environment

Top-right environment switcher → **Dev-Demo-Early**. If you don't see it,
your identity doesn't have access — fix that before continuing.

### 1.3 Create the SharePoint connection

1. Left nav → **Connections** → **+ New connection**
2. Search: `SharePoint`
3. Pick **SharePoint** (not "SharePoint with Azure AD") → **Create**
4. Sign in with an account that has access to the SharePoint sites the app
   needs to scan (typically the same admin account)

### 1.4 Create the Office 365 Outlook connection

1. **+ New connection** again
2. Search: `Office 365 Outlook` → **Create**
3. Sign in. (This is used if the README's Outlook notify-on-complete feature is
   enabled — required by the data source contract regardless.)

### 1.5 Capture the connection IDs

Back in the project shell:

```powershell
cd C:\Repos\D365KBLoaderPowerPlatform\D365KBLoader
$env:ENV_URL = 'https://orge8583733.crm.dynamics.com/'
npx power-apps list-connections
```

Expected output now includes rows for `shared_sharepointonline` and
`shared_office365`. Copy the connection IDs (32-char hex). Save them:

```powershell
$SP_CONN  = '<paste-sharepoint-id>'
$O365_CONN = '<paste-office365-id>'
```

---

## Workstream 2 — Register the data sources

```powershell
npx power-apps add-data-source -a shared_sharepointonline -c $SP_CONN  --non-interactive
npx power-apps add-data-source -a shared_office365         -c $O365_CONN --non-interactive
```

Both should finish in seconds and:

- Append entries to `power.config.json` under `connectionReferences` + an
  appropriate data sources section
- Regenerate `src/generated/services/` to add `SharePointOnlineService.ts` and
  `Office365Service.ts` (names may vary slightly — check what shows up)
- Update `src/generated/index.ts`

Verify:

```powershell
npx power-apps list-data-sources
Get-ChildItem .\src\generated\services\
```

---

## Workstream 3 — Refactor `PowerPlatformKbLoaderService.ts`

The README's wiring snippet is stale. The new generated shape is per-table
static service classes plus `MicrosoftDataverseService.executeAsync` for
unbound actions.

### 3.1 Replace the loader stubs (bottom of file, lines ~333–360)

**Delete** `loadDataverseClient`, `loadSharePointClient`, `loadEnvBaseUrl` and
replace the top of the file's imports with:

```ts
import {
  KnowledgearticlesService,
  SubjectsService,
  LanguagelocaleService,
  MicrosoftDataverseService,
  SharePointOnlineService,   // exact name TBD — check generated index.ts
  Office365Service,          // exact name TBD — check generated index.ts
} from '../generated';
```

### 3.2 Rewrite each method body

Map each old call to the new shape:

| Old | New |
|---|---|
| `dv.whoAmI?.()` | `MicrosoftDataverseService.executeAsync({ connectorOperation: { tableName: 'systemusers', operationName: 'WhoAmI', parameters: {} } })` |
| `dv.systemuser?.retrieve?.(id, ['fullname','internalemailaddress'])` | `MicrosoftDataverseService.executeAsync({ dataverseRequest: { action: 'retrieveRecord', parameters: { tableName: 'systemusers', id, select: ['fullname','internalemailaddress'] }}})` (or generate a `SystemusersService` via `add-data-source -t systemuser`) |
| `dv.org?.retrieveProvisionedLanguages?.()` | `LanguagelocaleService.getAll({ select: ['localeid','language','region'] })` |
| `dv.subject?.list?.({…})` | `SubjectsService.getAll({ select: […], filter: '…', top: 500 })` |
| `dv.knowledgearticle.list?.({…})` | `KnowledgearticlesService.getAll({ select: […], filter: '…', top: 500 })` |
| `dv.knowledgearticle.create({…})` | `KnowledgearticlesService.create({…})` |
| `dv.knowledgearticle.update(id,{…})` | `KnowledgearticlesService.update(id, {…})` |
| `dv.discovery?.listInstances?.()` | Drop it — the running code app already knows its env. Replace `loadEnvBaseUrl` callers with `window.location.origin` or a build-time constant. |
| `dv.fetch?.(url)` | `MicrosoftDataverseService.executeAsync({ dataverseRequest: { action: 'request', parameters: { method: 'GET', url } } })` — pattern may vary; check `MicrosoftDataverseService.ts` for the exact API. |
| `sp.GetAllSites?.()` | `SharePointOnlineService.executeAsync({ connectorOperation: { operationName: 'GetAllSites', parameters: {} } })` |
| `sp.GetFolderItemsByPath?.({…})` | `SharePointOnlineService.executeAsync({ connectorOperation: { operationName: 'GetFolderItemsByPath', parameters: {…} } })` |
| `sp.GetFolderFilesByPath({…})` | same pattern with `'GetFolderFilesByPath'` |
| `sp.GetFileContent({…})` | same pattern with `'GetFileContent'` |
| `sp.CreateFile({…})` | same pattern with `'CreateFile'` |

The exact `executeAsync` payload shape comes from the generated
`SharePointOnlineService.ts`. Open it and copy the type of the argument to
`executeAsync` — it'll be one of:

- `{ connectorOperation: { operationName, parameters } }`
- `{ apimOperation: { operationName, parameters } }`

Whichever the generated file uses is the right one.

### 3.3 Optional: generate a `Systemusers` service

To get strongly-typed `whoAmI` / current user lookup:

```powershell
npx power-apps add-data-source -a dataverse -t systemuser --non-interactive
```

Then use `SystemusersService.get(id)` instead of the executeAsync pattern.

### 3.4 Type-check often

After each method, run:

```powershell
npx tsc --noEmit
```

so you catch shape mismatches one method at a time instead of at the end.

---

## Workstream 4 — Flip flag, build, push, smoke test

### 4.1 Flip the flag

`D365KBLoader/.env.local`:

```env
VITE_USE_REAL_CONNECTORS=true
```

### 4.2 Build + push

```powershell
cd C:\Repos\D365KBLoaderPowerPlatform\D365KBLoader
npm run build
npx power-apps push --non-interactive
```

The CLI prints a new play URL (same app id, new version hash).

### 4.3 First-run consent

Open the play URL signed in as a user with access. Power Apps will prompt for
**Consent** on each connector (SharePoint, Office 365, Dataverse). Click
through. This only happens once per user.

### 4.4 Smoke test (mirror of README's Quick sanity checklist)

| Step | Expected |
|---|---|
| App loads, header shows env name | Env switcher resolves via real Dataverse call |
| **Configure** → language dropdown populates | `LanguagelocaleService.getAll` succeeded |
| **Configure** → subject picker populates | `SubjectsService.getAll` succeeded |
| **Scan** → Browse Site | SharePoint `GetAllSites` returns real sites |
| **Scan** → Browse Folder on a site | `GetFolderItemsByPath` returns real folders |
| **Scan** picks up files in a folder | `GetFolderFilesByPath` + `GetFileContent` work |
| **Review** shows existing-article overlap matches | `KnowledgearticlesService.getAll` with filter ran |
| **Load** creates a draft KB article | New row in `knowledgearticle` table in Dataverse |
| **Load** uploads an .xlsx report to the source SP folder | `CreateFile` ran |
| Re-running **Load** on an already-seen file updates the existing article | `KnowledgearticlesService.update` ran |

Run the SQL-style check in the env to confirm:

```
https://orge8583733.crm.dynamics.com/api/data/v9.2/knowledgearticles?$select=title,modifiedon&$top=10&$orderby=modifiedon desc
```

The most recent rows should be the test articles the smoke test created.

---

## Rollback

If real mode is broken in production and you need to revert without redeploying:

1. Flip `.env.local` back: `VITE_USE_REAL_CONNECTORS=false`
2. `npm run build`
3. `npx power-apps push --non-interactive`

Or, if you need to roll the deployed app back to a known-good version, use the
PPAC **App** → **Details** → **Versions** → **Restore** on the previous build.

---

## Known unknowns to validate during refactor

These are educated guesses based on `KnowledgearticlesService.ts`; verify by
opening the actual generated file before relying on them:

- Whether `SharePointOnlineService` uses `connectorOperation` or `apimOperation`
  in `executeAsync`.
- Whether `Office365Service` is the connector name (it might generate as
  `Office365OutlookService` or similar — check `src/generated/index.ts`).
- The exact field name for the new-row identifier returned by
  `KnowledgearticlesService.create` (`data.knowledgearticleid` is the bet).
- Whether `LanguagelocaleService.getAll` returns LCIDs in the field shape the
  old `LCID_TO_CODE` map expects, or whether it needs adaptation.

Each of these is a 30-second check in the generated `.ts` file — do it before
writing more than one method body against an assumption.
