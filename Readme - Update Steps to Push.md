# Readme — Update Steps to Push

How the D365 KB Loader code app was pushed to the demo environment
(`https://orge8583733.crm.dynamics.com/`, env id `9f0cea31-d244-eb65-a1ea-7175d8464b01`)
using the **newer npm-based `power-apps` CLI** instead of `pac code`.

The original README assumes `pac code add-data-source`, which in `pac` 2.7.4 hangs
silently with no output. The workflow below avoids it entirely and uses
`@microsoft/power-apps` (already a project devDep, so `npx power-apps …` just works).

---

## TL;DR

```powershell
cd C:\Repos\D365KBLoaderPowerPlatform\D365KBLoader

# 1. Auth (one-time per env)
pac auth create --environment https://orge8583733.crm.dynamics.com/

# 2. Point the new CLI at the env
$env:ENV_URL = 'https://orge8583733.crm.dynamics.com/'

# 3. Register data sources (Dataverse tables)
npx power-apps add-data-source -a dataverse -t knowledgearticle --non-interactive
npx power-apps add-data-source -a dataverse -t subject          --non-interactive
npx power-apps add-data-source -a dataverse -t languagelocale   --non-interactive

# 4. Build + push
npm install
npm run build
npx power-apps push --non-interactive
```

The push prints a `https://apps.powerapps.com/play/e/…/app/…` URL — that's the
deployed app.

---

## Why the new CLI

| | `pac code` (Power Platform CLI 2.7.4) | `npx power-apps` (`@microsoft/power-apps`) |
|---|---|---|
| `add-data-source` | Hangs silently, no output, ~1 s CPU — broken | Works, finishes in seconds |
| `push` | Works | Works |
| `list-connections` / `list-tables` | n/a | Built in |
| Non-interactive flags | Limited | First-class (`--non-interactive`, `-u`, `-c`, `-t`, `-a`) |
| Connection creation | Not supported | Not supported (must use maker portal — same limitation) |

`@microsoft/power-apps` ships as a project dependency, so `npx power-apps …`
resolves to the local version without a global install.

To discover commands:

```powershell
npx power-apps --help
npx power-apps add-data-source --help
```

> ⚠️ `npx --yes @microsoft/power-apps@latest --help` does **not** work
> (“could not determine executable”). Always use plain `npx power-apps`.

---

## Step-by-step

### 0. Prereqs

- Node 18+ (`node --version`)
- `pac` CLI installed and authed against the target env:
  ```powershell
  pac auth list
  pac auth select --index <n>   # if needed
  ```
  The new `power-apps` CLI piggybacks on the same Azure CLI / `pac` token cache.

### 1. Clone & install

```powershell
git clone <repo>  C:\Repos\D365KBLoaderPowerPlatform
cd C:\Repos\D365KBLoaderPowerPlatform\D365KBLoader
npm install
```

### 2. Initialize the code app (skip if `power.config.json` already exists)

```powershell
npx power-apps init   # or pac code init
```

This creates `power.config.json` and the `.power/` scaffolding.

### 3. Point the CLI at the environment

```powershell
$env:ENV_URL = 'https://orge8583733.crm.dynamics.com/'
```

Persistent across the shell session. Alternatively, pass `-u <url>` on every
command.

### 4. Register Dataverse data sources

For every table the app needs to read/write:

```powershell
npx power-apps add-data-source -a dataverse -t knowledgearticle --non-interactive
npx power-apps add-data-source -a dataverse -t subject          --non-interactive
npx power-apps add-data-source -a dataverse -t languagelocale   --non-interactive
```

Side effects:

- Appends entries under `databaseReferences."default.cds".dataSources` in
  `power.config.json`.
- (Re)generates per-table TypeScript clients in `src/generated/`:
  - `models/KnowledgearticlesModel.ts`, `SubjectsModel.ts`, `LanguagelocaleModel.ts`
  - `services/KnowledgearticlesService.ts`, `SubjectsService.ts`,
    `LanguagelocaleService.ts`, `MicrosoftDataverseService.ts`
  - Updated barrel `index.ts`
- Reuses the env's existing `shared_commondataserviceforapps` connection
  automatically.

### 5. (Optional) Register non-Dataverse connectors

The new CLI does **not** create connections — they must already exist in the
target env. Check first:

```powershell
npx power-apps list-connections
```

If you need SharePoint or Office 365 Outlook:

1. Sign in to `https://make.powerapps.com/?tenantId=<yourtenant>` as a maker on
   the target env.
2. **Connections → New connection** → create *SharePoint* and *Office 365 Outlook*.
3. Re-run `npx power-apps list-connections` and grab the connection IDs (the
   hex string in `Microsoft.PowerApps/.../connections/<id>`).
4. Register:

   ```powershell
   npx power-apps add-data-source -a shared_sharepointonline -c <SP-CONN-ID>  --non-interactive
   npx power-apps add-data-source -a shared_office365         -c <O365-CONN-ID> --non-interactive
   ```

### 6. Wire generated clients (only if `VITE_USE_REAL_CONNECTORS=true`)

The shape `power-apps` generates is **not** the same as the README's old
example. Each table gets a static class with `create/update/get/getAll/delete`
methods, and generic Dataverse calls go through `MicrosoftDataverseService`:

```ts
import {
  KnowledgearticlesService,
  SubjectsService,
  LanguagelocaleService,
  MicrosoftDataverseService,
} from '../generated';

// CRUD
await KnowledgearticlesService.create({ title: '…', /* … */ });
await KnowledgearticlesService.update(id, { statecode: 3 });
const { data } = await SubjectsService.getAll();

// Non-table / unbound action
await MicrosoftDataverseService.executeAsync({
  connectorOperation: {
    tableName: 'systemusers',
    operationName: 'WhoAmI',
    parameters: {},
  },
});
```

If you don't need real connectors yet, keep `.env.local`:

```env
VITE_USE_REAL_CONNECTORS=false
```

The app falls back to mock data and builds/deploys fine without the wiring.

### 7. Build

```powershell
npm run build
```

Runs `tsc -b && vite build`. A clean build prints chunk sizes and exits 0.

### 8. Push

```powershell
npx power-apps push --non-interactive
```

On success the CLI prints the play URL:

```
App pushed successfully. You can play your app at
https://apps.powerapps.com/play/e/<env-id>/app/<app-id>?tenantId=<tenant>
```

### 9. Verify

Open the play URL signed in as a user with access to the env. The app
loads in the Power Apps web player.

To iterate locally without re-pushing:

```powershell
npx power-apps run
```

---

## Files touched by this workflow

| Path | Owner | Notes |
|---|---|---|
| `D365KBLoader/power.config.json` | CLI | Connection refs + data source registry — edited by `add-data-source`. |
| `D365KBLoader/.power/` | CLI | Schemas (`dataSourcesInfo`) consumed by `getClient(...)`. |
| `D365KBLoader/src/generated/` | CLI | Per-table services + models. **Do not edit.** |
| `D365KBLoader/.env.local` | You | `VITE_USE_REAL_CONNECTORS` toggle. |
| `D365KBLoader/src/services/PowerPlatformKbLoaderService.ts` | You | Real-connector wiring (only needed when flipping to `true`). |

---

## Troubleshooting

- **`pac code add-data-source` hangs.** Known broken in `pac` 2.7.4. Use
  `npx power-apps add-data-source` instead.
- **`npx --yes @microsoft/power-apps@latest …` errors with “could not determine
  executable”.** Use plain `npx power-apps …` (the package is a local devDep).
- **`add-data-source` for SP/O365 says it needs a connection id.** The CLI
  doesn't create connections. Create them once in the maker portal, then pass
  `-c <id>`.
- **Push succeeds but play URL says “Sorry, we didn't find that app.”** Your
  browser is signed in as a user without access to the target env. Sign in as
  the env owner (or the identity `pac auth list` shows) or grant your user a
  role on the env.
- **TypeScript errors after `add-data-source`.** Your hand-written service file
  is calling the old client shape. Either keep `VITE_USE_REAL_CONNECTORS=false`
  or refactor to the new generated services (see step 6).
