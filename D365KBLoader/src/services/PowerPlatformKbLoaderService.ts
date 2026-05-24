import type { KbLoaderService } from './KbLoaderService';
import type { SourceFile, ProcessedArticle, KbConfig, LogEntry, SharePointSite, FolderItem, ReportResult, ArticleSuggestion, ExistingKbArticle, OverlapMatch, PowerPlatformEnvironment, KbLanguage, KbSubject, KbUser, SavedScanProfile } from '../types';
import { classify } from '../processing/pipeline';
import { buildReportWorkbook } from '../reporting/report';
import { buildMockSuggestion } from './copilotSuggest';
import { scoreOverlaps } from './overlapDetect';

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function assertGuid(name: string, value: string): string {
  if (!GUID_RE.test(value)) throw new Error(`${name} must be a GUID, got: ${value}`);
  return value;
}

function assertLcid(name: string, value: string | number): string {
  const s = String(value);
  if (!/^\d{1,6}$/.test(s)) throw new Error(`${name} must be a numeric LCID, got: ${s}`);
  return s;
}

/**
 * Real implementation that uses the generated Power Apps Code App SDK clients.
 *
 * After running:
 *   pac code add-data-source -a shared_sharepointonline
 *   pac code add-data-source -a shared_commondataserviceforapps -t knowledgearticle
 *
 * the CLI will create generated client classes under `src/Models/` (or similar).
 * Replace the dynamic imports below with the concrete imports it generated and
 * remove the runtime `throw` guards.
 *
 * Connector action references (current Power Platform connector schema):
 *   SharePoint:
 *     - GetAllSites / GetRootSiteCollection -> discover sites (Browse Site dialog)
 *     - GetFolderItemsByPath  -> list sub-folders (Browse Folder dialog)
 *     - GetFolderFilesByPath  -> list files in folder
 *     - GetFileContent         -> download bytes (FileIdentifier)
 *     - CreateFile             -> upload the per-run .xlsx report into the source folder
 *   Dataverse:
 *     - CreateRecord on table 'knowledgearticle'
 *       Required fields: title, content (HTML), description (optional), articlepublicnumber (auto)
 */
export class PowerPlatformKbLoaderService implements KbLoaderService {
  async getCurrentUser(): Promise<KbUser> {
    const dv = await loadDataverseClient();
    const who = await dv.whoAmI?.();
    const u = await dv.systemuser?.retrieve?.(who?.UserId, ['fullname', 'internalemailaddress']);
    return {
      id: who?.UserId ?? 'unknown',
      displayName: u?.fullname ?? 'Current user',
      email: u?.internalemailaddress,
    };
  }

  async listLanguages(): Promise<KbLanguage[]> {
    // Use the generated LanguagelocaleService — fast and works without admin
    // metadata endpoints. Filter to enabled languages (statecode 0 = Enabled).
    const dv = await loadDataverseClient();
    const res = await dv.language?.list?.({
      $select: 'languagelocaleid,localeid,code,name,region',
      $filter: 'statecode eq 0',
      $top: 500,
    });
    const rows: any[] = res?.value ?? [];
    if (!rows.length) {
      // Fallback to a sensible default so the dropdown is never empty
      return [{ id: '1033', code: 'en-US', displayName: 'English (United States)' }];
    }
    return rows.map(r => ({
      id: String(r.localeid ?? r.languagelocaleid),
      code: r.code ?? LCID_TO_CODE[r.localeid] ?? String(r.localeid),
      displayName: r.name ?? LCID_NAME[r.localeid] ?? `Locale ${r.localeid}`,
    } as KbLanguage));
  }

  async listSubjects(parentId?: string): Promise<KbSubject[]> {
    const dv = await loadDataverseClient();
    if (parentId !== undefined) assertGuid('parentId', parentId);
    const res = await dv.subject?.list?.({
      $select: 'subjectid,title,_parentsubject_value',
      $filter: parentId ? `_parentsubject_value eq ${parentId}` : '_parentsubject_value eq null',
      $top: 500,
    });
    return (res?.value ?? []).map((s: any) => ({
      id: s.subjectid,
      name: s.title,
      path: s.title,
      parentId: s._parentsubject_value ?? undefined,
    } as KbSubject));
  }

  async findArticleByTitle(title: string, environment?: PowerPlatformEnvironment): Promise<ExistingKbArticle | undefined> {
    const dv = await loadDataverseClient(environment);
    const normalized = title.trim().replace(/\s+/g, ' ');
    // Reject control chars (these would never appear in a real KB title and
    // could be used to construct malformed OData literals).
    if (/[\u0000-\u001F\u007F]/.test(normalized) || normalized.length === 0 || normalized.length > 500) {
      return undefined;
    }
    // OData v4 string literals only require single-quote doubling. Other
    // characters (%, \, _) are not special inside literals — see OData ABNF.
    const escaped = normalized.replace(/'/g, "''").toLowerCase();
    // Case-insensitive whitespace-tolerant match — OData supports tolower() and
    // trim() functions on string fields.
    const res = await dv.knowledgearticle.list?.({
      $select: 'knowledgearticleid,title,modifiedon',
      $filter: `tolower(trim(title)) eq '${escaped}'`,
      $top: 1,
    });
    const r = res?.value?.[0];
    if (!r) return undefined;
    const envBase = await loadEnvBaseUrl(environment);
    const url = envBase
      ? `${envBase.replace(/\/$/, '')}/main.aspx?etn=knowledgearticle&id=${r.knowledgearticleid}&pagetype=entityrecord`
      : undefined;
    return { id: r.knowledgearticleid, title: r.title, modifiedOn: r.modifiedon, url };
  }

  async updateKnowledgeArticle(existingId: string, article: ProcessedArticle, environment?: PowerPlatformEnvironment): Promise<void> {
    const dv = await loadDataverseClient(environment);
    await dv.knowledgearticle.update(existingId, {
      title: article.title,
      content: article.html,
      description: `Updated from ${article.source.path}`,
      ...(article.languageId ? { 'languagelocaleid@odata.bind': `/languagelocale(${assertLcid('languageId', article.languageId)})` } : {}),
      ...(article.subjectId ? { 'subjectid@odata.bind': `/subjects(${assertGuid('subjectId', article.subjectId)})` } : {}),
    });
  }

  /**
   * The Code App is bound to a single Dataverse environment at build time.
   * Discovery isn't exposed via the generated client, so we synthesize a
   * single-environment list from the build-time config. This keeps the
   * environment picker functional for demos.
   */
  async listEnvironments(): Promise<PowerPlatformEnvironment[]> {
    try {
      const dv = await loadDataverseClient();
      const res = await dv.discovery?.listInstances?.();
      if (res?.value?.length) {
        return res.value.map((i: any) => ({
          id: i.EnvironmentId ?? i.Id ?? i.UrlName,
          displayName: i.FriendlyName ?? i.Name,
          url: i.Url,
          region: i.Region,
          isDefault: !!i.IsDefault,
          knowledgebaseStatus: 'unknown' as const,
        } as PowerPlatformEnvironment));
      }
    } catch {
      // Fall through to single-env synthesis below.
    }
    const envUrl = typeof window !== 'undefined' ? window.location.origin : '';
    return [{
      id: 'current',
      displayName: 'Current environment',
      url: envUrl,
      region: 'prod',
      isDefault: true,
      knowledgebaseStatus: 'unknown' as const,
    }];
  }

  /**
   * Check whether the knowledgearticle table is present and queryable by doing
   * a 1-row probe via the generated client. If it succeeds the KB is present;
   * any error means missing / no permission / connection issue.
   */
  async checkKnowledgebase(env: PowerPlatformEnvironment): Promise<PowerPlatformEnvironment> {
    try {
      const dv = await loadDataverseClient();
      await dv.knowledgearticle.list?.({ $select: 'knowledgearticleid', $top: 1 });
      return { ...env, knowledgebaseStatus: 'present' };
    } catch (e: any) {
      const msg = String(e?.message ?? e);
      if (/not found|404|EntityType/i.test(msg)) return { ...env, knowledgebaseStatus: 'missing' };
      return { ...env, knowledgebaseStatus: 'error', knowledgebaseError: msg };
    }
  }

  async listSites(): Promise<SharePointSite[]> {
    // The SharePoint Online connector exposes site discovery via
    // `GetAllSites` (or, depending on connector version, `GetSitesAvailableForCurrentUser`).
    // Map each returned site to { id, name, url }.
    const sp = await loadSharePointClient();
    const res = await sp.GetAllSites?.();
    if (!res?.value) {
      throw new Error('SharePoint connector did not return any sites. Verify the connection has site-discovery permission.');
    }
    return res.value.map((s: any) => ({
      id: s.Id ?? s.Url,
      name: s.DisplayName ?? s.Title ?? s.Name,
      url: s.Url,
      description: s.Description,
    } as SharePointSite));
  }

  async listFolders(siteUrl: string, folderPath: string): Promise<FolderItem[]> {
    // The SharePoint connector returns folder + file entries together; filter to folders.
    // Most connector versions expose `GetFolderItemsByPath` (preferred) or
    // `GetFolderMetadataByPath` + a children listing.
    const sp = await loadSharePointClient();
    const path = folderPath && folderPath !== '/' ? folderPath : '/';
    const res = await sp.GetFolderItemsByPath?.({
      dataset: siteUrl,
      folderPath: path,
    });
    return (res?.value ?? [])
      .filter((it: any) => it.IsFolder)
      .map((it: any) => ({
        name: it.Name ?? it.FilenameWithExtension,
        path: it.FullPath ?? it.Path,
        hasChildren: it.HasChildren,
      } as FolderItem));
  }

  async listFiles(config: KbConfig): Promise<SourceFile[]> {
    const sp = await loadSharePointClient();
    const res = await sp.GetFolderFilesByPath({
      dataset: config.siteUrl,
      folderPath: config.folderPath,
      $top: 200
    });
    return (res.value ?? [])
      .filter((f: any) => !f.IsFolder)
      .map((f: any) => ({
        id: f.Identifier ?? f.ItemId ?? f.UniqueId,
        name: f.FilenameWithExtension ?? f.Name,
        path: f.FullPath ?? f.Path,
        size: f.Size ?? 0,
        modified: f.TimeLastModified ?? new Date().toISOString(),
        kind: classify(f.FilenameWithExtension ?? f.Name)
      } as SourceFile));
  }

  async downloadFile(file: SourceFile): Promise<ArrayBuffer> {
    const sp = await loadSharePointClient();
    const res = await sp.GetFileContent({ dataset: '', file: file.id });
    // The connector returns a Blob/Stream. Convert as the SDK exposes it.
    if (res instanceof ArrayBuffer) return res;
    if ((res as any).arrayBuffer) return await (res as any).arrayBuffer();
    return new TextEncoder().encode(String(res)).buffer;
  }

  async uploadImage(name: string, bytes: ArrayBuffer, contentType: string): Promise<string> {
    // TODO(real connector): Upload DOCX-extracted images into an `images/`
    // sub-folder beneath the selected SharePoint source folder so the converted
    // article can reference public URLs instead of inlined data URIs.
    //
    // Expected SharePoint Online connector flow:
    //   1) Ensure `${config.folderPath}/images` exists (Create new folder / EnsureFolderPath).
    //   2) Call CreateFile with dataset=siteUrl, folderPath=`<source>/images`,
    //      name, and body=bytes.
    //   3) Return the resulting server-relative / web URL from the connector
    //      response so mammoth can swap the inline image for the hosted URL.
    //
    // Until the generated client is wired up, fall back to the existing data URL
    // behavior so real-mode scans still produce usable previews.
    return `data:${contentType};base64,${arrayBufferToBase64(bytes)}`;
  }

  async listProfiles(): Promise<SavedScanProfile[]> {
    // TODO(real connector): Store profile rows in a lightweight Dataverse custom
    // table (e.g. jm_kbloaderprofile with name, config JSON, environment id,
    // owner) or in an MSAL-cached user-preferences payload so profiles roam per
    // maker without requiring extra infrastructure.
    return readProfilesFallback();
  }

  async saveProfile(profile: SavedScanProfile): Promise<SavedScanProfile> {
    // TODO(real connector): Upsert the profile row by id into Dataverse, scoped
    // to the current user. Fallback keeps real-mode demos functional until the
    // custom table and generated client are added.
    const next = { ...profile, name: profile.name.trim() };
    const profiles = readProfilesFallback().filter(existing => existing.id !== next.id);
    profiles.push(next);
    writeProfilesFallback(profiles);
    return next;
  }

  async deleteProfile(id: string): Promise<void> {
    // TODO(real connector): Delete the corresponding Dataverse custom-table row.
    writeProfilesFallback(readProfilesFallback().filter(profile => profile.id !== id));
  }

  async createKnowledgeArticle(article: ProcessedArticle, config?: KbConfig, environment?: PowerPlatformEnvironment): Promise<{ id: string; url?: string }> {
    const dv = await loadDataverseClient(environment);
    const langId = article.languageId ?? config?.defaultLanguageId;
    const subjId = article.subjectId ?? config?.defaultSubjectId;
    const publish = article.publish ?? config?.publishOnLoad ?? false;

    const created = await dv.knowledgearticle.create({
      title: article.title,
      content: article.html,
      description: `Imported from ${article.source.path}`,
      // 3 = Published (statecode 3, statuscode 7); leaving unset = Draft.
      ...(publish ? { statecode: 3, statuscode: 7 } : {}),
      ...(langId ? { 'languagelocaleid@odata.bind': `/languagelocale(${assertLcid('languageId', langId)})` } : {}),
      ...(subjId ? { 'subjectid@odata.bind': `/subjects(${assertGuid('subjectId', subjId)})` } : {}),
    });
    const id = created.knowledgearticleid ?? created.id;
    const envBase = await loadEnvBaseUrl(environment);
    const url = envBase
      ? `${envBase.replace(/\/$/, '')}/main.aspx?etn=knowledgearticle&id=${id}&pagetype=entityrecord`
      : undefined;
    return { id, url };
  }

  async writeReport(config: KbConfig, log: LogEntry[]): Promise<ReportResult> {
    const { buffer, fileName } = await buildReportWorkbook(config, log);
    const sp = await loadSharePointClient();
    // Upload the .xlsx into the same folder we scanned. The SharePoint connector's
    // CreateFile action accepts dataset (site), folderPath, name, and file bytes.
    await sp.CreateFile({
      dataset: config.siteUrl,
      folderPath: config.folderPath || '/',
      name: fileName,
      body: buffer,
    });
    const location = `${config.folderPath || '/'}/${fileName}`.replace(/\/+/g, '/');
    return { fileName, location };
  }

  async emailReport(to: string[], subject: string, html: string, attachment: { fileName: string; buffer: ArrayBuffer }): Promise<void> {
    // TODO(real connector): Call the Outlook connector's SendEmailV2 action with
    //   - To = comma-joined recipient list
    //   - Subject = subject
    //   - Body = html
    //   - Attachments = [{ Name: attachment.fileName, ContentBytes: base64(buffer) }]
    // so the generated .xlsx run report is mailed directly from the maker's
    // mailbox. Until that action is wired, log the payload for local validation.
    console.log('TODO SendEmailV2', { to, subject, html, attachment: { fileName: attachment.fileName, bytes: attachment.buffer.byteLength } });
  }

  async suggestEdits(article: ProcessedArticle): Promise<ArticleSuggestion> {
    // Preferred wiring is one of:
    //   1) An Azure OpenAI custom connector — call its "Chat completions" action
    //      with a system prompt like "You are a KB editor. Improve clarity,
    //      structure, and tone. Return JSON: { html, title?, summary, changes[] }."
    //   2) A Dataverse AI Prompt action (Power Platform "Prompts" feature) bound
    //      to an existing prompt template, executed via the Dataverse connector's
    //      ExecuteAction.
    // For now, fall back to the same deterministic heuristics the mock uses so
    // the button is always functional. Swap this body for the real connector
    // call when the connection is provisioned.
    return buildMockSuggestion(article);
  }

  async findOverlaps(articles: ProcessedArticle[], environment?: PowerPlatformEnvironment): Promise<Record<string, OverlapMatch[]>> {
    // Fetch a candidate pool of published / draft knowledgearticles from
    // Dataverse and score them client-side. For very large KBs you'd push
    // server-side relevance search instead (Dataverse `relevance search` or
    // a Cognitive Search index), but for typical KBs a few hundred rows is fine.
    const dv = await loadDataverseClient(environment);
    const res = await dv.knowledgearticle.list?.({
      $select: 'knowledgearticleid,title,description,modifiedon',
      $top: 5000,
      $filter: "statecode eq 0 or statecode eq 3", // draft or published
    });
    const candidates: ExistingKbArticle[] = (res?.value ?? []).map((r: any) => ({
      id: r.knowledgearticleid,
      title: r.title ?? '',
      excerpt: (r.description ?? '').toString().slice(0, 600),
      modifiedOn: r.modifiedon,
    }));
    return scoreOverlaps(articles, candidates);
  }
}

// ---- Generated-client loaders (replace after `pac code add-data-source`) ----

async function loadSharePointClient(): Promise<any> {
  // SharePoint Online connection isn't provisioned on this environment yet.
  // To make the app fully demo-able end-to-end, we return a small in-memory
  // "demo" SharePoint provider so the Browse-Site / Browse-Folder dialogs work
  // and surface a clear banner that this is sample data. The real Dataverse
  // write path is still live — load via the "Upload local files" drop zone
  // for actual KB ingestion.
  //
  // To switch to a real SharePoint connection:
  //   1. Create a SharePoint connection in make.powerapps.com (Connections > New)
  //   2. npx power-apps add-data-source -a shared_sharepointonline -c <SP-CONN-ID> --non-interactive
  //   3. Replace this stub with an adapter over the generated SharePointOnlineService
  return DEMO_SHAREPOINT_CLIENT;
}

const DEMO_SHAREPOINT_CLIENT = {
  __isDemo: true,
  async GetAllSites() {
    return {
      value: [
        { Id: 'demo-site-1', DisplayName: 'Customer Service KB (Demo)', Url: 'https://demo.sharepoint.com/sites/csv-kb', Description: 'Demo SharePoint site — KB source documents' },
        { Id: 'demo-site-2', DisplayName: 'Product Documentation (Demo)', Url: 'https://demo.sharepoint.com/sites/product-docs', Description: 'Demo SharePoint site — product manuals & FAQs' },
      ],
    };
  },
  async GetFolderItemsByPath() {
    return {
      value: [
        { Name: 'Articles', FullPath: '/Shared Documents/Articles', IsFolder: true, HasChildren: true },
        { Name: 'FAQs', FullPath: '/Shared Documents/FAQs', IsFolder: true, HasChildren: false },
        { Name: 'Release Notes', FullPath: '/Shared Documents/Release Notes', IsFolder: true, HasChildren: false },
      ],
    };
  },
  async GetFolderFilesByPath() {
    return {
      value: [
        { Identifier: 'demo-file-1', FilenameWithExtension: 'Getting Started.docx', FullPath: '/Shared Documents/Articles/Getting Started.docx', Size: 24576, TimeLastModified: new Date().toISOString(), IsFolder: false },
        { Identifier: 'demo-file-2', FilenameWithExtension: 'Troubleshooting.docx', FullPath: '/Shared Documents/Articles/Troubleshooting.docx', Size: 30720, TimeLastModified: new Date().toISOString(), IsFolder: false },
        { Identifier: 'demo-file-3', FilenameWithExtension: 'FAQ.pdf', FullPath: '/Shared Documents/Articles/FAQ.pdf', Size: 102400, TimeLastModified: new Date().toISOString(), IsFolder: false },
      ],
    };
  },
  async GetFileContent() {
    throw new Error(
      'Demo SharePoint provider — file download isn\'t implemented. ' +
      'Use the "Upload local files" drop zone in the Configure panel to ingest real content into Dataverse.'
    );
  },
  async CreateFile() {
    // No-op for demo — pretend the report was uploaded successfully.
    return { Id: 'demo-uploaded', Name: 'report.xlsx' };
  },
};

async function loadDataverseClient(_environment?: PowerPlatformEnvironment): Promise<any> {
  // Lazy-import the generated services so build-time tree shaking is preserved
  // and so this module still loads in pure-mock test contexts.
  const generated = await import('../generated');
  const KA = generated.KnowledgearticlesService;
  const SUB = generated.SubjectsService;
  const LL = generated.LanguagelocaleService;

  function unwrap<T>(r: any): T {
    if (r && r.error) {
      const msg = r.error?.message ?? r.error?.code ?? 'Dataverse error';
      throw new Error(msg);
    }
    return (r?.data ?? r) as T;
  }

  function odataToOpts(opts: any): any {
    if (!opts) return {};
    const out: any = {};
    if (opts.$select) out.select = String(opts.$select).split(',').map((s: string) => s.trim()).filter(Boolean);
    if (opts.$filter) out.filter = opts.$filter;
    if (opts.$top) out.top = Number(opts.$top);
    if (opts.$orderby) out.orderBy = [String(opts.$orderby)];
    return out;
  }

  return {
    knowledgearticle: {
      async create(body: any) {
        // Map "<field>@odata.bind" -> "<Navigation>@odata.bind" expected by
        // the generated client. The generated model exposes "Subject@odata.bind"
        // and "Language@odata.bind"; the existing call sites use lowercase
        // logical-name forms which the underlying connector also accepts in
        // most cases. Pass through as-is and let the SDK normalize.
        const data = unwrap<any>(await KA.create(body));
        return data ?? {};
      },
      async update(id: string, body: any) {
        return unwrap<any>(await KA.update(id, body));
      },
      async list(opts: any) {
        const data = unwrap<any[]>(await KA.getAll(odataToOpts(opts)));
        return { value: data ?? [] };
      },
    },
    subject: {
      async list(opts: any) {
        const data = unwrap<any[]>(await SUB.getAll(odataToOpts(opts)));
        return { value: data ?? [] };
      },
    },
    language: {
      async list(opts: any) {
        const data = unwrap<any[]>(await LL.getAll(odataToOpts(opts)));
        return { value: data ?? [] };
      },
    },
    // Best-effort: the generated MicrosoftDataverseService doesn't expose
    // WhoAmI / RetrieveProvisionedLanguages / discovery / raw fetch. Callers
    // already optional-chain these and degrade gracefully.
    whoAmI: undefined,
    systemuser: undefined,
    org: undefined,
    discovery: undefined,
    fetch: undefined,
  };
}

async function loadEnvBaseUrl(environment?: PowerPlatformEnvironment): Promise<string | undefined> {
  if (environment?.url) {
    return environment.url;
  }
  // The published Code App knows its Dataverse host from the connection
  // metadata Power Platform injects at runtime. Replace this with whatever
  // value the generated client exposes (e.g. dv.config.environmentUrl).
  return undefined;
}

function arrayBufferToBase64(bytes: ArrayBuffer): string {
  const array = new Uint8Array(bytes);
  let binary = '';
  for (const byte of array) {
    binary += String.fromCharCode(byte);
  }
  return btoa(binary);
}

function readProfilesFallback(): SavedScanProfile[] {
  if (typeof window === 'undefined') {
    return [];
  }
  try {
    const raw = window.localStorage.getItem(PROFILES_KEY);
    if (!raw) {
      return [];
    }
    const parsed = JSON.parse(raw) as SavedScanProfile[];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function writeProfilesFallback(profiles: SavedScanProfile[]) {
  if (typeof window === 'undefined') {
    return;
  }
  window.localStorage.setItem(PROFILES_KEY, JSON.stringify(profiles));
}

const PROFILES_KEY = 'kbloader.profiles';

// Minimal LCID lookup for the languages most D365 KB orgs use. Extend as needed.
const LCID_TO_CODE: Record<number, string> = {
  1033: 'en-US', 2057: 'en-GB', 1031: 'de-DE', 1036: 'fr-FR', 3082: 'es-ES',
  1041: 'ja-JP', 1040: 'it-IT', 1043: 'nl-NL', 1046: 'pt-BR', 2052: 'zh-CN',
};
const LCID_NAME: Record<number, string> = {
  1033: 'English (United States)', 2057: 'English (United Kingdom)',
  1031: 'German (Germany)', 1036: 'French (France)', 3082: 'Spanish (Spain)',
  1041: 'Japanese', 1040: 'Italian (Italy)', 1043: 'Dutch (Netherlands)',
  1046: 'Portuguese (Brazil)', 2052: 'Chinese (Simplified)',
};
