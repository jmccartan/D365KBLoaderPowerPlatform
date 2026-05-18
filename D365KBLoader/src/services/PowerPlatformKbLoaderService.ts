import type { KbLoaderService } from './KbLoaderService';
import type { SourceFile, ProcessedArticle, KbConfig, LogEntry, SharePointSite, FolderItem, ReportResult, ArticleSuggestion, ExistingKbArticle, OverlapMatch, PowerPlatformEnvironment, KbLanguage, KbSubject, KbUser, SavedScanProfile } from '../types';
import { classify } from '../processing/pipeline';
import { buildReportWorkbook } from '../reporting/report';
import { buildMockSuggestion } from './copilotSuggest';
import { scoreOverlaps } from './overlapDetect';

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
    // GET /api/data/v9.2/RetrieveProvisionedLanguages returns LCIDs enabled for the env.
    const dv = await loadDataverseClient();
    const res = await dv.org?.retrieveProvisionedLanguages?.();
    if (!res?.LocaleIds?.length) throw new Error('Could not retrieve provisioned languages from Dataverse.');
    return res.LocaleIds.map((lcid: number) => ({
      id: String(lcid),
      code: LCID_TO_CODE[lcid] ?? String(lcid),
      displayName: LCID_NAME[lcid] ?? `Locale ${lcid}`,
    }));
  }

  async listSubjects(parentId?: string): Promise<KbSubject[]> {
    const dv = await loadDataverseClient();
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

  async findArticleByTitle(title: string): Promise<ExistingKbArticle | undefined> {
    const dv = await loadDataverseClient();
    const normalized = title.trim().replace(/\s+/g, ' ');
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
    return { id: r.knowledgearticleid, title: r.title, modifiedOn: r.modifiedon };
  }

  async updateKnowledgeArticle(existingId: string, article: ProcessedArticle): Promise<void> {
    const dv = await loadDataverseClient();
    await dv.knowledgearticle.update(existingId, {
      title: article.title,
      content: article.html,
      description: `Updated from ${article.source.path}`,
      ...(article.languageId ? { 'languagelocaleid@odata.bind': `/languagelocale(${article.languageId})` } : {}),
      ...(article.subjectId ? { 'subjectid@odata.bind': `/subjects(${article.subjectId})` } : {}),
    });
  }

  /**
   * The Code App is bound to a single Dataverse environment at build time
   * (set by `pac code init` and the connection references in the published
   * app). The Global Discovery Service lets the same signed-in user enumerate
   * all environments they have access to so the UI can switch between them.
   *
   * Endpoint: `https://globaldisco.crm.dynamics.com/api/discovery/v2.0/Instances`
   * Auth: same AAD bearer token used by the Dataverse connector.
   */
  async listEnvironments(): Promise<PowerPlatformEnvironment[]> {
    const dv = await loadDataverseClient();
    const res = await dv.discovery?.listInstances?.();
    if (!res?.value) {
      throw new Error(
        'Could not enumerate environments. The default Dataverse connection must include the ' +
        'Global Discovery endpoint, or wire this to the BAP API directly.'
      );
    }
    return res.value.map((i: any) => ({
      id: i.EnvironmentId ?? i.Id ?? i.UrlName,
      displayName: i.FriendlyName ?? i.Name,
      url: i.Url,
      region: i.Region,
      isDefault: !!i.IsDefault,
      knowledgebaseStatus: 'unknown' as const,
    } as PowerPlatformEnvironment));
  }

  /**
   * Check whether the `knowledgearticle` table exists in the target environment.
   * A 200 response from EntityDefinitions means the Customer Service / Knowledge
   * Management solution is installed; 404 means it's missing.
   */
  async checkKnowledgebase(env: PowerPlatformEnvironment): Promise<PowerPlatformEnvironment> {
    try {
      const url = `${env.url.replace(/\/$/, '')}/api/data/v9.2/EntityDefinitions(LogicalName='knowledgearticle')?$select=LogicalName`;
      const dv = await loadDataverseClient();
      const res = await dv.fetch?.(url);
      if (res?.status === 404) return { ...env, knowledgebaseStatus: 'missing' };
      if (res?.ok === false) {
        return { ...env, knowledgebaseStatus: 'error', knowledgebaseError: `HTTP ${res.status}` };
      }
      return { ...env, knowledgebaseStatus: 'present' };
    } catch (e: any) {
      return { ...env, knowledgebaseStatus: 'error', knowledgebaseError: String(e?.message ?? e) };
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

  async createKnowledgeArticle(article: ProcessedArticle, config?: KbConfig): Promise<{ id: string; url?: string }> {
    const dv = await loadDataverseClient();
    const langId = article.languageId ?? config?.defaultLanguageId;
    const subjId = article.subjectId ?? config?.defaultSubjectId;
    const publish = article.publish ?? config?.publishOnLoad ?? false;

    const created = await dv.knowledgearticle.create({
      title: article.title,
      content: article.html,
      description: `Imported from ${article.source.path}`,
      // 3 = Published (statecode 3, statuscode 7); leaving unset = Draft.
      ...(publish ? { statecode: 3, statuscode: 7 } : {}),
      ...(langId ? { 'languagelocaleid@odata.bind': `/languagelocale(${langId})` } : {}),
      ...(subjId ? { 'subjectid@odata.bind': `/subjects(${subjId})` } : {}),
    });
    const id = created.knowledgearticleid ?? created.id;
    const envBase = await loadEnvBaseUrl();
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

  async findOverlaps(articles: ProcessedArticle[]): Promise<Record<string, OverlapMatch[]>> {
    // Fetch a candidate pool of published / draft knowledgearticles from
    // Dataverse and score them client-side. For very large KBs you'd push
    // server-side relevance search instead (Dataverse `relevance search` or
    // a Cognitive Search index), but for typical KBs a few hundred rows is fine.
    const dv = await loadDataverseClient();
    const res = await dv.knowledgearticle.list?.({
      $select: 'knowledgearticleid,title,description,modifiedon',
      $top: 500,
      $filter: "statuscode eq 3 or statecode eq 0", // published or draft
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
  // Example after generation:
  //   const mod = await import('../Models/SharePointOnlineService');
  //   return new mod.SharePointOnlineService();
  throw new Error(
    'SharePoint client not wired up. Run `pac code add-data-source -a shared_sharepointonline` ' +
    'and update PowerPlatformKbLoaderService.ts with the generated import.'
  );
}

async function loadDataverseClient(): Promise<any> {
  // Example after generation:
  //   const mod = await import('../Models/knowledgearticleService');
  //   return { knowledgearticle: mod.knowledgearticleService };
  throw new Error(
    'Dataverse client not wired up. Run `pac code add-data-source -a shared_commondataserviceforapps -t knowledgearticle` ' +
    'and update PowerPlatformKbLoaderService.ts with the generated import.'
  );
}

async function loadEnvBaseUrl(): Promise<string | undefined> {
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
