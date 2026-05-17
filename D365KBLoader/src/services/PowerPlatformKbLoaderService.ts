import type { KbLoaderService } from './KbLoaderService';
import type { SourceFile, ProcessedArticle, KbConfig, LogEntry, SharePointSite, FolderItem, ReportResult } from '../types';
import { classify } from '../processing/pipeline';
import { buildReportWorkbook } from '../reporting/report';

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

  async createKnowledgeArticle(article: ProcessedArticle): Promise<string> {
    const dv = await loadDataverseClient();
    const created = await dv.knowledgearticle.create({
      title: article.title,
      content: article.html,
      description: `Imported from ${article.source.path}`,
      languagelocaleid_value: undefined  // optionally set the language lookup
    });
    return created.knowledgearticleid ?? created.id;
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
