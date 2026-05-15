import type { KbLoaderService } from './KbLoaderService';
import type { SourceFile, ProcessedArticle, KbConfig, LogEntry } from '../types';
import { classify } from '../processing/pipeline';

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
 *     - GetFolderFilesByPath  -> list files in folder
 *     - GetFileContent         -> download bytes (FileIdentifier)
 *     - PostItem               -> add list item to log list
 *     - GetItems               -> read log list
 *   Dataverse:
 *     - CreateRecord on table 'knowledgearticle'
 *       Required fields: title, content (HTML), description (optional), articlepublicnumber (auto)
 */
export class PowerPlatformKbLoaderService implements KbLoaderService {
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

  async appendLog(config: KbConfig, entry: LogEntry): Promise<void> {
    const sp = await loadSharePointClient();
    await sp.PostItem({
      dataset: config.siteUrl,
      table: config.logListName,
      item: {
        Title: entry.fileName,
        Action: entry.action,
        Status: entry.status,
        Message: entry.message,
        SourcePath: entry.sourcePath ?? '',
        KnowledgeArticleId: entry.knowledgeArticleId ?? '',
        Timestamp: entry.timestamp
      }
    });
  }

  async getLog(config: KbConfig, top = 100): Promise<LogEntry[]> {
    const sp = await loadSharePointClient();
    const res = await sp.GetItems({
      dataset: config.siteUrl,
      table: config.logListName,
      $top: top,
      $orderby: 'Created desc'
    });
    return (res.value ?? []).map((it: any) => ({
      id: String(it.ID),
      timestamp: it.Timestamp ?? it.Created,
      fileName: it.Title,
      action: it.Action,
      status: it.Status,
      message: it.Message,
      sourcePath: it.SourcePath,
      knowledgeArticleId: it.KnowledgeArticleId
    }));
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
