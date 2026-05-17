import type { SourceFile, ProcessedArticle, KbConfig, LogEntry, SharePointSite, FolderItem, ReportResult, ArticleSuggestion, OverlapMatch, PowerPlatformEnvironment } from '../types';

/**
 * Service contract. The MockService implementation lets the UI run locally
 * without connectors. The PowerPlatformService implementation uses the
 * generated Power Platform Code App SDK clients (created by `pac code
 * add-data-source`) to call SharePoint and Dataverse via the connectors.
 *
 * Wire-up steps for the real service are documented in README.md.
 */
export interface KbLoaderService {
  /** List Power Platform / Dataverse environments the signed-in user can access. */
  listEnvironments(): Promise<PowerPlatformEnvironment[]>;
  /** Verify the given environment has the knowledgearticle table available. */
  checkKnowledgebase(env: PowerPlatformEnvironment): Promise<PowerPlatformEnvironment>;
  /** Discover SharePoint sites the user has access to (for the Browse dialog). */
  listSites(): Promise<SharePointSite[]>;
  /** List sub-folders of the given path within a site (folderPath '/' = root). */
  listFolders(siteUrl: string, folderPath: string): Promise<FolderItem[]>;
  listFiles(config: KbConfig): Promise<SourceFile[]>;
  downloadFile(file: SourceFile): Promise<ArrayBuffer>;
  createKnowledgeArticle(article: ProcessedArticle): Promise<string>;
  /** Persist a formatted Excel report of the run to the source folder (or download in mock mode). */
  writeReport(config: KbConfig, log: LogEntry[]): Promise<ReportResult>;
  /** Generate AI-assisted edit suggestions for a single article (Copilot review). */
  suggestEdits(article: ProcessedArticle): Promise<ArticleSuggestion>;
  /**
   * Compare each candidate against the existing D365 knowledgebase and return
   * possible overlaps, keyed by the candidate article id.
   */
  findOverlaps(articles: ProcessedArticle[]): Promise<Record<string, OverlapMatch[]>>;
}
