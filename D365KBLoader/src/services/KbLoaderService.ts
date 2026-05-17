import type { SourceFile, ProcessedArticle, KbConfig, LogEntry, SharePointSite, FolderItem, ReportResult, ArticleSuggestion, OverlapMatch, PowerPlatformEnvironment, KbLanguage, KbSubject, KbUser, ExistingKbArticle } from '../types';

/**
 * Service contract. The MockService implementation lets the UI run locally
 * without connectors. The PowerPlatformService implementation uses the
 * generated Power Platform Code App SDK clients (created by `pac code
 * add-data-source`) to call SharePoint and Dataverse via the connectors.
 *
 * Wire-up steps for the real service are documented in README.md.
 */
export interface KbLoaderService {
  /** Identity of the signed-in user (for audit logging and ownership). */
  getCurrentUser(): Promise<KbUser>;
  /** List Power Platform / Dataverse environments the signed-in user can access. */
  listEnvironments(): Promise<PowerPlatformEnvironment[]>;
  /** Verify the given environment has the knowledgearticle table available. */
  checkKnowledgebase(env: PowerPlatformEnvironment): Promise<PowerPlatformEnvironment>;
  /** D365 KB metadata: languages and the subject (category) tree. */
  listLanguages(): Promise<KbLanguage[]>;
  /** Tree-style subject lookup. Pass parentId=undefined for root nodes. */
  listSubjects(parentId?: string): Promise<KbSubject[]>;
  /** Look up an existing article by exact title (used for idempotency). */
  findArticleByTitle(title: string): Promise<ExistingKbArticle | undefined>;
  /** Update an existing article (used when duplicateAction = 'update-existing'). */
  updateKnowledgeArticle(existingId: string, article: ProcessedArticle): Promise<void>;
  /** Discover SharePoint sites the user has access to (for the Browse dialog). */
  listSites(): Promise<SharePointSite[]>;
  /** List sub-folders of the given path within a site (folderPath '/' = root). */
  listFolders(siteUrl: string, folderPath: string): Promise<FolderItem[]>;
  listFiles(config: KbConfig): Promise<SourceFile[]>;
  downloadFile(file: SourceFile): Promise<ArrayBuffer>;
  createKnowledgeArticle(article: ProcessedArticle, config?: KbConfig): Promise<{ id: string; url?: string }>;
  /** Persist a formatted Excel report of the run to the source folder (or download in mock mode). */
  writeReport(config: KbConfig, log: LogEntry[]): Promise<ReportResult>;
  /** Read a prior run report (.xlsx) and return file paths that were loaded successfully — used for incremental mode. */
  readPriorReports?(config: KbConfig): Promise<Set<string>>;
  /** Send the report by email through the Outlook connector (real) or log (mock). */
  emailReport?(to: string[], subject: string, html: string, attachment: { fileName: string; buffer: ArrayBuffer }): Promise<void>;
  /** Generate AI-assisted edit suggestions for a single article (Copilot review). */
  suggestEdits(article: ProcessedArticle): Promise<ArticleSuggestion>;
  /**
   * Compare each candidate against the existing D365 knowledgebase and return
   * possible overlaps, keyed by the candidate article id.
   */
  findOverlaps(articles: ProcessedArticle[]): Promise<Record<string, OverlapMatch[]>>;
}
