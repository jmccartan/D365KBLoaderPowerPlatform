import type { SourceFile, ProcessedArticle, KbConfig, LogEntry } from '../types';

/**
 * Service contract. The MockService implementation lets the UI run locally
 * without connectors. The PowerPlatformService implementation uses the
 * generated Power Platform Code App SDK clients (created by `pac code
 * add-data-source`) to call SharePoint and Dataverse via the connectors.
 *
 * Wire-up steps for the real service are documented in README.md.
 */
export interface KbLoaderService {
  listFiles(config: KbConfig): Promise<SourceFile[]>;
  downloadFile(file: SourceFile): Promise<ArrayBuffer>;
  createKnowledgeArticle(article: ProcessedArticle): Promise<string>;
  appendLog(config: KbConfig, entry: LogEntry): Promise<void>;
  getLog(config: KbConfig, top?: number): Promise<LogEntry[]>;
}
