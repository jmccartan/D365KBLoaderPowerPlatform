/**
 * Domain types for the KB loader.
 */
export type SourceFileKind = 'docx' | 'html' | 'unsupported';

export interface SourceFile {
  id: string;
  name: string;
  path: string;
  size: number;
  modified: string;
  kind: SourceFileKind;
}

export interface ProcessedArticle {
  id: string;
  source: SourceFile;
  title: string;
  /** Sanitized HTML body ready for KB. */
  html: string;
  /** Raw HTML before sanitization (for the Raw view). */
  rawHtml: string;
  warnings: string[];
  selected: boolean;
  /** Set after a successful load. */
  knowledgeArticleId?: string;
  loadStatus: 'pending' | 'loading' | 'success' | 'error';
  loadError?: string;
}

export interface KbConfig {
  siteUrl: string;
  folderPath: string;
  /** Optional SharePoint list to use as activity log. */
  logListName: string;
}

export interface LogEntry {
  id?: string;
  timestamp: string;
  fileName: string;
  action: 'scan' | 'process' | 'load' | 'skip';
  status: 'success' | 'error' | 'info';
  message: string;
  knowledgeArticleId?: string;
  sourcePath?: string;
}
