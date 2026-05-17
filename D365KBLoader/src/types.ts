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
}

/** Result of writing the per-run Excel report. */
export interface ReportResult {
  fileName: string;
  /** Where the file ended up: server-relative path for SharePoint, or 'download' for browser download. */
  location: string;
  /** True if the file was streamed to the user's browser instead of uploaded. */
  downloaded?: boolean;
}

/** A SharePoint site the user can pick from. */
export interface SharePointSite {
  id: string;
  name: string;
  url: string;
  description?: string;
}

/** A folder inside a SharePoint document library, used by the folder browser. */
export interface FolderItem {
  name: string;
  /** Server-relative path, e.g. /Shared Documents/KB */
  path: string;
  /** Optional hint: true if known to have no sub-folders (lets the UI hide the chevron). */
  hasChildren?: boolean;
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
