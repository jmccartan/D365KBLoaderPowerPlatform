/**
 * Domain types for the KB loader.
 */
export type SourceFileKind = 'docx' | 'html' | 'pdf' | 'md' | 'unsupported';

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
  /** Direct deep-link to the article in D365 (set when known). */
  knowledgeArticleUrl?: string;
  loadStatus: 'pending' | 'loading' | 'success' | 'error' | 'skipped';
  loadError?: string;
  /** Possible overlap with existing D365 KB articles (set after a scan). */
  overlaps?: OverlapMatch[];
  /** Override the run-level defaults for this article. */
  languageId?: string;
  subjectId?: string;
  /** When true, publish on create instead of leaving as Draft. */
  publish?: boolean;
}

/** A languagelocaleid row used by the picker. */
export interface KbLanguage {
  id: string;
  code: string;           // e.g. "en-US"
  displayName: string;    // e.g. "English (United States)"
}

/** A subject (category) tree node. */
export interface KbSubject {
  id: string;
  name: string;
  /** Server-relative-style path for display, e.g. "/IT/Networking" */
  path: string;
  parentId?: string;
  hasChildren?: boolean;
}

/** A signed-in user — used for audit logging and owner assignment. */
export interface KbUser {
  id: string;
  displayName: string;
  email?: string;
}

/** What to do when a candidate article already exists in the KB. */
export type DuplicateAction = 'create-new' | 'skip' | 'update-existing';

export interface KbConfig {
  siteUrl: string;
  folderPath: string;
  /** Default language for newly loaded articles (per-article can override). */
  defaultLanguageId?: string;
  /** Default subject/category (per-article can override). */
  defaultSubjectId?: string;
  /** When true, articles are published on create; otherwise left as Draft. */
  publishOnLoad?: boolean;
  /** What to do when a candidate already has a matching article in the KB. */
  duplicateAction?: DuplicateAction;
  /** Recurse into sub-folders during scan. */
  recursive?: boolean;
  /** Only include files modified on/after this ISO timestamp. */
  modifiedSince?: string;
  /** Skip files already loaded successfully in a prior run (read from prior report). */
  incremental?: boolean;
}

/** A Copilot-generated suggestion for editing a KB article. */
export interface ArticleSuggestion {
  /** Proposed sanitized HTML body. */
  html: string;
  /** Optional revised title — undefined means keep the existing one. */
  title?: string;
  /** Short, human-readable summary of what changed and why. */
  summary: string;
  /** Bullet list of specific changes the suggestion makes. */
  changes: string[];
}

/** An existing knowledgearticle row used for overlap detection. */
export interface ExistingKbArticle {
  id: string;
  title: string;
  /** Short plain-text excerpt for the overlap UI. */
  excerpt?: string;
  /** Optional deep-link into D365. */
  url?: string;
  modifiedOn?: string;
}

/** One match between a candidate import and an existing KB article. */
export interface OverlapMatch {
  article: ExistingKbArticle;
  /** Similarity score between 0 and 1. */
  score: number;
  /** Human-readable reasons (e.g., title token overlap, shared keywords). */
  reasons: string[];
}

/** A Power Platform / Dataverse environment the user can target. */
export interface PowerPlatformEnvironment {
  id: string;
  displayName: string;
  /** Dataverse Web API base URL, e.g. https://contoso.crm.dynamics.com */
  url: string;
  region?: string;
  /** True if this environment is the user's default. */
  isDefault?: boolean;
  /** Set by checkKnowledgebase. */
  knowledgebaseStatus: 'unknown' | 'checking' | 'present' | 'missing' | 'error';
  knowledgebaseError?: string;
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
  action: 'scan' | 'process' | 'load' | 'skip' | 'update';
  status: 'success' | 'error' | 'info';
  message: string;
  knowledgeArticleId?: string;
  sourcePath?: string;
  /** Identity of the user who triggered the action. */
  userDisplayName?: string;
  userEmail?: string;
}
