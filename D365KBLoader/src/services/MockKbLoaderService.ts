import type { KbLoaderService } from './KbLoaderService';
import type { SourceFile, ProcessedArticle, KbConfig, LogEntry, SharePointSite, FolderItem, ReportResult } from '../types';
import { buildReportWorkbook, downloadBlob } from '../reporting/report';

const SAMPLE_DOCX_HTML = `
  <h1>Sample KB Article (mock)</h1>
  <p>This is a <strong>mock</strong> article generated locally because no
  SharePoint connector is wired up yet. Once you run
  <code>pac code add-data-source</code> for SharePoint, real files will appear here.</p>
  <h2>Steps</h2>
  <ol><li>Pick site &amp; folder</li><li>Review</li><li>Load</li></ol>
`;

function makeMockFiles(): SourceFile[] {
  return [
    { id: 'mock-1', name: 'Reset-Password.html', path: '/Shared Documents/KB/Reset-Password.html', size: 1024, modified: new Date().toISOString(), kind: 'html' },
    { id: 'mock-2', name: 'VPN-Setup.docx', path: '/Shared Documents/KB/VPN-Setup.docx', size: 22048, modified: new Date().toISOString(), kind: 'docx' },
    { id: 'mock-3', name: 'Onboarding-Checklist.docx', path: '/Shared Documents/KB/Onboarding-Checklist.docx', size: 8192, modified: new Date().toISOString(), kind: 'docx' },
    { id: 'mock-4', name: 'Printer-Troubleshooting.htm', path: '/Shared Documents/KB/Printer-Troubleshooting.htm', size: 3050, modified: new Date().toISOString(), kind: 'html' }
  ];
}

const MOCK_SITES: SharePointSite[] = [
  { id: 's1', name: 'Support', url: 'https://contoso.sharepoint.com/sites/Support', description: 'Customer support knowledgebase' },
  { id: 's2', name: 'IT Help Desk', url: 'https://contoso.sharepoint.com/sites/ITHelp', description: 'Internal IT articles and runbooks' },
  { id: 's3', name: 'HR', url: 'https://contoso.sharepoint.com/sites/HR', description: 'Policies, benefits, onboarding' },
  { id: 's4', name: 'Field Services', url: 'https://contoso.sharepoint.com/sites/FieldServices', description: 'On-site technician procedures' },
];

// A small fake folder tree per site, keyed by site URL.
const MOCK_TREE: Record<string, Record<string, string[]>> = {
  default: {
    '/': ['Shared Documents', 'Site Assets'],
    '/Shared Documents': ['KB', 'Archived', 'Drafts', 'Images'],
    '/Shared Documents/KB': ['Networking', 'Accounts', 'Hardware'],
    '/Shared Documents/KB/Networking': [],
    '/Shared Documents/KB/Accounts': [],
    '/Shared Documents/KB/Hardware': [],
    '/Shared Documents/Archived': ['2023', '2024'],
    '/Shared Documents/Archived/2023': [],
    '/Shared Documents/Archived/2024': [],
    '/Shared Documents/Drafts': [],
    '/Shared Documents/Images': [],
    '/Site Assets': [],
  },
};

export class MockKbLoaderService implements KbLoaderService {
  async listSites(): Promise<SharePointSite[]> {
    await delay(250);
    return [...MOCK_SITES];
  }

  async listFolders(_siteUrl: string, folderPath: string): Promise<FolderItem[]> {
    await delay(200);
    const tree = MOCK_TREE.default;
    const normalized = normalizePath(folderPath);
    const children = tree[normalized] ?? [];
    return children.map(name => {
      const path = normalized === '/' ? `/${name}` : `${normalized}/${name}`;
      return {
        name,
        path,
        hasChildren: (tree[path] ?? []).length > 0,
      };
    });
  }

  async listFiles(_config: KbConfig): Promise<SourceFile[]> {
    await delay(400);
    return makeMockFiles();
  }

  async downloadFile(file: SourceFile): Promise<ArrayBuffer> {
    await delay(150);
    if (file.kind === 'html') {
      return new TextEncoder().encode(`<html><body>${SAMPLE_DOCX_HTML}</body></html>`).buffer;
    }
    // Return text bytes; mammoth would normally read .docx zip — for the mock
    // we cheat and inject HTML directly via processFile path (handled by pipeline).
    return new TextEncoder().encode(SAMPLE_DOCX_HTML).buffer;
  }

  async createKnowledgeArticle(article: ProcessedArticle): Promise<string> {
    await delay(500);
    if (article.title.toLowerCase().includes('fail')) throw new Error('Simulated failure');
    return `mock-ka-${Math.random().toString(36).slice(2, 10)}`;
  }

  async writeReport(config: KbConfig, log: LogEntry[]): Promise<ReportResult> {
    const { buffer, fileName } = await buildReportWorkbook(config, log);
    // No SharePoint connector available locally — give it to the user as a download.
    downloadBlob(buffer, fileName);
    return { fileName, location: 'browser download', downloaded: true };
  }
}

function delay(ms: number) { return new Promise(res => setTimeout(res, ms)); }

function normalizePath(p: string): string {
  if (!p || p === '/' || p.trim() === '') return '/';
  let n = p.trim();
  if (!n.startsWith('/')) n = '/' + n;
  if (n.length > 1 && n.endsWith('/')) n = n.slice(0, -1);
  return n;
}
