import type { KbLoaderService } from './KbLoaderService';
import type { SourceFile, ProcessedArticle, KbConfig, LogEntry, SharePointSite, FolderItem, ReportResult, ArticleSuggestion, ExistingKbArticle, OverlapMatch, PowerPlatformEnvironment, KbLanguage, KbSubject, KbUser } from '../types';
import { buildReportWorkbook, downloadBlob } from '../reporting/report';
import { buildMockSuggestion } from './copilotSuggest';
import { scoreOverlaps } from './overlapDetect';

const SAMPLE_DOCX_HTML = `
  <h1>Sample KB Article (mock)</h1>
  <p>This is a <strong>mock</strong> article generated locally because no
  SharePoint connector is wired up yet. Once you run
  <code>pac code add-data-source</code> for SharePoint, real files will appear here.</p>
  <p>Need help? Email helpdesk@contoso.com or use the pilot card 4111 1111 1111 1111 in the test tenant. The content is intentionally vivid so demos feel realistic and the PII scanner has believable findings.</p>
  <h2>Steps</h2>
  <ol><li>Pick site &amp; folder</li><li>Review</li><li>Load</li></ol>
`;

function makeMockFiles(folder: string = '/Shared Documents/KB'): SourceFile[] {
  const f = folder.endsWith('/') ? folder.slice(0, -1) : folder;
  const variant = f.split('/').pop() ?? '';
  const suffix = variant && variant !== 'KB' ? `-${variant.toLowerCase()}` : '';
  return [
    { id: `mock-1${suffix}`, name: `Reset-Password${suffix}.html`, path: `${f}/Reset-Password${suffix}.html`, size: 1024, modified: new Date().toISOString(), kind: 'html' },
    { id: `mock-2${suffix}`, name: `VPN-Setup${suffix}.docx`, path: `${f}/VPN-Setup${suffix}.docx`, size: 22048, modified: new Date().toISOString(), kind: 'docx' },
    { id: `mock-3${suffix}`, name: `Onboarding-Checklist${suffix}.docx`, path: `${f}/Onboarding-Checklist${suffix}.docx`, size: 8192, modified: new Date().toISOString(), kind: 'docx' },
    { id: `mock-4${suffix}`, name: `Printer-Troubleshooting${suffix}.htm`, path: `${f}/Printer-Troubleshooting${suffix}.htm`, size: 3050, modified: new Date().toISOString(), kind: 'html' },
  ];
}

const MOCK_SITES: SharePointSite[] = [
  { id: 's1', name: 'Support', url: 'https://contoso.sharepoint.com/sites/Support', description: 'Customer support knowledgebase' },
  { id: 's2', name: 'IT Help Desk', url: 'https://contoso.sharepoint.com/sites/ITHelp', description: 'Internal IT articles and runbooks' },
  { id: 's3', name: 'HR', url: 'https://contoso.sharepoint.com/sites/HR', description: 'Policies, benefits, onboarding' },
  { id: 's4', name: 'Field Services', url: 'https://contoso.sharepoint.com/sites/FieldServices', description: 'On-site technician procedures' },
];

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
  async getCurrentUser(): Promise<KbUser> {
    await delay(50);
    return { id: 'mock-user', displayName: 'Mock User', email: 'mock.user@contoso.com' };
  }

  async listEnvironments(): Promise<PowerPlatformEnvironment[]> {
    await delay(300);
    return MOCK_ENVIRONMENTS.map(environment => ({ ...environment }));
  }

  async checkKnowledgebase(env: PowerPlatformEnvironment): Promise<PowerPlatformEnvironment> {
    await delay(450);
    const canned = MOCK_ENVIRONMENTS.find(environment => environment.id === env.id);
    if (canned?.knowledgebaseStatus === 'error') {
      return { ...env, knowledgebaseStatus: 'error', knowledgebaseError: 'Could not reach environment (simulated).' };
    }
    return { ...env, knowledgebaseStatus: canned?.knowledgebaseStatus ?? 'unknown' };
  }

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

  async listFiles(config: KbConfig): Promise<SourceFile[]> {
    await delay(400);
    let files = makeMockFiles(config.folderPath || '/');
    if (config.recursive) {
      files = [
        ...files,
        ...makeMockFiles(`${config.folderPath || '/'}/Networking`.replace(/\/+/g, '/')),
        ...makeMockFiles(`${config.folderPath || '/'}/Accounts`.replace(/\/+/g, '/')),
      ];
    }
    if (config.modifiedSince) {
      const since = new Date(config.modifiedSince).getTime();
      files = files.filter(file => new Date(file.modified).getTime() >= since);
    }
    return files;
  }

  async readPriorReports(_config: KbConfig): Promise<Set<string>> {
    await delay(80);
    return new Set();
  }

  async downloadFile(file: SourceFile): Promise<ArrayBuffer> {
    await delay(150);
    if (file.kind === 'html') {
      return new TextEncoder().encode(`<html><body>${SAMPLE_DOCX_HTML}</body></html>`).buffer;
    }
    return new TextEncoder().encode(SAMPLE_DOCX_HTML).buffer;
  }

  async uploadImage(_name: string, bytes: ArrayBuffer, contentType: string): Promise<string> {
    await delay(60);
    return `data:${contentType};base64,${arrayBufferToBase64(bytes)}`;
  }

  async createKnowledgeArticle(article: ProcessedArticle): Promise<{ id: string; url?: string }> {
    await delay(500);
    if (article.title.toLowerCase().includes('fail')) throw new Error('Simulated failure');
    const id = `mock-ka-${Math.random().toString(36).slice(2, 10)}`;
    return { id, url: `https://contoso.crm.dynamics.com/main.aspx?etn=knowledgearticle&id=${id}` };
  }

  async updateKnowledgeArticle(_existingId: string, article: ProcessedArticle): Promise<void> {
    await delay(400);
    if (article.title.toLowerCase().includes('fail')) throw new Error('Simulated update failure');
  }

  async findArticleByTitle(title: string): Promise<ExistingKbArticle | undefined> {
    await delay(120);
    const normalizedTitle = title.trim().toLowerCase();
    return MOCK_EXISTING_KB.find(article => article.title.trim().toLowerCase() === normalizedTitle);
  }

  async listLanguages(): Promise<KbLanguage[]> {
    await delay(80);
    return [...MOCK_LANGUAGES];
  }

  async listSubjects(parentId?: string): Promise<KbSubject[]> {
    await delay(120);
    return MOCK_SUBJECTS.filter(subject => (subject.parentId ?? null) === (parentId ?? null));
  }

  async writeReport(config: KbConfig, log: LogEntry[]): Promise<ReportResult> {
    const { buffer, fileName } = await buildReportWorkbook(config, log);
    downloadBlob(buffer, fileName);
    return { fileName, location: 'browser download', downloaded: true };
  }

  async suggestEdits(article: ProcessedArticle): Promise<ArticleSuggestion> {
    await delay(800);
    return buildMockSuggestion(article);
  }

  async findOverlaps(articles: ProcessedArticle[]): Promise<Record<string, OverlapMatch[]>> {
    await delay(500);
    return scoreOverlaps(articles, MOCK_EXISTING_KB);
  }
}

const MOCK_ENVIRONMENTS: PowerPlatformEnvironment[] = [
  {
    id: 'env-prod',
    displayName: 'Contoso – Production',
    url: 'https://contoso.crm.dynamics.com',
    region: 'United States',
    isDefault: true,
    knowledgebaseStatus: 'present',
  },
  {
    id: 'env-test',
    displayName: 'Contoso – Test',
    url: 'https://contoso-test.crm.dynamics.com',
    region: 'United States',
    knowledgebaseStatus: 'present',
  },
  {
    id: 'env-dev',
    displayName: 'Contoso – Dev sandbox',
    url: 'https://contoso-dev.crm.dynamics.com',
    region: 'United States',
    knowledgebaseStatus: 'missing',
  },
  {
    id: 'env-eu',
    displayName: 'Contoso EU – Production',
    url: 'https://contoso-eu.crm4.dynamics.com',
    region: 'Europe',
    knowledgebaseStatus: 'error',
  },
];

const MOCK_LANGUAGES: KbLanguage[] = [
  { id: '1033', code: 'en-US', displayName: 'English (United States)' },
  { id: '2057', code: 'en-GB', displayName: 'English (United Kingdom)' },
  { id: '1031', code: 'de-DE', displayName: 'German (Germany)' },
  { id: '1036', code: 'fr-FR', displayName: 'French (France)' },
  { id: '3082', code: 'es-ES', displayName: 'Spanish (Spain)' },
  { id: '1041', code: 'ja-JP', displayName: 'Japanese' },
];

const MOCK_SUBJECTS: KbSubject[] = [
  { id: 'sub-it', name: 'IT', path: '/IT', hasChildren: true },
  { id: 'sub-it-net', name: 'Networking', path: '/IT/Networking', parentId: 'sub-it' },
  { id: 'sub-it-acct', name: 'Accounts & access', path: '/IT/Accounts', parentId: 'sub-it' },
  { id: 'sub-it-hw', name: 'Hardware', path: '/IT/Hardware', parentId: 'sub-it' },
  { id: 'sub-hr', name: 'HR', path: '/HR', hasChildren: true },
  { id: 'sub-hr-onb', name: 'Onboarding', path: '/HR/Onboarding', parentId: 'sub-hr' },
  { id: 'sub-hr-ben', name: 'Benefits', path: '/HR/Benefits', parentId: 'sub-hr' },
  { id: 'sub-support', name: 'Customer Support', path: '/Customer Support', hasChildren: false },
  { id: 'sub-field', name: 'Field Service', path: '/Field Service', hasChildren: false },
];

const MOCK_EXISTING_KB: ExistingKbArticle[] = [
  {
    id: 'ka-1001',
    title: 'How to reset your Windows password',
    excerpt: 'Step-by-step guide for resetting your domain account password via the self-service portal. Covers forgotten passwords, account lockouts, and MFA reset.',
    url: 'https://contoso.crm.dynamics.com/main.aspx?etn=knowledgearticle&id=ka-1001',
    modifiedOn: '2026-02-14',
  },
  {
    id: 'ka-1002',
    title: 'VPN setup for remote workers (Windows + macOS)',
    excerpt: 'Configure the corporate VPN client, connect, and troubleshoot common authentication errors. Includes split-tunnel and MFA setup.',
    url: 'https://contoso.crm.dynamics.com/main.aspx?etn=knowledgearticle&id=ka-1002',
    modifiedOn: '2026-04-02',
  },
  {
    id: 'ka-1003',
    title: 'Connect to corporate WiFi (CONTOSO-CORP)',
    excerpt: 'Join the CONTOSO-CORP wireless network with your work account. Covers certificate-based auth and BYOD enrollment.',
    url: 'https://contoso.crm.dynamics.com/main.aspx?etn=knowledgearticle&id=ka-1003',
    modifiedOn: '2025-11-20',
  },
  {
    id: 'ka-1004',
    title: 'Printer troubleshooting: paper jams and offline status',
    excerpt: 'Diagnose and clear paper jams, reset printer queues, and resolve offline / driver issues on Windows.',
    url: 'https://contoso.crm.dynamics.com/main.aspx?etn=knowledgearticle&id=ka-1004',
    modifiedOn: '2026-01-12',
  },
  {
    id: 'ka-1005',
    title: 'Onboarding checklist for new employees',
    excerpt: 'Day-one checklist: laptop, badge, account setup, mandatory training, and benefits enrollment.',
    url: 'https://contoso.crm.dynamics.com/main.aspx?etn=knowledgearticle&id=ka-1005',
    modifiedOn: '2026-03-30',
  },
  {
    id: 'ka-1006',
    title: 'Email outage runbook',
    excerpt: 'Tier-1 response steps when corporate email is unreachable. Includes Exchange health checks and customer communication template.',
    url: 'https://contoso.crm.dynamics.com/main.aspx?etn=knowledgearticle&id=ka-1006',
    modifiedOn: '2025-10-08',
  },
];

function delay(ms: number) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function normalizePath(pathValue: string): string {
  if (!pathValue || pathValue === '/' || pathValue.trim() === '') return '/';
  let normalized = pathValue.trim();
  if (!normalized.startsWith('/')) normalized = '/' + normalized;
  if (normalized.length > 1 && normalized.endsWith('/')) normalized = normalized.slice(0, -1);
  return normalized;
}

function arrayBufferToBase64(bytes: ArrayBuffer): string {
  const array = new Uint8Array(bytes);
  let binary = '';
  for (const byte of array) {
    binary += String.fromCharCode(byte);
  }
  return btoa(binary);
}
