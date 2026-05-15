import type { KbLoaderService } from './KbLoaderService';
import type { SourceFile, ProcessedArticle, KbConfig, LogEntry } from '../types';

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

const log: LogEntry[] = [];

export class MockKbLoaderService implements KbLoaderService {
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

  async appendLog(_config: KbConfig, entry: LogEntry): Promise<void> {
    log.unshift({ ...entry, id: String(log.length + 1) });
  }

  async getLog(_config: KbConfig): Promise<LogEntry[]> {
    return [...log];
  }
}

function delay(ms: number) { return new Promise(res => setTimeout(res, ms)); }
