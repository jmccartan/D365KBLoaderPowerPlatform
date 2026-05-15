import { useEffect, useMemo, useState } from 'react';
import {
  makeStyles, tokens, Text, Tab, TabList, Spinner, MessageBar, MessageBarBody, MessageBarTitle
} from '@fluentui/react-components';
import { ConfigPanel } from './components/ConfigPanel';
import { ReviewPanel } from './components/ReviewPanel';
import { ProgressPanel } from './components/ProgressPanel';
import { processFile, classify } from './processing/pipeline';
import { getService } from './services';
import type { KbConfig, ProcessedArticle, LogEntry } from './types';

const useStyles = makeStyles({
  shell: { padding: tokens.spacingHorizontalXXL, display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalL, fontFamily: tokens.fontFamilyBase, minHeight: '100vh', backgroundColor: tokens.colorNeutralBackground2 },
  header: { display: 'flex', alignItems: 'baseline', justifyContent: 'space-between', gap: tokens.spacingHorizontalM },
  title: { fontSize: tokens.fontSizeHero700, fontWeight: tokens.fontWeightSemibold },
  sub: { color: tokens.colorNeutralForeground3 }
});

type Stage = 'config' | 'review' | 'progress';

export function App() {
  const s = useStyles();
  const svc = useMemo(() => getService(), []);

  const [config, setConfig] = useState<KbConfig>({
    siteUrl: '',
    folderPath: '',
    logListName: 'KB Loader Log'
  });

  const [stage, setStage] = useState<Stage>('config');
  const [scanning, setScanning] = useState(false);
  const [scanError, setScanError] = useState<string | undefined>();
  const [articles, setArticles] = useState<ProcessedArticle[]>([]);

  const [loading, setLoading] = useState(false);
  const [done, setDone] = useState(0);
  const [errors, setErrors] = useState(0);
  const [log, setLog] = useState<LogEntry[]>([]);

  async function refreshLog() {
    try { setLog(await svc.getLog(config)); } catch { /* ignore */ }
  }

  useEffect(() => { refreshLog(); /* on mount */ }, []);

  async function handleScan() {
    setScanning(true);
    setScanError(undefined);
    try {
      const files = await svc.listFiles(config);
      const supported = files.filter(f => f.kind !== 'unsupported');
      const skipped = files.filter(f => f.kind === 'unsupported');

      const processed: ProcessedArticle[] = [];
      for (const f of supported) {
        try {
          const buf = await svc.downloadFile(f);
          const a = await processFile(f, buf);
          processed.push(a);
          await svc.appendLog(config, mkLog(f.name, 'process', 'success', 'Converted to HTML', f.path));
        } catch (e: any) {
          processed.push({
            id: f.id, source: f, title: f.name, html: '', rawHtml: '',
            warnings: [String(e?.message ?? e)], selected: false, loadStatus: 'error', loadError: String(e?.message ?? e)
          });
          await svc.appendLog(config, mkLog(f.name, 'process', 'error', String(e?.message ?? e), f.path));
        }
      }
      for (const f of skipped) {
        await svc.appendLog(config, mkLog(f.name, 'skip', 'info', `Unsupported type`, f.path));
      }
      setArticles(processed);
      setStage('review');
      await refreshLog();
    } catch (e: any) {
      setScanError(String(e?.message ?? e));
    } finally {
      setScanning(false);
    }
  }

  async function handleLoad() {
    setLoading(true);
    setStage('progress');
    setDone(0); setErrors(0);
    const targets = articles.filter(a => a.selected && a.loadStatus !== 'success');
    let d = 0, err = 0;

    for (const a of targets) {
      setArticles(prev => prev.map(p => p.id === a.id ? { ...p, loadStatus: 'loading', loadError: undefined } : p));
      try {
        const id = await svc.createKnowledgeArticle(a);
        setArticles(prev => prev.map(p => p.id === a.id ? { ...p, loadStatus: 'success', knowledgeArticleId: id } : p));
        await svc.appendLog(config, mkLog(a.source.name, 'load', 'success', `Created knowledgearticle`, a.source.path, id));
        d++;
      } catch (e: any) {
        const msg = String(e?.message ?? e);
        setArticles(prev => prev.map(p => p.id === a.id ? { ...p, loadStatus: 'error', loadError: msg } : p));
        await svc.appendLog(config, mkLog(a.source.name, 'load', 'error', msg, a.source.path));
        err++;
      }
      setDone(d + err);
      setErrors(err);
    }
    await refreshLog();
    setLoading(false);
  }

  return (
    <div className={s.shell}>
      <div className={s.header}>
        <div>
          <div className={s.title}>D365 KB Loader</div>
          <Text className={s.sub}>SharePoint folder → Dynamics 365 Knowledgebase, with preview, edit, and audit log.</Text>
        </div>
        <TabList selectedValue={stage} onTabSelect={(_, d) => setStage(d.value as Stage)}>
          <Tab value="config">1 · Configure</Tab>
          <Tab value="review" disabled={articles.length === 0}>2 · Review ({articles.length})</Tab>
          <Tab value="progress">3 · Progress &amp; log</Tab>
        </TabList>
      </div>

      {import.meta.env?.VITE_USE_REAL_CONNECTORS !== 'true' && (
        <MessageBar intent="info">
          <MessageBarBody>
            <MessageBarTitle>Mock mode</MessageBarTitle>
            Running with sample data. Run <code>pac code add-data-source</code> for SharePoint and Dataverse,
            then set <code>VITE_USE_REAL_CONNECTORS=true</code> in <code>.env.local</code> to use live connectors.
          </MessageBarBody>
        </MessageBar>
      )}

      {stage === 'config' && (
        <ConfigPanel config={config} onChange={setConfig} onScan={handleScan} scanning={scanning} error={scanError} />
      )}
      {stage === 'review' && (
        <ReviewPanel articles={articles} onChange={setArticles} onLoad={handleLoad} loading={loading} />
      )}
      {stage === 'progress' && (
        <ProgressPanel
          done={done}
          total={articles.filter(a => a.selected).length}
          errors={errors}
          log={log}
          onRefresh={refreshLog}
        />
      )}
    </div>
  );
}

function mkLog(fileName: string, action: LogEntry['action'], status: LogEntry['status'], message: string, sourcePath?: string, knowledgeArticleId?: string): LogEntry {
  return { timestamp: new Date().toISOString(), fileName, action, status, message, sourcePath, knowledgeArticleId };
}
