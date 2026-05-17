import { useMemo, useState } from 'react';
import {
  makeStyles, tokens, Text, MessageBar, MessageBarBody, MessageBarTitle
} from '@fluentui/react-components';
import { BookDatabase24Filled, Sparkle20Filled } from '@fluentui/react-icons';
import { ConfigPanel } from './components/ConfigPanel';
import { ReviewPanel } from './components/ReviewPanel';
import { ProgressPanel } from './components/ProgressPanel';
import { Stepper } from './components/Stepper';
import { EnvironmentPicker } from './components/EnvironmentPicker';
import { processFile } from './processing/pipeline';
import { getService } from './services';
import { heroGradient } from './theme';
import type { KbConfig, ProcessedArticle, LogEntry, ReportResult, PowerPlatformEnvironment } from './types';

const useStyles = makeStyles({
  shell: {
    minHeight: '100vh',
    backgroundColor: tokens.colorNeutralBackground2,
    display: 'flex',
    flexDirection: 'column',
  },
  hero: {
    background: heroGradient,
    color: '#FFFFFF',
    padding: `${tokens.spacingVerticalXL} ${tokens.spacingHorizontalXXL}`,
    boxShadow: '0 2px 12px rgba(15, 60, 110, 0.25)',
    position: 'relative',
    overflow: 'hidden',
  },
  heroDecor: {
    position: 'absolute',
    right: '-80px',
    top: '-80px',
    width: '320px',
    height: '320px',
    borderRadius: '50%',
    background: 'radial-gradient(circle, rgba(255,255,255,0.18) 0%, rgba(255,255,255,0) 70%)',
    pointerEvents: 'none',
  },
  heroDecor2: {
    position: 'absolute',
    left: '40%',
    bottom: '-120px',
    width: '260px',
    height: '260px',
    borderRadius: '50%',
    background: 'radial-gradient(circle, rgba(255,255,255,0.10) 0%, rgba(255,255,255,0) 70%)',
    pointerEvents: 'none',
  },
  heroInner: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalL,
    position: 'relative',
  },
  logo: {
    width: '52px',
    height: '52px',
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: 'rgba(255,255,255,0.18)',
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    color: '#FFFFFF',
    backdropFilter: 'blur(6px)',
    boxShadow: 'inset 0 0 0 1px rgba(255,255,255,0.25)',
  },
  title: {
    fontSize: '28px',
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: '34px',
    letterSpacing: '-0.01em',
    color: '#FFFFFF',
  },
  sub: {
    color: 'rgba(255,255,255,0.85)',
    fontSize: tokens.fontSizeBase300,
  },
  modeChip: {
    marginLeft: 'auto',
    display: 'inline-flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusCircular,
    backgroundColor: 'rgba(255,255,255,0.18)',
    color: '#FFFFFF',
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    border: '1px solid rgba(255,255,255,0.3)',
  },
  stepperBar: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalXXL}`,
    boxShadow: '0 1px 2px rgba(15, 60, 110, 0.04)',
  },
  body: {
    padding: tokens.spacingHorizontalXXL,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalL,
    flex: 1,
  },
});

type Stage = 'config' | 'review' | 'progress';

export function App() {
  const s = useStyles();
  const svc = useMemo(() => getService(), []);

  const [config, setConfig] = useState<KbConfig>({
    siteUrl: '',
    folderPath: '',
  });

  const [stage, setStage] = useState<Stage>('config');
  const [scanning, setScanning] = useState(false);
  const [scanError, setScanError] = useState<string | undefined>();
  const [articles, setArticles] = useState<ProcessedArticle[]>([]);

  const [loading, setLoading] = useState(false);
  const [done, setDone] = useState(0);
  const [errors, setErrors] = useState(0);
  const [log, setLog] = useState<LogEntry[]>([]);
  const [reportSaving, setReportSaving] = useState(false);
  const [reportResult, setReportResult] = useState<ReportResult | undefined>();
  const [reportError, setReportError] = useState<string | undefined>();
  const [environment, setEnvironment] = useState<PowerPlatformEnvironment | undefined>();

  const isMock = import.meta.env?.VITE_USE_REAL_CONNECTORS !== 'true';
  const envReady = !!environment && environment.knowledgebaseStatus === 'present';

  function appendLog(entry: LogEntry) {
    setLog(prev => [{ ...entry, id: String(prev.length + 1) }, ...prev]);
  }

  async function handleSaveReport(currentLog?: LogEntry[]) {
    setReportSaving(true);
    setReportError(undefined);
    try {
      const result = await svc.writeReport(config, currentLog ?? log);
      setReportResult(result);
    } catch (e: any) {
      setReportError(String(e?.message ?? e));
    } finally {
      setReportSaving(false);
    }
  }

  async function handleScan() {
    setScanning(true);
    setScanError(undefined);
    setLog([]);
    setReportResult(undefined);
    setReportError(undefined);
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
          appendLog(mkLog(f.name, 'process', 'success', 'Converted to HTML', f.path));
        } catch (e: any) {
          processed.push({
            id: f.id, source: f, title: f.name, html: '', rawHtml: '',
            warnings: [String(e?.message ?? e)], selected: false, loadStatus: 'error', loadError: String(e?.message ?? e)
          });
          appendLog(mkLog(f.name, 'process', 'error', String(e?.message ?? e), f.path));
        }
      }
      for (const f of skipped) {
        appendLog(mkLog(f.name, 'skip', 'info', `Unsupported type`, f.path));
      }
      setArticles(processed);
      setStage('review');
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
    setReportResult(undefined);
    setReportError(undefined);
    const targets = articles.filter(a => a.selected && a.loadStatus !== 'success');
    let d = 0, err = 0;
    const newEntries: LogEntry[] = [];

    for (const a of targets) {
      setArticles(prev => prev.map(p => p.id === a.id ? { ...p, loadStatus: 'loading', loadError: undefined } : p));
      try {
        const id = await svc.createKnowledgeArticle(a);
        setArticles(prev => prev.map(p => p.id === a.id ? { ...p, loadStatus: 'success', knowledgeArticleId: id } : p));
        const entry = mkLog(a.source.name, 'load', 'success', `Created knowledgearticle`, a.source.path, id);
        newEntries.push(entry);
        appendLog(entry);
        d++;
      } catch (e: any) {
        const msg = String(e?.message ?? e);
        setArticles(prev => prev.map(p => p.id === a.id ? { ...p, loadStatus: 'error', loadError: msg } : p));
        const entry = mkLog(a.source.name, 'load', 'error', msg, a.source.path);
        newEntries.push(entry);
        appendLog(entry);
        err++;
      }
      setDone(d + err);
      setErrors(err);
    }
    setLoading(false);

    // Auto-save the run report. Combine prior log entries (scan/process) with this run's load entries.
    const combined = [...newEntries.slice().reverse(), ...log];
    handleSaveReport(combined);
  }

  return (
    <div className={s.shell}>
      <header className={s.hero}>
        <div className={s.heroDecor} />
        <div className={s.heroDecor2} />
        <div className={s.heroInner}>
          <span className={s.logo}>
            <BookDatabase24Filled />
          </span>
          <div>
            <div className={s.title}>D365 Knowledgebase Loader</div>
            <Text className={s.sub}>
              SharePoint folder → Dynamics 365 Knowledgebase, with preview, edit, and an audit trail.
            </Text>
          </div>
          {isMock && (
            <span className={s.modeChip}>
              <Sparkle20Filled /> Mock mode
            </span>
          )}
          <EnvironmentPicker
            service={svc}
            selected={environment}
            onChange={setEnvironment}
          />
        </div>
      </header>

      <div className={s.stepperBar}>
        <Stepper
          active={stage}
          onSelect={v => setStage(v as Stage)}
          steps={[
            { value: 'config', label: 'Configure' },
            { value: 'review', label: 'Review', count: articles.length, disabled: articles.length === 0 },
            { value: 'progress', label: 'Progress & report' },
          ]}
        />
      </div>

      <main className={s.body}>
        {!environment && (
          <MessageBar intent="warning">
            <MessageBarBody>
              <MessageBarTitle>Pick a target environment</MessageBarTitle>
              Use the <strong>Environment</strong> chip in the header to choose
              which Power Platform environment to load articles into. The app
              will verify the Dynamics 365 Knowledgebase is available there.
            </MessageBarBody>
          </MessageBar>
        )}
        {environment && environment.knowledgebaseStatus === 'missing' && (
          <MessageBar intent="error">
            <MessageBarBody>
              <MessageBarTitle>Knowledgebase not installed</MessageBarTitle>
              <strong>{environment.displayName}</strong> doesn't have the
              <code> knowledgearticle</code> table. Install the Dynamics 365
              Customer Service (or another Knowledge-enabled) solution, or
              pick a different environment.
            </MessageBarBody>
          </MessageBar>
        )}
        {environment && environment.knowledgebaseStatus === 'error' && (
          <MessageBar intent="warning">
            <MessageBarBody>
              <MessageBarTitle>Couldn't verify the Knowledgebase</MessageBarTitle>
              {environment.knowledgebaseError ?? 'The check failed. You can still try loading, but it may not work.'}
            </MessageBarBody>
          </MessageBar>
        )}

        {isMock && (
          <MessageBar intent="info">
            <MessageBarBody>
              <MessageBarTitle>Running with sample data</MessageBarTitle>
              Run <code>pac code add-data-source</code> for SharePoint and Dataverse,
              then set <code>VITE_USE_REAL_CONNECTORS=true</code> in <code>.env.local</code> to use live connectors.
            </MessageBarBody>
          </MessageBar>
        )}

        {stage === 'config' && (
          <ConfigPanel config={config} onChange={setConfig} onScan={handleScan} scanning={scanning} error={scanError} />
        )}
        {stage === 'review' && (
          <ReviewPanel
            articles={articles}
            onChange={setArticles}
            onLoad={handleLoad}
            loading={loading}
            canLoad={envReady}
            disabledReason={
              !environment
                ? 'Pick a Power Platform environment first.'
                : environment.knowledgebaseStatus === 'missing'
                  ? 'The selected environment has no knowledgearticle table.'
                  : environment.knowledgebaseStatus !== 'present'
                    ? 'Knowledgebase availability not confirmed for this environment.'
                    : undefined
            }
          />
        )}
        {stage === 'progress' && (
          <ProgressPanel
            done={done}
            total={articles.filter(a => a.selected).length}
            errors={errors}
            log={log}
            onSaveReport={() => handleSaveReport()}
            reportSaving={reportSaving}
            reportResult={reportResult}
            reportError={reportError}
          />
        )}
      </main>
    </div>
  );
}

function mkLog(fileName: string, action: LogEntry['action'], status: LogEntry['status'], message: string, sourcePath?: string, knowledgeArticleId?: string): LogEntry {
  return { timestamp: new Date().toISOString(), fileName, action, status, message, sourcePath, knowledgeArticleId };
}
