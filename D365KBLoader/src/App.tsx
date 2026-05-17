import { useEffect, useMemo, useState } from 'react';
import {
  makeStyles, tokens, Text, MessageBar, MessageBarBody, MessageBarTitle,
} from '@fluentui/react-components';
import { BookDatabase24Filled, Sparkle20Filled } from '@fluentui/react-icons';
import { ConfigPanel } from './components/ConfigPanel';
import { KbDefaultsCard } from './components/KbDefaultsCard';
import { LocalFilesDropZone } from './components/LocalFilesDropZone';
import { ProfilesBar } from './components/ProfilesBar';
import { ReviewPanel } from './components/ReviewPanel';
import { ProgressPanel } from './components/ProgressPanel';
import { Stepper } from './components/Stepper';
import { EnvironmentPicker } from './components/EnvironmentPicker';
import { processFile } from './processing/pipeline';
import { buildReportWorkbook } from './reporting/report';
import { getService } from './services';
import { findPii } from './services/piiScan';
import { heroGradient } from './theme';
import type { KbConfig, ProcessedArticle, LogEntry, ReportResult, PowerPlatformEnvironment, KbUser, SavedScanProfile, SourceFile } from './types';

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
    borderTopWidth: '1px',
    borderRightWidth: '1px',
    borderBottomWidth: '1px',
    borderLeftWidth: '1px',
    borderTopStyle: 'solid',
    borderRightStyle: 'solid',
    borderBottomStyle: 'solid',
    borderLeftStyle: 'solid',
    borderTopColor: 'rgba(255,255,255,0.3)',
    borderRightColor: 'rgba(255,255,255,0.3)',
    borderBottomColor: 'rgba(255,255,255,0.3)',
    borderLeftColor: 'rgba(255,255,255,0.3)',
  },
  stepperBar: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottomWidth: '1px',
    borderBottomStyle: 'solid',
    borderBottomColor: tokens.colorNeutralStroke2,
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
type EmailStatus = { kind: 'success' | 'error'; message: string } | undefined;

export function App() {
  const s = useStyles();
  const svc = useMemo(() => getService(), []);

  const [config, setConfig] = useState<KbConfig>({
    siteUrl: '',
    folderPath: '',
    publishOnLoad: false,
    duplicateAction: 'skip',
    recursive: false,
    incremental: false,
    blockPiiOnLoad: false,
  });

  const [stage, setStage] = useState<Stage>('config');
  const [scanning, setScanning] = useState(false);
  const [scanError, setScanError] = useState<string | undefined>();
  const [articles, setArticles] = useState<ProcessedArticle[]>([]);
  const [profiles, setProfiles] = useState<SavedScanProfile[]>([]);
  const [profilesLoading, setProfilesLoading] = useState(false);
  const [selectedProfileId, setSelectedProfileId] = useState<string | undefined>();

  const [loading, setLoading] = useState(false);
  const [done, setDone] = useState(0);
  const [errors, setErrors] = useState(0);
  const [log, setLog] = useState<LogEntry[]>([]);
  const [reportSaving, setReportSaving] = useState(false);
  const [reportResult, setReportResult] = useState<ReportResult | undefined>();
  const [reportError, setReportError] = useState<string | undefined>();
  const [emailSending, setEmailSending] = useState(false);
  const [emailStatus, setEmailStatus] = useState<EmailStatus>();
  const [environment, setEnvironment] = useState<PowerPlatformEnvironment | undefined>();
  const [currentUser, setCurrentUser] = useState<KbUser | undefined>();

  const isMock = import.meta.env?.VITE_USE_REAL_CONNECTORS !== 'true';
  const envReady = !!environment && environment.knowledgebaseStatus === 'present';
  const uploadImage = svc.uploadImage ? svc.uploadImage.bind(svc) : undefined;

  useEffect(() => {
    svc.getCurrentUser().then(setCurrentUser).catch(() => undefined);
  }, [svc]);

  useEffect(() => {
    void refreshProfiles();
  }, [svc]);

  useEffect(() => {
    if (!config.blockPiiOnLoad) {
      return;
    }
    setArticles(previous => previous.map(article => article.findings.length > 0 ? { ...article, selected: false } : article));
  }, [config.blockPiiOnLoad]);

  async function refreshProfiles() {
    setProfilesLoading(true);
    try {
      setProfiles(await svc.listProfiles());
    } finally {
      setProfilesLoading(false);
    }
  }

  function updateConfig(next: KbConfig) {
    setConfig(next);
  }

  function mkLog(fileName: string, action: LogEntry['action'], status: LogEntry['status'], message: string, sourcePath?: string, knowledgeArticleId?: string): LogEntry {
    return {
      timestamp: new Date().toISOString(),
      fileName,
      action,
      status,
      message,
      sourcePath,
      knowledgeArticleId,
      userDisplayName: currentUser?.displayName,
      userEmail: currentUser?.email,
    };
  }

  function appendLog(entry: LogEntry) {
    setLog(previous => [{ ...entry, id: String(previous.length + 1) }, ...previous]);
  }

  async function handleSaveReport(currentLog?: LogEntry[]) {
    setReportSaving(true);
    setReportError(undefined);
    try {
      const result = await svc.writeReport(config, currentLog ?? log);
      setReportResult(result);
    } catch (error: unknown) {
      setReportError(String(error instanceof Error ? error.message : error));
    } finally {
      setReportSaving(false);
    }
  }

  async function handleEmailReport(to: string[], subject: string, html: string) {
    setEmailSending(true);
    setEmailStatus(undefined);
    try {
      if (!svc.emailReport) {
        throw new Error('Email report is not wired for this service.');
      }
      const attachment = await buildReportWorkbook(config, log);
      await svc.emailReport(to, subject, html, attachment);
      setEmailStatus({ kind: 'success', message: `Sent ${attachment.fileName} to ${to.join(', ')}.` });
    } catch (error: unknown) {
      setEmailStatus({ kind: 'error', message: String(error instanceof Error ? error.message : error) });
      throw error;
    } finally {
      setEmailSending(false);
    }
  }

  async function buildProcessedArticle(file: SourceFile, buffer: ArrayBuffer): Promise<ProcessedArticle> {
    const processed = await processFile(file, buffer, uploadImage);
    const findings = findPii(`${processed.title}\n${processed.html}`);
    return {
      ...processed,
      findings,
      selected: findings.length > 0 && config.blockPiiOnLoad ? false : processed.selected,
    };
  }

  async function handleScan() {
    setScanning(true);
    setScanError(undefined);
    setLog([]);
    setReportResult(undefined);
    setReportError(undefined);
    setEmailStatus(undefined);
    try {
      const files = await svc.listFiles(config);
      const supported = files.filter(file => file.kind !== 'unsupported');
      const skipped = files.filter(file => file.kind === 'unsupported');

      let alreadyLoaded = new Set<string>();
      if (config.incremental && svc.readPriorReports) {
        try {
          alreadyLoaded = await svc.readPriorReports(config);
        } catch {
          alreadyLoaded = new Set<string>();
        }
      }

      const toProcess = supported.filter(file => !alreadyLoaded.has(file.path));
      const incSkipped = supported.length - toProcess.length;
      if (incSkipped > 0) {
        appendLog(mkLog('(incremental)', 'skip', 'info', `Skipped ${incSkipped} files already loaded in a prior run.`));
      }

      const processed: ProcessedArticle[] = [];
      for (const file of toProcess) {
        try {
          const buffer = await svc.downloadFile(file);
          const article = await buildProcessedArticle(file, buffer);
          processed.push(article);
          appendLog(mkLog(file.name, 'process', 'success', 'Converted to HTML', file.path));
          if (article.findings.length > 0) {
            appendLog(mkLog(file.name, 'process', 'info', `Sensitive-content scan flagged ${article.findings.map(finding => `${finding.kind} × ${finding.count}`).join(', ')}`, file.path));
          }
        } catch (error: unknown) {
          const message = String(error instanceof Error ? error.message : error);
          processed.push({
            id: file.id,
            source: file,
            title: file.name,
            html: '',
            rawHtml: '',
            warnings: [message],
            findings: [],
            selected: false,
            loadStatus: 'error',
            loadError: message,
          });
          appendLog(mkLog(file.name, 'process', 'error', message, file.path));
        }
      }
      for (const file of skipped) {
        appendLog(mkLog(file.name, 'skip', 'info', 'Unsupported type', file.path));
      }
      setArticles(processed);
      setStage('review');
    } catch (error: unknown) {
      setScanError(String(error instanceof Error ? error.message : error));
    } finally {
      setScanning(false);
    }
  }

  async function ingestLocalFiles(items: Array<{ file: File; source: SourceFile }>) {
    setScanning(true);
    setScanError(undefined);
    try {
      const processed: ProcessedArticle[] = [...articles];
      for (const item of items) {
        if (item.source.kind === 'unsupported') {
          appendLog(mkLog(item.file.name, 'skip', 'info', 'Unsupported type', item.source.path));
          continue;
        }
        try {
          const buffer = await item.file.arrayBuffer();
          const article = await buildProcessedArticle(item.source, buffer);
          processed.push(article);
          appendLog(mkLog(item.file.name, 'process', 'success', 'Converted to HTML (local upload)', item.source.path));
          if (article.findings.length > 0) {
            appendLog(mkLog(item.file.name, 'process', 'info', `Sensitive-content scan flagged ${article.findings.map(finding => `${finding.kind} × ${finding.count}`).join(', ')}`, item.source.path));
          }
        } catch (error: unknown) {
          appendLog(mkLog(item.file.name, 'process', 'error', String(error instanceof Error ? error.message : error), item.source.path));
        }
      }
      setArticles(processed);
      if (processed.length > 0) setStage('review');
    } finally {
      setScanning(false);
    }
  }

  async function handleLoad() {
    setLoading(true);
    setStage('progress');
    setDone(0);
    setErrors(0);
    setReportResult(undefined);
    setReportError(undefined);
    setEmailStatus(undefined);
    const targets = articles.filter(article => article.selected && article.loadStatus !== 'success' && (!config.blockPiiOnLoad || article.findings.length === 0));
    let completed = 0;
    let failureCount = 0;
    const newEntries: LogEntry[] = [];

    for (const article of targets) {
      setArticles(previous => previous.map(candidate => candidate.id === article.id ? { ...candidate, loadStatus: 'loading', loadError: undefined } : candidate));
      try {
        const existing = config.duplicateAction !== 'create-new'
          ? await svc.findArticleByTitle(article.title)
          : undefined;

        if (existing && config.duplicateAction === 'skip') {
          setArticles(previous => previous.map(candidate => candidate.id === article.id ? {
            ...candidate,
            loadStatus: 'skipped',
            knowledgeArticleId: existing.id,
            knowledgeArticleUrl: existing.url,
          } : candidate));
          const entry = mkLog(article.source.name, 'skip', 'info', `Skipped — article "${existing.title}" already exists`, article.source.path, existing.id);
          newEntries.push(entry);
          appendLog(entry);
        } else if (existing && config.duplicateAction === 'update-existing') {
          await svc.updateKnowledgeArticle(existing.id, article);
          setArticles(previous => previous.map(candidate => candidate.id === article.id ? {
            ...candidate,
            loadStatus: 'success',
            knowledgeArticleId: existing.id,
            knowledgeArticleUrl: existing.url,
          } : candidate));
          const entry = mkLog(article.source.name, 'update', 'success', `Updated existing knowledgearticle "${existing.title}"`, article.source.path, existing.id);
          newEntries.push(entry);
          appendLog(entry);
          completed += 1;
        } else {
          const result = await svc.createKnowledgeArticle(article, config);
          setArticles(previous => previous.map(candidate => candidate.id === article.id ? {
            ...candidate,
            loadStatus: 'success',
            knowledgeArticleId: result.id,
            knowledgeArticleUrl: result.url,
          } : candidate));
          const entry = mkLog(article.source.name, 'load', 'success', 'Created knowledgearticle', article.source.path, result.id);
          newEntries.push(entry);
          appendLog(entry);
          completed += 1;
        }
      } catch (error: unknown) {
        const message = String(error instanceof Error ? error.message : error);
        setArticles(previous => previous.map(candidate => candidate.id === article.id ? { ...candidate, loadStatus: 'error', loadError: message } : candidate));
        const entry = mkLog(article.source.name, 'load', 'error', message, article.source.path);
        newEntries.push(entry);
        appendLog(entry);
        failureCount += 1;
      }
      setDone(completed + failureCount);
      setErrors(failureCount);
    }
    setLoading(false);

    const combined = [...newEntries.slice().reverse(), ...log];
    handleSaveReport(combined);
  }

  async function handleApplyProfile(profile: SavedScanProfile) {
    updateConfig({ ...profile.config });
    setSelectedProfileId(profile.id);
    if (!profile.environmentId) {
      return;
    }
    try {
      const environments = await svc.listEnvironments();
      const match = environments.find(environmentOption => environmentOption.id === profile.environmentId);
      if (match) {
        setEnvironment(await svc.checkKnowledgebase(match));
      }
    } catch {
      // Leave the current environment unchanged if profile environment resolution fails.
    }
  }

  async function handleSaveProfile(name: string) {
    const profile: SavedScanProfile = {
      id: `profile-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      name,
      config: { ...config },
      environmentId: environment?.id,
    };
    const saved = await svc.saveProfile(profile);
    await refreshProfiles();
    setSelectedProfileId(saved.id);
  }

  async function handleDeleteProfile(id: string) {
    await svc.deleteProfile(id);
    await refreshProfiles();
    if (selectedProfileId === id) {
      setSelectedProfileId(undefined);
    }
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
          onSelect={value => setStage(value as Stage)}
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
          <>
            <ProfilesBar
              profiles={profiles}
              selectedProfileId={selectedProfileId}
              loading={profilesLoading}
              onApply={profile => { void handleApplyProfile(profile); }}
              onSave={handleSaveProfile}
              onDelete={handleDeleteProfile}
            />
            <ConfigPanel config={config} onChange={updateConfig} onScan={handleScan} scanning={scanning} error={scanError} />
            <KbDefaultsCard service={svc} config={config} onChange={updateConfig} />
            <Text size={300} weight="semibold" style={{ marginTop: tokens.spacingVerticalL }}>
              …or upload local files
            </Text>
            <LocalFilesDropZone onFiles={ingestLocalFiles} />
          </>
        )}
        {stage === 'review' && (
          <ReviewPanel
            articles={articles}
            onChange={setArticles}
            onLoad={handleLoad}
            loading={loading}
            canLoad={envReady}
            blockPiiOnLoad={!!config.blockPiiOnLoad}
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
            total={articles.filter(article => article.selected).length}
            errors={errors}
            log={log}
            onSaveReport={() => handleSaveReport()}
            reportSaving={reportSaving}
            reportResult={reportResult}
            reportError={reportError}
            canEmailReport={!!svc.emailReport}
            emailSending={emailSending}
            emailStatus={emailStatus}
            onEmailReport={handleEmailReport}
          />
        )}
      </main>
    </div>
  );
}
