import { useEffect, useMemo, useState } from 'react';
import {
  Card, Text, Button, Checkbox, makeStyles, tokens, mergeClasses,
  Tab, TabList, Input, Textarea, Badge, Tooltip, Spinner, Divider,
  MessageBar, MessageBarBody, MessageBarTitle,
} from '@fluentui/react-components';
import {
  Edit24Regular, Eye24Regular, Code24Regular, DocumentText24Regular,
  Warning20Filled, CloudArrowUp20Filled, CheckmarkCircle20Filled, DismissCircle20Filled,
  Sparkle20Filled, BranchCompare20Regular, Open16Regular,
} from '@fluentui/react-icons';
import type { ProcessedArticle, ArticleSuggestion, OverlapMatch, PIIFinding } from '../types';
import { sanitizeArticleHtml } from '../processing/pipeline';
import { getService } from '../services';
import { RichTextEditor } from './RichTextEditor';
import { SuggestEditsDialog } from './SuggestEditsDialog';

const useStyles = makeStyles({
  card: {
    padding: 0,
    overflow: 'hidden',
    boxShadow: tokens.shadow8,
    borderRadius: tokens.borderRadiusXLarge,
    display: 'flex',
    flexDirection: 'column',
    minHeight: 'calc(100vh - 280px)',
  },
  accent: {
    height: '4px',
    background: 'linear-gradient(90deg, #0B3A6F 0%, #1278D2 50%, #4F9DE8 100%)',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalL} ${tokens.spacingHorizontalXL}`,
    flexWrap: 'wrap',
  },
  headerText: { display: 'flex', flexDirection: 'column', flex: 1, minWidth: '200px' },
  title: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  sub: { color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 },
  toolbar: {
    display: 'flex',
    gap: tokens.spacingHorizontalM,
    alignItems: 'center',
    flexWrap: 'wrap',
  },
  countChip: {
    display: 'inline-flex',
    alignItems: 'center',
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalS}`,
    borderRadius: tokens.borderRadiusCircular,
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
  },
  wrap: {
    display: 'grid',
    gridTemplateColumns: '380px 1fr',
    gap: 0,
    flex: 1,
    minHeight: 0,
  },
  list: {
    overflow: 'auto',
    padding: tokens.spacingHorizontalS,
    backgroundColor: tokens.colorNeutralBackground2,
    borderRightWidth: '1px',
    borderRightStyle: 'solid',
    borderRightColor: tokens.colorNeutralStroke2,
  },
  row: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderRadius: tokens.borderRadiusMedium,
    cursor: 'pointer',
    borderTopWidth: '1px',
    borderRightWidth: '1px',
    borderBottomWidth: '1px',
    borderLeftWidth: '1px',
    borderTopStyle: 'solid',
    borderRightStyle: 'solid',
    borderBottomStyle: 'solid',
    borderLeftStyle: 'solid',
    borderTopColor: 'transparent',
    borderRightColor: 'transparent',
    borderBottomColor: 'transparent',
    borderLeftColor: 'transparent',
    marginBottom: tokens.spacingVerticalXS,
    transitionProperty: 'background-color, border-color',
    transitionDuration: tokens.durationFast,
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  rowActive: {
    backgroundColor: tokens.colorBrandBackground2,
    borderTopColor: tokens.colorBrandStroke2,
    borderRightColor: tokens.colorBrandStroke2,
    borderBottomColor: tokens.colorBrandStroke2,
    borderLeftColor: tokens.colorBrandStroke2,
    ':hover': { backgroundColor: tokens.colorBrandBackground2 },
  },
  fileIcon: {
    width: '32px',
    height: '32px',
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorBrandForeground1,
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
  },
  badges: { display: 'flex', gap: tokens.spacingHorizontalXS, alignItems: 'center' },
  detail: {
    padding: tokens.spacingHorizontalXL,
    overflow: 'auto',
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  preview: {
    borderTopWidth: '1px',
    borderRightWidth: '1px',
    borderBottomWidth: '1px',
    borderLeftWidth: '1px',
    borderTopStyle: 'solid',
    borderRightStyle: 'solid',
    borderBottomStyle: 'solid',
    borderLeftStyle: 'solid',
    borderTopColor: tokens.colorNeutralStroke2,
    borderRightColor: tokens.colorNeutralStroke2,
    borderBottomColor: tokens.colorNeutralStroke2,
    borderLeftColor: tokens.colorNeutralStroke2,
    borderRadius: tokens.borderRadiusLarge,
    padding: tokens.spacingHorizontalL,
    minHeight: '320px',
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow2,
  },
  raw: {
    fontFamily: 'Consolas, Cascadia Code, monospace',
    fontSize: '12.5px',
    minHeight: '320px',
  },
  labelText: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    marginBottom: tokens.spacingVerticalXS,
    display: 'block',
  },
  warningBox: {
    borderTopWidth: '1px',
    borderRightWidth: '1px',
    borderBottomWidth: '1px',
    borderLeftWidth: '1px',
    borderTopStyle: 'solid',
    borderRightStyle: 'solid',
    borderBottomStyle: 'solid',
    borderLeftStyle: 'solid',
    borderTopColor: tokens.colorPaletteYellowBorder1,
    borderRightColor: tokens.colorPaletteYellowBorder1,
    borderBottomColor: tokens.colorPaletteYellowBorder1,
    borderLeftColor: tokens.colorPaletteYellowBorder1,
    backgroundColor: tokens.colorPaletteYellowBackground1,
    borderRadius: tokens.borderRadiusMedium,
    padding: tokens.spacingHorizontalM,
  },
  empty: {
    padding: tokens.spacingHorizontalXL,
    textAlign: 'center',
    color: tokens.colorNeutralForeground3,
  },
  copilotBar: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: tokens.spacingHorizontalM,
    marginTop: tokens.spacingVerticalS,
    padding: tokens.spacingHorizontalM,
    borderRadius: tokens.borderRadiusLarge,
    background: 'linear-gradient(135deg, rgba(18,120,210,0.08) 0%, rgba(79,157,232,0.12) 100%)',
    borderTopWidth: '1px',
    borderRightWidth: '1px',
    borderBottomWidth: '1px',
    borderLeftWidth: '1px',
    borderTopStyle: 'solid',
    borderRightStyle: 'solid',
    borderBottomStyle: 'solid',
    borderLeftStyle: 'solid',
    borderTopColor: tokens.colorBrandStroke2,
    borderRightColor: tokens.colorBrandStroke2,
    borderBottomColor: tokens.colorBrandStroke2,
    borderLeftColor: tokens.colorBrandStroke2,
  },
});

export interface ReviewPanelProps {
  articles: ProcessedArticle[];
  onChange: (articles: ProcessedArticle[]) => void;
  onLoad: () => void;
  loading: boolean;
  canLoad?: boolean;
  disabledReason?: string;
  blockPiiOnLoad?: boolean;
}

export function ReviewPanel({ articles, onChange, onLoad, loading, canLoad = true, disabledReason, blockPiiOnLoad = false }: ReviewPanelProps) {
  const s = useStyles();
  const svc = useMemo(() => getService(), []);
  const [activeId, setActiveId] = useState<string>(articles[0]?.id ?? '');
  const [tab, setTab] = useState<'preview' | 'raw' | 'edit'>('preview');
  const [editorMode, setEditorMode] = useState<'visual' | 'source'>('visual');
  const active = articles.find(article => article.id === activeId) ?? articles[0];
  const [editDraft, setEditDraft] = useState(active?.html ?? '');

  const [copilotOpen, setCopilotOpen] = useState(false);
  const [copilotLoading, setCopilotLoading] = useState(false);
  const [copilotError, setCopilotError] = useState<string | undefined>();
  const [suggestion, setSuggestion] = useState<ArticleSuggestion | undefined>();
  const [copilotMode, setCopilotMode] = useState<'single' | 'bulk'>('single');
  const [bulkQueue, setBulkQueue] = useState<string[]>([]);
  const [bulkIndex, setBulkIndex] = useState(0);

  const [overlapScanning, setOverlapScanning] = useState(false);
  const [overlapError, setOverlapError] = useState<string | undefined>();
  const [overlapBanner, setOverlapBanner] = useState<string | undefined>();

  useEffect(() => {
    if (!activeId && articles[0]) {
      setActiveId(articles[0].id);
    }
  }, [activeId, articles]);

  useEffect(() => {
    setEditDraft(active?.html ?? '');
  }, [active?.id, active?.html]);

  const suggestionTarget = copilotMode === 'bulk'
    ? articles.find(article => article.id === bulkQueue[bulkIndex])
    : active;
  const allSelected = articles.length > 0 && articles.every(article => article.selected || (blockPiiOnLoad && article.findings.length > 0));
  const someSelected = articles.some(article => article.selected) && !allSelected;
  const selectedCount = articles.filter(article => article.selected).length;
  const selectedWithFindings = articles.filter(article => article.selected && article.findings.length > 0).length;
  const loadTooltip = !canLoad
    ? (disabledReason ?? 'Select an environment first')
    : blockPiiOnLoad && selectedWithFindings > 0
      ? `${selectedWithFindings} selected article${selectedWithFindings === 1 ? '' : 's'} contain possible PII. Clean them up or leave blocking mode off to continue.`
      : selectedWithFindings > 0
        ? `${selectedWithFindings} selected article${selectedWithFindings === 1 ? '' : 's'} contain possible PII. Loading is allowed, but review the findings first.`
        : undefined;

  const update = (id: string, patch: Partial<ProcessedArticle>) => {
    onChange(articles.map(article => article.id === id ? { ...article, ...patch } : article));
  };

  const toggleAll = () => {
    const next = !allSelected;
    onChange(articles.map(article => ({
      ...article,
      selected: next && (!blockPiiOnLoad || article.findings.length === 0),
    })));
  };

  async function openSuggestion(target: ProcessedArticle, mode: 'single' | 'bulk', queue: string[] = [], index = 0) {
    setCopilotMode(mode);
    setBulkQueue(queue);
    setBulkIndex(index);
    setActiveId(target.id);
    setCopilotOpen(true);
    setCopilotLoading(true);
    setCopilotError(undefined);
    setSuggestion(undefined);
    try {
      const result = await svc.suggestEdits(target);
      setSuggestion(result);
    } catch (error: unknown) {
      setCopilotError(String(error instanceof Error ? error.message : error));
    } finally {
      setCopilotLoading(false);
    }
  }

  function closeCopilot() {
    setCopilotOpen(false);
    setSuggestion(undefined);
    setCopilotError(undefined);
    setCopilotMode('single');
    setBulkQueue([]);
    setBulkIndex(0);
  }

  async function requestSuggestions(target: ProcessedArticle) {
    await openSuggestion(target, 'single');
  }

  async function startBulkSuggestions() {
    const queue = articles.filter(article => article.selected).map(article => article.id);
    if (queue.length === 0) {
      return;
    }
    const first = articles.find(article => article.id === queue[0]);
    if (!first) {
      return;
    }
    await openSuggestion(first, 'bulk', queue, 0);
  }

  async function goToNextBulk(currentIndex: number) {
    const nextIndex = currentIndex + 1;
    if (nextIndex >= bulkQueue.length) {
      closeCopilot();
      return;
    }
    const nextTarget = articles.find(article => article.id === bulkQueue[nextIndex]);
    if (!nextTarget) {
      closeCopilot();
      return;
    }
    await openSuggestion(nextTarget, 'bulk', bulkQueue, nextIndex);
  }

  async function handleDeclineSuggestion() {
    if (copilotMode === 'bulk') {
      await goToNextBulk(bulkIndex);
      return;
    }
    closeCopilot();
  }

  async function acceptSuggestion(target: ProcessedArticle, nextSuggestion: ArticleSuggestion) {
    const cleaned = sanitizeArticleHtml(nextSuggestion.html);
    update(target.id, {
      html: cleaned,
      ...(nextSuggestion.title ? { title: nextSuggestion.title } : {}),
    });
    setEditDraft(cleaned);
    if (copilotMode === 'bulk') {
      await goToNextBulk(bulkIndex);
      return;
    }
    closeCopilot();
  }

  function commitEditHtml() {
    if (!active) {
      return;
    }
    const cleaned = sanitizeArticleHtml(editDraft);
    setEditDraft(cleaned);
    update(active.id, { html: cleaned });
  }

  async function scanForOverlap() {
    setOverlapScanning(true);
    setOverlapError(undefined);
    setOverlapBanner(undefined);
    try {
      const map = await svc.findOverlaps(articles);
      let autoDeselected = 0;
      let flagged = 0;
      const next = articles.map(article => {
        const overlaps = map[article.id] ?? [];
        if (overlaps.length === 0) return { ...article, overlaps: [] };
        flagged++;
        const topScore = overlaps[0].score;
        const shouldAutoDeselect = topScore >= 0.8 && article.selected;
        if (shouldAutoDeselect) autoDeselected++;
        return {
          ...article,
          overlaps,
          selected: shouldAutoDeselect ? false : article.selected,
        };
      });
      onChange(next);
      if (flagged === 0) {
        setOverlapBanner('No likely duplicates found in the existing knowledgebase.');
      } else if (autoDeselected > 0) {
        setOverlapBanner(`${flagged} article${flagged === 1 ? '' : 's'} flagged for potential overlap. ${autoDeselected} high-confidence duplicate${autoDeselected === 1 ? '' : 's'} auto-deselected.`);
      } else {
        setOverlapBanner(`${flagged} article${flagged === 1 ? '' : 's'} flagged for potential overlap — review before loading.`);
      }
    } catch (error: unknown) {
      setOverlapError(String(error instanceof Error ? error.message : error));
    } finally {
      setOverlapScanning(false);
    }
  }

  return (
    <Card className={s.card}>
      <div className={s.accent} />
      <div className={s.header}>
        <div className={s.headerText}>
          <span className={s.title}>Review &amp; edit</span>
          <span className={s.sub}>
            {articles.length} files processed
            {' · '}
            <span className={s.countChip}>{selectedCount} selected</span>
          </span>
        </div>
        <div className={s.toolbar}>
          <Checkbox
            label="Select all"
            checked={allSelected ? true : someSelected ? 'mixed' : false}
            onChange={toggleAll}
          />
          <Button
            appearance="secondary"
            size="large"
            icon={<Sparkle20Filled />}
            onClick={() => { void startBulkSuggestions(); }}
            disabled={selectedCount === 0 || copilotLoading}
          >
            Apply Copilot to selected
          </Button>
          <Button
            appearance="secondary"
            size="large"
            icon={overlapScanning ? <Spinner size="tiny" /> : <BranchCompare20Regular />}
            onClick={scanForOverlap}
            disabled={overlapScanning || articles.length === 0 || !canLoad}
            title={
              !canLoad
                ? (disabledReason ?? 'Select an environment first')
                : 'Compare these candidates against existing D365 KB articles'
            }
          >
            {overlapScanning ? 'Scanning…' : 'Scan for overlap'}
          </Button>
          <Tooltip content={loadTooltip ?? ''} relationship="label" visible={loadTooltip ? undefined : false}>
            <Button
              appearance="primary"
              size="large"
              icon={loading ? <Spinner size="tiny" /> : <CloudArrowUp20Filled />}
              onClick={onLoad}
              disabled={loading || selectedCount === 0 || !canLoad || (blockPiiOnLoad && selectedWithFindings > 0)}
            >
              Load {selectedCount} into KB
            </Button>
          </Tooltip>
        </div>
      </div>
      {(overlapBanner || overlapError) && (
        <div style={{ padding: `0 ${tokens.spacingHorizontalXL} ${tokens.spacingVerticalM}` }}>
          {overlapError ? (
            <MessageBar intent="error">
              <MessageBarBody>
                <MessageBarTitle>Overlap scan failed</MessageBarTitle>
                {overlapError}
              </MessageBarBody>
            </MessageBar>
          ) : (
            <MessageBar intent={overlapBanner!.startsWith('No likely') ? 'success' : 'warning'}>
              <MessageBarBody>{overlapBanner}</MessageBarBody>
            </MessageBar>
          )}
        </div>
      )}
      <Divider />
      <div className={s.wrap}>
        <div className={s.list}>
          {articles.length === 0 && (
            <div className={s.empty}>No articles yet.</div>
          )}
          {articles.map(article => {
            const isActive = article.id === active?.id;
            const piiSummary = summarizeFindings(article.findings);
            const selectionBlocked = blockPiiOnLoad && article.findings.length > 0;
            return (
              <div
                key={article.id}
                className={mergeClasses(s.row, isActive && s.rowActive)}
                onClick={() => setActiveId(article.id)}
              >
                <Tooltip content={selectionBlocked ? 'PII blocking is enabled for this run. Resolve the finding or switch to warn-only mode to select this article.' : ''} relationship="label" visible={selectionBlocked ? undefined : false}>
                  <Checkbox
                    checked={article.selected}
                    disabled={selectionBlocked}
                    onClick={event => event.stopPropagation()}
                    onChange={(_, data) => update(article.id, { selected: !!data.checked })}
                  />
                </Tooltip>
                <span className={s.fileIcon}><DocumentText24Regular /></span>
                <div style={{ flex: 1, overflow: 'hidden' }}>
                  <Text truncate block weight="semibold">{article.title}</Text>
                  <Text size={100} truncate block style={{ color: tokens.colorNeutralForeground3 }}>
                    {article.source.name}
                  </Text>
                </div>
                <div className={s.badges}>
                  {article.findings.length > 0 && (
                    <Tooltip content={piiSummary} relationship="description">
                      <Badge color="warning" appearance="tint" size="small">PII</Badge>
                    </Tooltip>
                  )}
                  {article.overlaps && article.overlaps.length > 0 && (
                    <Tooltip
                      content={`Top match: ${article.overlaps[0].article.title} (${Math.round(article.overlaps[0].score * 100)}%)`}
                      relationship="description"
                    >
                      <Badge color="warning" appearance="tint" size="small">
                        {article.overlaps.length} overlap{article.overlaps.length === 1 ? '' : 's'}
                      </Badge>
                    </Tooltip>
                  )}
                  {article.warnings.length > 0 && (
                    <Tooltip content={article.warnings.join('\n')} relationship="description">
                      <Warning20Filled style={{ color: tokens.colorPaletteYellowForeground1 }} />
                    </Tooltip>
                  )}
                  <StatusBadge status={article.loadStatus} />
                </div>
              </div>
            );
          })}
        </div>
        {active && (
          <div className={s.detail}>
            <div>
              <span className={s.labelText}>Title</span>
              <Input
                value={active.title}
                onChange={(_, data) => update(active.id, { title: data.value })}
                size="large"
              />
            </div>
            <TabList selectedValue={tab} onTabSelect={(_, data) => setTab(data.value as 'preview' | 'edit' | 'raw')} appearance="subtle-circular">
              <Tab icon={<Eye24Regular />} value="preview">Preview</Tab>
              <Tab icon={<Edit24Regular />} value="edit">Edit HTML</Tab>
              <Tab icon={<Code24Regular />} value="raw">Raw source</Tab>
            </TabList>
            {tab === 'preview' && (
              <div className={s.preview} dangerouslySetInnerHTML={{ __html: active.html }} />
            )}
            {tab === 'edit' && (
              <>
                <TabList
                  selectedValue={editorMode}
                  onTabSelect={(_, data) => {
                    commitEditHtml();
                    setEditorMode(data.value as 'visual' | 'source');
                  }}
                  appearance="subtle-circular"
                >
                  <Tab value="visual">Visual</Tab>
                  <Tab value="source">Source</Tab>
                </TabList>
                {editorMode === 'visual' ? (
                  <RichTextEditor value={editDraft} onCommit={html => { setEditDraft(html); update(active.id, { html }); }} />
                ) : (
                  <Textarea
                    value={editDraft}
                    className={s.raw}
                    onChange={(_, data) => setEditDraft(data.value)}
                    onBlur={commitEditHtml}
                    resize="vertical"
                    rows={18}
                  />
                )}
                <div className={s.copilotBar}>
                  <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                    Let Copilot review this article and propose structure and clarity improvements.
                  </Text>
                  <Button
                    appearance="primary"
                    icon={<Sparkle20Filled />}
                    onClick={() => { void requestSuggestions(active); }}
                    disabled={copilotLoading}
                  >
                    Suggest edits with Copilot
                  </Button>
                </div>
              </>
            )}
            {tab === 'raw' && (
              <Textarea value={active.rawHtml} className={s.raw} readOnly resize="vertical" rows={18} />
            )}
            {active.findings.length > 0 && (
              <MessageBar intent="warning">
                <MessageBarBody>
                  <MessageBarTitle>Possible sensitive content detected</MessageBarTitle>
                  <ul style={{ margin: `${tokens.spacingVerticalXS} 0 0`, paddingLeft: 18 }}>
                    {active.findings.map(finding => (
                      <li key={finding.kind}>
                        <Text size={200}>
                          {finding.kind} × {finding.count}
                          {finding.snippets.length > 0 ? ` — ${finding.snippets.join(' · ')}` : ''}
                        </Text>
                      </li>
                    ))}
                  </ul>
                </MessageBarBody>
              </MessageBar>
            )}
            {active.knowledgeArticleUrl && active.loadStatus === 'success' && (
              <MessageBar intent="success">
                <MessageBarBody>
                  Loaded successfully. <a href={active.knowledgeArticleUrl} target="_blank" rel="noreferrer" style={{ color: tokens.colorBrandForeground1, fontWeight: tokens.fontWeightSemibold }}>Open in D365 <Open16Regular style={{ verticalAlign: 'middle' }} /></a>
                </MessageBarBody>
              </MessageBar>
            )}
            {active.overlaps && active.overlaps.length > 0 && (
              <OverlapSection overlaps={active.overlaps} />
            )}
            {active.warnings.length > 0 && (
              <div className={s.warningBox}>
                <Text weight="semibold" block style={{ marginBottom: 4 }}>Conversion warnings</Text>
                <ul style={{ margin: 0, paddingLeft: 18 }}>
                  {active.warnings.map((warning, index) => (
                    <li key={index}><Text size={200}>{warning}</Text></li>
                  ))}
                </ul>
              </div>
            )}
            {active.loadStatus === 'error' && active.loadError && (
              <Text style={{ color: tokens.colorPaletteRedForeground1 }}>
                Error: {active.loadError}
              </Text>
            )}
          </div>
        )}
      </div>

      <SuggestEditsDialog
        open={copilotOpen}
        loading={copilotLoading}
        error={copilotError}
        suggestion={suggestion}
        currentTitle={suggestionTarget?.title ?? ''}
        titleText={copilotMode === 'bulk' ? 'Copilot review queue' : 'Copilot suggestions'}
        subtitleText={suggestionTarget ? `${suggestionTarget.title} · ${suggestionTarget.source.name}` : undefined}
        queueLabel={copilotMode === 'bulk' ? `${bulkIndex + 1} of ${bulkQueue.length}` : undefined}
        onAccept={nextSuggestion => suggestionTarget && acceptSuggestion(suggestionTarget, nextSuggestion)}
        onDecline={() => { void handleDeclineSuggestion(); }}
        onRegenerate={() => suggestionTarget && void openSuggestion(suggestionTarget, copilotMode, bulkQueue, bulkIndex)}
        onSkipAll={copilotMode === 'bulk' ? closeCopilot : undefined}
      />
    </Card>
  );
}

function summarizeFindings(findings: PIIFinding[]): string {
  return findings.map(finding => `${finding.kind} × ${finding.count}${finding.snippets.length > 0 ? ` — ${finding.snippets.join(' · ')}` : ''}`).join('\n');
}

function StatusBadge({ status }: { status: ProcessedArticle['loadStatus'] }) {
  switch (status) {
    case 'success':
      return <Badge color="success" appearance="tint" icon={<CheckmarkCircle20Filled />}>Loaded</Badge>;
    case 'error':
      return <Badge color="danger" appearance="tint" icon={<DismissCircle20Filled />}>Failed</Badge>;
    case 'loading':
      return <Badge color="brand" appearance="tint">Loading…</Badge>;
    case 'skipped':
      return <Badge color="informative" appearance="tint">Skipped</Badge>;
    default:
      return <Badge color="informative" appearance="tint">Pending</Badge>;
  }
}

const useOverlapStyles = makeStyles({
  wrap: {
    borderTopWidth: '1px',
    borderRightWidth: '1px',
    borderBottomWidth: '1px',
    borderLeftWidth: '1px',
    borderTopStyle: 'solid',
    borderRightStyle: 'solid',
    borderBottomStyle: 'solid',
    borderLeftStyle: 'solid',
    borderTopColor: tokens.colorPaletteDarkOrangeBorder1,
    borderRightColor: tokens.colorPaletteDarkOrangeBorder1,
    borderBottomColor: tokens.colorPaletteDarkOrangeBorder1,
    borderLeftColor: tokens.colorPaletteDarkOrangeBorder1,
    backgroundColor: tokens.colorPaletteDarkOrangeBackground1,
    borderRadius: tokens.borderRadiusLarge,
    padding: tokens.spacingHorizontalL,
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    marginBottom: tokens.spacingVerticalS,
  },
  item: {
    display: 'flex',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalS} 0`,
    borderTopWidth: '0',
    borderRightWidth: '0',
    borderBottomWidth: '1px',
    borderLeftWidth: '0',
    borderBottomStyle: 'solid',
    borderBottomColor: tokens.colorPaletteDarkOrangeBorder1,
    ':last-child': { borderBottomWidth: '0' },
  },
  score: {
    flexShrink: 0,
    width: '64px',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: '2px',
  },
  scoreBar: {
    width: '100%',
    height: '6px',
    borderRadius: '3px',
    backgroundColor: tokens.colorNeutralBackground4,
    overflow: 'hidden',
  },
  scoreFill: {
    height: '100%',
    backgroundColor: tokens.colorPaletteDarkOrangeForeground1,
  },
  scoreLabel: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorPaletteDarkOrangeForeground1,
  },
  body: { flex: 1, minWidth: 0 },
  reasons: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: tokens.spacingHorizontalXS,
    marginTop: tokens.spacingVerticalXS,
  },
});

function OverlapSection({ overlaps }: { overlaps: OverlapMatch[] }) {
  const o = useOverlapStyles();
  return (
    <div className={o.wrap}>
      <div className={o.header}>
        <BranchCompare20Regular style={{ color: tokens.colorPaletteDarkOrangeForeground1 }} />
        <Text weight="semibold">
          Possible overlap with existing KB ({overlaps.length})
        </Text>
      </div>
      <Text size={200} block style={{ color: tokens.colorNeutralForeground2, marginBottom: tokens.spacingVerticalS }}>
        Consider unchecking this article if the matches below already cover the topic, or update the existing article instead of creating a duplicate.
      </Text>
      {overlaps.map(match => (
        <div key={match.article.id} className={o.item}>
          <div className={o.score}>
            <div className={o.scoreBar}>
              <div className={o.scoreFill} style={{ width: `${Math.round(match.score * 100)}%` }} />
            </div>
            <span className={o.scoreLabel}>{Math.round(match.score * 100)}%</span>
          </div>
          <div className={o.body}>
            <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalXS }}>
              <Text weight="semibold">{match.article.title}</Text>
              {match.article.url && (
                <a href={match.article.url} target="_blank" rel="noreferrer" title="Open in D365"
                   style={{ display: 'inline-flex', color: tokens.colorBrandForeground1 }}>
                  <Open16Regular />
                </a>
              )}
              {match.article.modifiedOn && (
                <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>
                  · modified {new Date(match.article.modifiedOn).toLocaleDateString()}
                </Text>
              )}
            </div>
            {match.article.excerpt && (
              <Text size={200} block style={{ color: tokens.colorNeutralForeground2 }}>
                {match.article.excerpt}
              </Text>
            )}
            <div className={o.reasons}>
              {match.reasons.map((reason, index) => (
                <Badge key={index} appearance="outline" color="warning" size="small">{reason}</Badge>
              ))}
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}
