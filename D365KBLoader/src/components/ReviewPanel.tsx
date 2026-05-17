import { useMemo, useState } from 'react';
import {
  Card, Text, Button, Checkbox, makeStyles, tokens, mergeClasses,
  Tab, TabList, Input, Textarea, Badge, Tooltip, Spinner, Divider
} from '@fluentui/react-components';
import {
  Edit24Regular, Eye24Regular, Code24Regular, DocumentText24Regular,
  Warning20Filled, CloudArrowUp20Filled, CheckmarkCircle20Filled, DismissCircle20Filled,
  Sparkle20Filled,
} from '@fluentui/react-icons';
import type { ProcessedArticle, ArticleSuggestion } from '../types';
import { getService } from '../services';
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
    borderRight: `1px solid ${tokens.colorNeutralStroke2}`,
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
    border: `1px solid ${tokens.colorNeutralStroke2}`,
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
    border: `1px solid ${tokens.colorPaletteYellowBorder1}`,
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
}

export function ReviewPanel({ articles, onChange, onLoad, loading }: ReviewPanelProps) {
  const s = useStyles();
  const svc = useMemo(() => getService(), []);
  const [activeId, setActiveId] = useState<string>(articles[0]?.id ?? '');
  const [tab, setTab] = useState<'preview' | 'raw' | 'edit'>('preview');
  const active = articles.find(a => a.id === activeId) ?? articles[0];

  // Copilot suggestion state
  const [copilotOpen, setCopilotOpen] = useState(false);
  const [copilotLoading, setCopilotLoading] = useState(false);
  const [copilotError, setCopilotError] = useState<string | undefined>();
  const [suggestion, setSuggestion] = useState<ArticleSuggestion | undefined>();

  const allSelected = articles.length > 0 && articles.every(a => a.selected);
  const someSelected = articles.some(a => a.selected) && !allSelected;
  const selectedCount = articles.filter(a => a.selected).length;

  const update = (id: string, patch: Partial<ProcessedArticle>) => {
    onChange(articles.map(a => a.id === id ? { ...a, ...patch } : a));
  };
  const toggleAll = () => {
    const next = !allSelected;
    onChange(articles.map(a => ({ ...a, selected: next })));
  };

  async function requestSuggestions(target: ProcessedArticle) {
    setCopilotOpen(true);
    setCopilotLoading(true);
    setCopilotError(undefined);
    setSuggestion(undefined);
    try {
      const result = await svc.suggestEdits(target);
      setSuggestion(result);
    } catch (e: any) {
      setCopilotError(String(e?.message ?? e));
    } finally {
      setCopilotLoading(false);
    }
  }

  function acceptSuggestion(target: ProcessedArticle, sugg: ArticleSuggestion) {
    update(target.id, {
      html: sugg.html,
      ...(sugg.title ? { title: sugg.title } : {}),
    });
    setCopilotOpen(false);
    setSuggestion(undefined);
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
            appearance="primary"
            size="large"
            icon={loading ? <Spinner size="tiny" /> : <CloudArrowUp20Filled />}
            onClick={onLoad}
            disabled={loading || selectedCount === 0}
          >
            Load {selectedCount} into KB
          </Button>
        </div>
      </div>
      <Divider />
      <div className={s.wrap}>
        <div className={s.list}>
          {articles.length === 0 && (
            <div className={s.empty}>No articles yet.</div>
          )}
          {articles.map(a => {
            const isActive = a.id === active?.id;
            return (
              <div
                key={a.id}
                className={mergeClasses(s.row, isActive && s.rowActive)}
                onClick={() => setActiveId(a.id)}
              >
                <Checkbox
                  checked={a.selected}
                  onClick={e => e.stopPropagation()}
                  onChange={(_, d) => update(a.id, { selected: !!d.checked })}
                />
                <span className={s.fileIcon}><DocumentText24Regular /></span>
                <div style={{ flex: 1, overflow: 'hidden' }}>
                  <Text truncate block weight="semibold">{a.title}</Text>
                  <Text size={100} truncate block style={{ color: tokens.colorNeutralForeground3 }}>
                    {a.source.name}
                  </Text>
                </div>
                <div className={s.badges}>
                  {a.warnings.length > 0 && (
                    <Tooltip content={a.warnings.join('\n')} relationship="description">
                      <Warning20Filled style={{ color: tokens.colorPaletteYellowForeground1 }} />
                    </Tooltip>
                  )}
                  <StatusBadge status={a.loadStatus} />
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
                onChange={(_, d) => update(active.id, { title: d.value })}
                size="large"
              />
            </div>
            <TabList selectedValue={tab} onTabSelect={(_, d) => setTab(d.value as any)} appearance="subtle-circular">
              <Tab icon={<Eye24Regular />} value="preview">Preview</Tab>
              <Tab icon={<Edit24Regular />} value="edit">Edit HTML</Tab>
              <Tab icon={<Code24Regular />} value="raw">Raw source</Tab>
            </TabList>
            {tab === 'preview' && (
              <div className={s.preview} dangerouslySetInnerHTML={{ __html: active.html }} />
            )}
            {tab === 'edit' && (
              <>
                <Textarea
                  value={active.html}
                  className={s.raw}
                  onChange={(_, d) => update(active.id, { html: d.value })}
                  resize="vertical"
                  rows={18}
                />
                <div className={s.copilotBar}>
                  <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                    Let Copilot review this article and propose structure and clarity improvements.
                  </Text>
                  <Button
                    appearance="primary"
                    icon={<Sparkle20Filled />}
                    onClick={() => requestSuggestions(active)}
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
            {active.warnings.length > 0 && (
              <div className={s.warningBox}>
                <Text weight="semibold" block style={{ marginBottom: 4 }}>Conversion warnings</Text>
                <ul style={{ margin: 0, paddingLeft: 18 }}>
                  {active.warnings.map((w, i) => (
                    <li key={i}><Text size={200}>{w}</Text></li>
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
        currentTitle={active?.title ?? ''}
        onAccept={sugg => active && acceptSuggestion(active, sugg)}
        onDecline={() => { setCopilotOpen(false); setSuggestion(undefined); }}
        onRegenerate={() => active && requestSuggestions(active)}
      />
    </Card>
  );
}

function StatusBadge({ status }: { status: ProcessedArticle['loadStatus'] }) {
  switch (status) {
    case 'success':
      return <Badge color="success" appearance="tint" icon={<CheckmarkCircle20Filled />}>Loaded</Badge>;
    case 'error':
      return <Badge color="danger" appearance="tint" icon={<DismissCircle20Filled />}>Failed</Badge>;
    case 'loading':
      return <Badge color="brand" appearance="tint">Loading…</Badge>;
    default:
      return <Badge color="informative" appearance="tint">Pending</Badge>;
  }
}
