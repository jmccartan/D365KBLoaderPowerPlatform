import { useState } from 'react';
import {
  Card, CardHeader, Text, Button, Checkbox, makeStyles, tokens,
  Tab, TabList, Input, Textarea, Badge, Divider, Tooltip
} from '@fluentui/react-components';
import { Edit24Regular, Eye24Regular, Code24Regular } from '@fluentui/react-icons';
import type { ProcessedArticle } from '../types';

const useStyles = makeStyles({
  wrap: { display: 'grid', gridTemplateColumns: '380px 1fr', gap: tokens.spacingHorizontalL, height: 'calc(100vh - 280px)' },
  list: { overflow: 'auto', padding: tokens.spacingHorizontalS },
  detail: { padding: tokens.spacingHorizontalL, overflow: 'auto', display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM },
  row: { display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS, padding: tokens.spacingVerticalS, borderRadius: tokens.borderRadiusMedium, cursor: 'pointer' },
  rowActive: { backgroundColor: tokens.colorNeutralBackground2Selected },
  preview: { border: `1px solid ${tokens.colorNeutralStroke2}`, borderRadius: tokens.borderRadiusMedium, padding: tokens.spacingHorizontalM, minHeight: '300px', backgroundColor: tokens.colorNeutralBackground1 },
  raw: { fontFamily: 'Consolas, monospace', fontSize: '12px', minHeight: '300px' },
  toolbar: { display: 'flex', gap: tokens.spacingHorizontalM, alignItems: 'center', flexWrap: 'wrap' },
  badges: { display: 'flex', gap: tokens.spacingHorizontalXS }
});

export interface ReviewPanelProps {
  articles: ProcessedArticle[];
  onChange: (articles: ProcessedArticle[]) => void;
  onLoad: () => void;
  loading: boolean;
}

export function ReviewPanel({ articles, onChange, onLoad, loading }: ReviewPanelProps) {
  const s = useStyles();
  const [activeId, setActiveId] = useState<string>(articles[0]?.id ?? '');
  const [tab, setTab] = useState<'preview' | 'raw' | 'edit'>('preview');
  const active = articles.find(a => a.id === activeId) ?? articles[0];

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

  return (
    <Card>
      <CardHeader
        header={<Text weight="semibold" size={500}>2. Review &amp; edit</Text>}
        description={<Text size={200}>{articles.length} files processed · {selectedCount} selected for load</Text>}
        action={
          <div className={s.toolbar}>
            <Checkbox label="Select all" checked={allSelected ? true : someSelected ? 'mixed' : false} onChange={toggleAll} />
            <Button appearance="primary" onClick={onLoad} disabled={loading || selectedCount === 0}>
              Load {selectedCount} into KB
            </Button>
          </div>
        }
      />
      <Divider />
      <div className={s.wrap}>
        <div className={s.list}>
          {articles.map(a => {
            const isActive = a.id === active?.id;
            return (
              <div key={a.id} className={`${s.row} ${isActive ? s.rowActive : ''}`} onClick={() => setActiveId(a.id)}>
                <Checkbox
                  checked={a.selected}
                  onClick={e => e.stopPropagation()}
                  onChange={(_, d) => update(a.id, { selected: !!d.checked })}
                />
                <div style={{ flex: 1, overflow: 'hidden' }}>
                  <Text truncate block weight="semibold">{a.title}</Text>
                  <Text size={100} truncate block>{a.source.name}</Text>
                </div>
                <div className={s.badges}>
                  {a.warnings.length > 0 && (
                    <Tooltip content={a.warnings.join('\n')} relationship="description">
                      <Badge color="warning" size="small">!</Badge>
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
            <Field label="Title">
              <Input value={active.title} onChange={(_, d) => update(active.id, { title: d.value })} />
            </Field>
            <TabList selectedValue={tab} onTabSelect={(_, d) => setTab(d.value as any)}>
              <Tab icon={<Eye24Regular />} value="preview">Preview</Tab>
              <Tab icon={<Edit24Regular />} value="edit">Edit HTML</Tab>
              <Tab icon={<Code24Regular />} value="raw">Raw source</Tab>
            </TabList>
            {tab === 'preview' && (
              <div className={s.preview} dangerouslySetInnerHTML={{ __html: active.html }} />
            )}
            {tab === 'edit' && (
              <Textarea value={active.html} className={s.raw} onChange={(_, d) => update(active.id, { html: d.value })} resize="vertical" rows={18} />
            )}
            {tab === 'raw' && (
              <Textarea value={active.rawHtml} className={s.raw} readOnly resize="vertical" rows={18} />
            )}
            {active.warnings.length > 0 && (
              <div>
                <Text weight="semibold">Conversion warnings</Text>
                <ul>{active.warnings.map((w, i) => <li key={i}><Text size={200}>{w}</Text></li>)}</ul>
              </div>
            )}
            {active.loadStatus === 'error' && active.loadError && (
              <Text style={{ color: tokens.colorPaletteRedForeground1 }}>Error: {active.loadError}</Text>
            )}
          </div>
        )}
      </div>
    </Card>
  );
}

function StatusBadge({ status }: { status: ProcessedArticle['loadStatus'] }) {
  switch (status) {
    case 'success': return <Badge color="success" size="small">Loaded</Badge>;
    case 'error': return <Badge color="danger" size="small">Failed</Badge>;
    case 'loading': return <Badge color="brand" size="small">…</Badge>;
    default: return <Badge color="informative" size="small">Pending</Badge>;
  }
}

// Local Field — Fluent's Field needs a label prop and we want compact usage
function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div>
      <Text weight="semibold" size={200} block style={{ marginBottom: 4 }}>{label}</Text>
      {children}
    </div>
  );
}
