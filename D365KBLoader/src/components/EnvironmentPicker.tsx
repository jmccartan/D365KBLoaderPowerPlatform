import { useEffect, useState } from 'react';
import {
  Dialog, DialogSurface, DialogTitle, DialogBody, DialogActions, DialogContent,
  Button, Spinner, Text, makeStyles, tokens, mergeClasses, MessageBar, MessageBarBody, Badge,
} from '@fluentui/react-components';
import {
  Cloud24Filled, CheckmarkCircle20Filled, DismissCircle20Filled, Warning20Filled,
  ArrowClockwise20Regular, ChevronDown16Regular,
} from '@fluentui/react-icons';
import type { KbLoaderService } from '../services/KbLoaderService';
import type { PowerPlatformEnvironment } from '../types';

const useStyles = makeStyles({
  chip: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
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
    cursor: 'pointer',
    ':hover': { backgroundColor: 'rgba(255,255,255,0.26)' },
  },
  chipMeta: { display: 'flex', flexDirection: 'column', alignItems: 'flex-start', lineHeight: 1.15 },
  chipLabel: { fontSize: '10px', textTransform: 'uppercase', letterSpacing: '0.05em', opacity: 0.85 },
  chipValue: { fontSize: tokens.fontSizeBase300 },
  list: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    maxHeight: '480px',
    overflow: 'auto',
    marginTop: tokens.spacingVerticalM,
  },
  row: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: tokens.spacingHorizontalM,
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
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  rowActive: {
    backgroundColor: tokens.colorBrandBackground2,
    borderTopColor: tokens.colorBrandStroke2,
    borderRightColor: tokens.colorBrandStroke2,
    borderBottomColor: tokens.colorBrandStroke2,
    borderLeftColor: tokens.colorBrandStroke2,
    ':hover': { backgroundColor: tokens.colorBrandBackground2 },
  },
  icon: {
    width: '36px',
    height: '36px',
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
  },
  meta: { display: 'flex', flexDirection: 'column', overflow: 'hidden', flex: 1 },
  url: { color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 },
  empty: { padding: tokens.spacingHorizontalL, textAlign: 'center', color: tokens.colorNeutralForeground3 },
  toolbar: { display: 'flex', gap: tokens.spacingHorizontalS, alignItems: 'center', marginBottom: tokens.spacingVerticalS },
});

export interface EnvironmentPickerProps {
  service: KbLoaderService;
  selected?: PowerPlatformEnvironment;
  onChange: (env: PowerPlatformEnvironment) => void;
}

export function EnvironmentPicker({ service, selected, onChange }: EnvironmentPickerProps) {
  const s = useStyles();
  const [open, setOpen] = useState(false);

  return (
    <>
      <button className={s.chip} onClick={() => setOpen(true)} type="button">
        <Cloud24Filled />
        <span className={s.chipMeta}>
          <span className={s.chipLabel}>Environment</span>
          <span className={s.chipValue}>
            {selected ? selected.displayName : 'Select…'}
          </span>
        </span>
        {selected && <KbStatusIcon status={selected.knowledgebaseStatus} />}
        <ChevronDown16Regular />
      </button>
      <EnvironmentDialog
        open={open}
        service={service}
        selectedId={selected?.id}
        onClose={() => setOpen(false)}
        onPick={env => { onChange(env); setOpen(false); }}
      />
    </>
  );
}

interface DialogProps {
  open: boolean;
  service: KbLoaderService;
  selectedId?: string;
  onPick: (env: PowerPlatformEnvironment) => void;
  onClose: () => void;
}

function EnvironmentDialog({ open, service, selectedId, onPick, onClose }: DialogProps) {
  const s = useStyles();
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>();
  const [envs, setEnvs] = useState<PowerPlatformEnvironment[]>([]);
  const [highlight, setHighlight] = useState<string | undefined>(selectedId);

  async function load() {
    setLoading(true);
    setError(undefined);
    try {
      const list = await service.listEnvironments();
      // Mark "checking" then resolve each KB status in parallel
      setEnvs(list.map(e => ({ ...e, knowledgebaseStatus: 'checking' as const })));
      const checked = await Promise.all(list.map(e => service.checkKnowledgebase(e)));
      setEnvs(checked);
    } catch (e: any) {
      setError(String(e?.message ?? e));
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    if (open) {
      setHighlight(selectedId);
      load();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  const picked = envs.find(e => e.id === highlight);

  return (
    <Dialog open={open} onOpenChange={(_, d) => { if (!d.open) onClose(); }} modalType="modal">
      <DialogSurface style={{ maxWidth: '760px' }}>
        <DialogBody>
          <DialogTitle>Choose a Power Platform environment</DialogTitle>
          <DialogContent>
            <Text size={200} block style={{ color: tokens.colorNeutralForeground3, marginBottom: tokens.spacingVerticalS }}>
              The app validates each environment for the Dynamics 365 Knowledgebase
              (<code>knowledgearticle</code> table) so you don't accidentally load into an environment that doesn't support it.
            </Text>

            <div className={s.toolbar}>
              <Button
                icon={<ArrowClockwise20Regular />}
                appearance="subtle"
                onClick={load}
                disabled={loading}
              >
                Refresh
              </Button>
            </div>

            {error && (
              <MessageBar intent="error">
                <MessageBarBody>{error}</MessageBarBody>
              </MessageBar>
            )}

            <div className={s.list}>
              {loading && envs.length === 0 && (
                <div className={s.empty}><Spinner size="small" label="Loading environments…" /></div>
              )}
              {envs.map(env => {
                const isSel = highlight === env.id;
                return (
                  <div
                    key={env.id}
                    className={mergeClasses(s.row, isSel && s.rowActive)}
                    onClick={() => setHighlight(env.id)}
                    onDoubleClick={() => onPick(env)}
                  >
                    <span className={s.icon}><Cloud24Filled /></span>
                    <div className={s.meta}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalXS }}>
                        <Text weight="semibold" truncate block>{env.displayName}</Text>
                        {env.isDefault && (
                          <Badge appearance="tint" color="brand" size="small">Default</Badge>
                        )}
                        {env.region && (
                          <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>· {env.region}</Text>
                        )}
                      </div>
                      <Text className={s.url} truncate block>{env.url}</Text>
                      <KbStatusBadge status={env.knowledgebaseStatus} error={env.knowledgebaseError} />
                    </div>
                  </div>
                );
              })}
            </div>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose}>Cancel</Button>
            <Button
              appearance="primary"
              disabled={!picked || picked.knowledgebaseStatus === 'checking'}
              onClick={() => picked && onPick(picked)}
            >
              Use this environment
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

function KbStatusIcon({ status }: { status: PowerPlatformEnvironment['knowledgebaseStatus'] }) {
  if (status === 'present') return <CheckmarkCircle20Filled style={{ color: '#7FE0A3' }} />;
  if (status === 'missing') return <DismissCircle20Filled style={{ color: '#FFB3B3' }} />;
  if (status === 'error') return <Warning20Filled style={{ color: '#FFD479' }} />;
  if (status === 'checking') return <Spinner size="tiny" />;
  return null;
}

function KbStatusBadge({
  status, error,
}: { status: PowerPlatformEnvironment['knowledgebaseStatus']; error?: string }) {
  if (status === 'checking') {
    return (
      <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4, marginTop: 4 }}>
        <Spinner size="tiny" /> <Text size={200}>Checking knowledgebase…</Text>
      </span>
    );
  }
  if (status === 'present') {
    return <Badge appearance="tint" color="success" icon={<CheckmarkCircle20Filled />} style={{ marginTop: 4 }}>Knowledgebase available</Badge>;
  }
  if (status === 'missing') {
    return <Badge appearance="tint" color="danger" icon={<DismissCircle20Filled />} style={{ marginTop: 4 }}>Knowledgebase not installed</Badge>;
  }
  if (status === 'error') {
    return <Badge appearance="tint" color="warning" icon={<Warning20Filled />} style={{ marginTop: 4 }}>{error ?? 'Check failed'}</Badge>;
  }
  return null;
}
