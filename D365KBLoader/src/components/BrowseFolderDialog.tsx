import { useEffect, useState } from 'react';
import {
  Dialog, DialogSurface, DialogTitle, DialogBody, DialogActions, DialogContent,
  Button, Spinner, Text, makeStyles, tokens, MessageBar, MessageBarBody,
} from '@fluentui/react-components';
import {
  Folder24Filled, FolderArrowUp24Regular, ChevronRight20Regular, Home20Regular, ArrowClockwise20Regular,
} from '@fluentui/react-icons';
import type { KbLoaderService } from '../services/KbLoaderService';
import type { FolderItem } from '../types';

const useStyles = makeStyles({
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    marginBottom: tokens.spacingVerticalS,
    flexWrap: 'wrap',
  },
  crumbs: {
    display: 'flex',
    alignItems: 'center',
    flexWrap: 'wrap',
    gap: tokens.spacingHorizontalXS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    flex: 1,
    minWidth: '200px',
  },
  crumbButton: {
    padding: `2px ${tokens.spacingHorizontalXS}`,
    minWidth: 'auto',
  },
  sep: { color: tokens.colorNeutralForeground3, display: 'inline-flex', alignItems: 'center' },
  list: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    maxHeight: '420px',
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
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
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
  chev: { color: tokens.colorNeutralForeground3 },
  meta: { display: 'flex', flexDirection: 'column', overflow: 'hidden', flex: 1 },
  empty: { padding: tokens.spacingHorizontalL, textAlign: 'center', color: tokens.colorNeutralForeground3 },
});

export interface BrowseFolderDialogProps {
  open: boolean;
  service: KbLoaderService;
  siteUrl: string;
  initialPath?: string;
  onPick: (path: string) => void;
  onClose: () => void;
}

export function BrowseFolderDialog({ open, service, siteUrl, initialPath, onPick, onClose }: BrowseFolderDialogProps) {
  const s = useStyles();
  const [path, setPath] = useState<string>(normalize(initialPath ?? '/'));
  const [folders, setFolders] = useState<FolderItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>();

  async function load(p: string) {
    setLoading(true);
    setError(undefined);
    try {
      setFolders(await service.listFolders(siteUrl, p));
    } catch (e: any) {
      setError(String(e?.message ?? e));
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    if (open) {
      const start = normalize(initialPath ?? '/');
      setPath(start);
      load(start);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open, siteUrl]);

  function go(p: string) {
    const next = normalize(p);
    setPath(next);
    load(next);
  }

  function goUp() {
    if (path === '/' || !path) return;
    const idx = path.lastIndexOf('/');
    go(idx <= 0 ? '/' : path.slice(0, idx));
  }

  const segments = path === '/' ? [] : path.split('/').filter(Boolean);

  return (
    <Dialog open={open} onOpenChange={(_, d) => { if (!d.open) onClose(); }} modalType="modal">
      <DialogSurface style={{ maxWidth: '760px' }}>
        <DialogBody>
          <DialogTitle>Browse folder</DialogTitle>
          <DialogContent>
            {!siteUrl && (
              <MessageBar intent="warning">
                <MessageBarBody>Pick a SharePoint site first.</MessageBarBody>
              </MessageBar>
            )}
            <div className={s.toolbar}>
              <div className={s.crumbs} aria-label="Breadcrumbs">
                <Button
                  size="small"
                  appearance="subtle"
                  icon={<Home20Regular />}
                  className={s.crumbButton}
                  onClick={() => go('/')}
                >
                  Root
                </Button>
                {segments.map((seg, i) => {
                  const sub = '/' + segments.slice(0, i + 1).join('/');
                  return (
                    <span key={sub} style={{ display: 'inline-flex', alignItems: 'center' }}>
                      <span className={s.sep}><ChevronRight20Regular /></span>
                      <Button
                        size="small"
                        appearance="subtle"
                        className={s.crumbButton}
                        onClick={() => go(sub)}
                      >
                        {seg}
                      </Button>
                    </span>
                  );
                })}
              </div>
              <Button icon={<FolderArrowUp24Regular />} appearance="subtle" onClick={goUp} disabled={path === '/'}>
                Up
              </Button>
              <Button icon={<ArrowClockwise20Regular />} appearance="subtle" onClick={() => load(path)} disabled={loading}>
                Refresh
              </Button>
            </div>

            {error && (
              <MessageBar intent="error" style={{ marginTop: 12 }}>
                <MessageBarBody>{error}</MessageBarBody>
              </MessageBar>
            )}

            <div className={s.list}>
              {loading && (
                <div className={s.empty}><Spinner size="small" label="Loading folders…" /></div>
              )}
              {!loading && folders.length === 0 && (
                <div className={s.empty}>No sub-folders here. Click "Use this folder" to pick this location.</div>
              )}
              {!loading && folders.map(f => (
                <div key={f.path} className={s.row} onClick={() => go(f.path)} onDoubleClick={() => { onPick(f.path); onClose(); }}>
                  <span className={s.icon}><Folder24Filled /></span>
                  <div className={s.meta}>
                    <Text weight="semibold" truncate block>{f.name}</Text>
                    <Text size={200} truncate block style={{ color: tokens.colorNeutralForeground3 }}>{f.path}</Text>
                  </div>
                  <ChevronRight20Regular className={s.chev} />
                </div>
              ))}
            </div>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose}>Cancel</Button>
            <Button appearance="primary" disabled={!siteUrl} onClick={() => { onPick(path); onClose(); }}>
              Use this folder
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

function normalize(p: string): string {
  if (!p || p.trim() === '') return '/';
  let n = p.trim();
  if (!n.startsWith('/')) n = '/' + n;
  if (n.length > 1 && n.endsWith('/')) n = n.slice(0, -1);
  return n;
}
