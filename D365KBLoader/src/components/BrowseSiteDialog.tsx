import { useEffect, useState } from 'react';
import {
  Dialog, DialogSurface, DialogTitle, DialogBody, DialogActions, DialogContent,
  Button, Spinner, Text, Input, makeStyles, tokens, mergeClasses, MessageBar, MessageBarBody,
} from '@fluentui/react-components';
import { Globe24Filled, Search20Regular, ArrowClockwise20Regular } from '@fluentui/react-icons';
import type { KbLoaderService } from '../services/KbLoaderService';
import type { SharePointSite } from '../types';

const useStyles = makeStyles({
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
  toolbar: { display: 'flex', gap: tokens.spacingHorizontalS, alignItems: 'center' },
});

export interface BrowseSiteDialogProps {
  open: boolean;
  service: KbLoaderService;
  onPick: (url: string) => void;
  onClose: () => void;
}

export function BrowseSiteDialog({ open, service, onPick, onClose }: BrowseSiteDialogProps) {
  const s = useStyles();
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>();
  const [sites, setSites] = useState<SharePointSite[]>([]);
  const [filter, setFilter] = useState('');
  const [selected, setSelected] = useState<string | undefined>();

  async function load() {
    setLoading(true);
    setError(undefined);
    try {
      setSites(await service.listSites());
    } catch (e: any) {
      setError(String(e?.message ?? e));
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    if (open) {
      setSelected(undefined);
      setFilter('');
      load();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  const visible = sites.filter(site => {
    const q = filter.trim().toLowerCase();
    if (!q) return true;
    return site.name.toLowerCase().includes(q)
      || site.url.toLowerCase().includes(q)
      || (site.description ?? '').toLowerCase().includes(q);
  });

  return (
    <Dialog open={open} onOpenChange={(_, d) => { if (!d.open) onClose(); }} modalType="modal">
      <DialogSurface style={{ maxWidth: '720px' }}>
        <DialogBody>
          <DialogTitle>Pick a SharePoint site</DialogTitle>
          <DialogContent>
            <div className={s.toolbar}>
              <Input
                placeholder="Filter by name, URL, description…"
                value={filter}
                onChange={(_, d) => setFilter(d.value)}
                contentBefore={<Search20Regular />}
                style={{ flex: 1 }}
              />
              <Button icon={<ArrowClockwise20Regular />} appearance="subtle" onClick={load} disabled={loading}>
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
                <div className={s.empty}><Spinner size="small" label="Loading sites…" /></div>
              )}
              {!loading && visible.length === 0 && (
                <div className={s.empty}>No matching sites.</div>
              )}
              {!loading && visible.map(site => {
                const isSel = selected === site.url;
                return (
                  <div
                    key={site.id}
                    className={mergeClasses(s.row, isSel && s.rowActive)}
                    onClick={() => setSelected(site.url)}
                    onDoubleClick={() => { onPick(site.url); onClose(); }}
                  >
                    <span className={s.icon}><Globe24Filled /></span>
                    <div className={s.meta}>
                      <Text weight="semibold" truncate block>{site.name}</Text>
                      <Text className={s.url} truncate block>{site.url}</Text>
                      {site.description && (
                        <Text size={200} truncate block style={{ color: tokens.colorNeutralForeground3 }}>
                          {site.description}
                        </Text>
                      )}
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
              disabled={!selected}
              onClick={() => { if (selected) { onPick(selected); onClose(); } }}
            >
              Use this site
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
