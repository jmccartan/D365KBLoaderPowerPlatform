import { useEffect, useMemo, useState } from 'react';
import {
  Dialog, DialogSurface, DialogTitle, DialogBody, DialogActions, DialogContent,
  Button, Spinner, Text, Input, makeStyles, tokens, mergeClasses, MessageBar, MessageBarBody, Divider,
} from '@fluentui/react-components';
import { Globe24Filled, Search20Regular, ArrowClockwise20Regular, ArrowRight20Regular } from '@fluentui/react-icons';
import type { KbLoaderService } from '../services/KbLoaderService';
import type { SharePointSite } from '../types';

const RECENT_SITES_KEY = 'd365kb.recentSites';
const MAX_RECENTS = 6;

function readRecents(): string[] {
  try {
    const raw = localStorage.getItem(RECENT_SITES_KEY);
    if (!raw) return [];
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr.filter(x => typeof x === 'string') : [];
  } catch { return []; }
}

function pushRecent(url: string) {
  try {
    const cur = readRecents().filter(u => u !== url);
    cur.unshift(url);
    localStorage.setItem(RECENT_SITES_KEY, JSON.stringify(cur.slice(0, MAX_RECENTS)));
  } catch { /* noop */ }
}

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
  const [manualUrl, setManualUrl] = useState('');
  const [recents, setRecents] = useState<string[]>([]);

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
      setManualUrl('');
      setRecents(readRecents());
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

  const manualValid = useMemo(() => /^https:\/\/.+\.sharepoint\.com\/.+/i.test(manualUrl.trim()), [manualUrl]);
  function commitPick(url: string) {
    pushRecent(url);
    onPick(url);
    onClose();
  }

  return (
    <Dialog open={open} onOpenChange={(_, d) => { if (!d.open) onClose(); }} modalType="modal">
      <DialogSurface style={{ maxWidth: '720px' }}>
        <DialogBody>
          <DialogTitle>Pick a SharePoint site</DialogTitle>
          <DialogContent>
            <Text size={200} block style={{ color: tokens.colorNeutralForeground3, marginBottom: tokens.spacingVerticalS }}>
              Paste any SharePoint Online site URL — or pick from sites discovered via the connector below.
            </Text>

            <div className={s.toolbar}>
              <Input
                placeholder="https://contoso.sharepoint.com/sites/Support"
                value={manualUrl}
                onChange={(_, d) => setManualUrl(d.value)}
                onKeyDown={e => { if (e.key === 'Enter' && manualValid) commitPick(manualUrl.trim()); }}
                style={{ flex: 1 }}
              />
              <Button
                appearance="primary"
                icon={<ArrowRight20Regular />}
                disabled={!manualValid}
                onClick={() => commitPick(manualUrl.trim())}
              >
                Use URL
              </Button>
            </div>

            {recents.length > 0 && (
              <div style={{ marginTop: tokens.spacingVerticalS, display: 'flex', flexWrap: 'wrap', gap: tokens.spacingHorizontalXS, alignItems: 'center' }}>
                <Text size={200} style={{ color: tokens.colorNeutralForeground3, marginRight: tokens.spacingHorizontalXS }}>Recent:</Text>
                {recents.map(url => (
                  <Button
                    key={url}
                    size="small"
                    appearance="subtle"
                    onClick={() => commitPick(url)}
                    title={url}
                  >
                    {url.replace(/^https:\/\//, '').replace(/\/$/, '')}
                  </Button>
                ))}
              </div>
            )}

            <Divider style={{ marginTop: tokens.spacingVerticalM, marginBottom: tokens.spacingVerticalS }}>or browse</Divider>

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
            {!loading && !error && sites.length > 0 && sites[0].id?.startsWith('demo-site') && (
              <MessageBar intent="info" style={{ marginTop: 12 }}>
                <MessageBarBody>
                  Site discovery via the SharePoint connector isn't wired yet, so this list shows sample sites.
                  Paste your real site URL above (e.g. <code>https://yourtenant.sharepoint.com/sites/YourSite</code>) — it'll be remembered for next time.
                  KB writes to Dataverse are live; for real document ingest today, use the "Upload local files" drop zone in Configure.
                </MessageBarBody>
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
                    onDoubleClick={() => commitPick(site.url)}
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
              onClick={() => { if (selected) commitPick(selected); }}
            >
              Use this site
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
