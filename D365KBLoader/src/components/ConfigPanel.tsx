import { useEffect, useMemo, useState } from 'react';
import {
  Field, Input, Button, Card, Text, makeStyles, tokens, Spinner, MessageBar, MessageBarBody
} from '@fluentui/react-components';
import {
  FolderOpen24Filled, GlobeSearch24Regular, FolderLink24Regular, Search20Filled,
  Globe20Regular, FolderSearch20Regular,
} from '@fluentui/react-icons';
import type { KbConfig } from '../types';
import { getService } from '../services';
import { BrowseSiteDialog } from './BrowseSiteDialog';
import { BrowseFolderDialog } from './BrowseFolderDialog';

const useStyles = makeStyles({
  card: {
    maxWidth: '100%',
    flex: 1,
    minWidth: 0,
    padding: 0,
    overflow: 'hidden',
    boxShadow: tokens.shadow8,
    borderRadius: tokens.borderRadiusXLarge,
  },
  accent: {
    height: '4px',
    background: 'linear-gradient(90deg, #0B3A6F 0%, #1278D2 50%, #4F9DE8 100%)',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalL} ${tokens.spacingHorizontalXL} ${tokens.spacingVerticalM}`,
  },
  iconWrap: {
    width: '44px',
    height: '44px',
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
  },
  titleRow: { display: 'flex', flexDirection: 'column' },
  title: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  sub: { color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 },
  body: {
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalXL} ${tokens.spacingVerticalXL}`,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalL,
  },
  actions: {
    display: 'flex',
    gap: tokens.spacingHorizontalM,
    alignItems: 'center',
  },
  hint: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  inputRow: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    alignItems: 'stretch',
  },
});

export interface ConfigPanelProps {
  config: KbConfig;
  onChange: (cfg: KbConfig) => void;
  onScan: () => void;
  scanning: boolean;
  error?: string;
}

export function ConfigPanel({ config, onChange, onScan, scanning, error }: ConfigPanelProps) {
  const s = useStyles();
  const svc = useMemo(() => getService(), []);
  const [local, setLocal] = useState<KbConfig>(config);
  const [siteOpen, setSiteOpen] = useState(false);
  const [folderOpen, setFolderOpen] = useState(false);

  useEffect(() => {
    setLocal(config);
  }, [config]);

  const update = (patch: Partial<KbConfig>) => {
    const next = { ...local, ...patch };
    setLocal(next);
    onChange(next);
  };
  const ready = !!local.siteUrl && !!local.folderPath;

  return (
    <Card className={s.card}>
      <div className={s.accent} />
      <div className={s.header}>
        <span className={s.iconWrap}><FolderOpen24Filled /></span>
        <div className={s.titleRow}>
          <span className={s.title}>Point at the SharePoint source</span>
          <span className={s.sub}>Files in the folder will be classified, converted to HTML, and previewed for review.</span>
        </div>
      </div>
      <div className={s.body}>
        <Field label="SharePoint site URL" hint="Pick from your sites or paste a URL.">
          <div className={s.inputRow}>
            <Input
              value={local.siteUrl}
              onChange={(_, d) => update({ siteUrl: d.value })}
              placeholder="https://contoso.sharepoint.com/sites/Support"
              contentBefore={<GlobeSearch24Regular />}
              size="large"
              style={{ flex: 1 }}
            />
            <Button
              appearance="secondary"
              size="large"
              icon={<Globe20Regular />}
              onClick={() => setSiteOpen(true)}
            >
              Browse…
            </Button>
          </div>
        </Field>
        <Field label="Folder path" hint="Drill into the site's libraries to pick a folder. A run report (.xlsx) will be saved here after each load.">
          <div className={s.inputRow}>
            <Input
              value={local.folderPath}
              onChange={(_, d) => update({ folderPath: d.value })}
              placeholder="/Shared Documents/KB"
              contentBefore={<FolderLink24Regular />}
              size="large"
              style={{ flex: 1 }}
            />
            <Button
              appearance="secondary"
              size="large"
              icon={<FolderSearch20Regular />}
              onClick={() => setFolderOpen(true)}
              disabled={!local.siteUrl}
              title={!local.siteUrl ? 'Pick a site first' : 'Browse folders'}
            >
              Browse…
            </Button>
          </div>
        </Field>

        {error && (
          <MessageBar intent="error">
            <MessageBarBody>{error}</MessageBarBody>
          </MessageBar>
        )}

        <div className={s.actions}>
          <Button
            appearance="primary"
            size="large"
            icon={scanning ? <Spinner size="tiny" /> : <Search20Filled />}
            onClick={onScan}
            disabled={scanning || !ready}
          >
            {scanning ? 'Scanning…' : 'Scan folder'}
          </Button>
          {!ready && (
            <Text className={s.hint}>Fill in site URL and folder path to continue.</Text>
          )}
        </div>
      </div>

      <BrowseSiteDialog
        open={siteOpen}
        service={svc}
        onPick={url => update({ siteUrl: url, folderPath: '/' })}
        onClose={() => setSiteOpen(false)}
      />
      <BrowseFolderDialog
        open={folderOpen}
        service={svc}
        siteUrl={local.siteUrl}
        initialPath={local.folderPath || '/'}
        onPick={path => update({ folderPath: path })}
        onClose={() => setFolderOpen(false)}
      />
    </Card>
  );
}
