import { useState } from 'react';
import {
  Field, Input, Button, Card, CardHeader, Text, makeStyles, tokens, Spinner, MessageBar, MessageBarBody
} from '@fluentui/react-components';
import { FolderOpen24Regular } from '@fluentui/react-icons';
import type { KbConfig } from '../types';

const useStyles = makeStyles({
  card: { padding: tokens.spacingHorizontalL, maxWidth: '720px' },
  row: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM, marginTop: tokens.spacingVerticalM },
  actions: { display: 'flex', gap: tokens.spacingHorizontalM, marginTop: tokens.spacingVerticalL }
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
  const [local, setLocal] = useState<KbConfig>(config);
  const update = (patch: Partial<KbConfig>) => {
    const next = { ...local, ...patch };
    setLocal(next);
    onChange(next);
  };
  return (
    <Card className={s.card}>
      <CardHeader
        image={<FolderOpen24Regular />}
        header={<Text weight="semibold" size={500}>1. Pick the SharePoint source</Text>}
        description={<Text size={200}>Files in the folder will be classified, converted to HTML, and previewed for review.</Text>}
      />
      <div className={s.row}>
        <Field label="SharePoint site URL" hint="Example: https://contoso.sharepoint.com/sites/Support">
          <Input value={local.siteUrl} onChange={(_, d) => update({ siteUrl: d.value })} placeholder="https://contoso.sharepoint.com/sites/Support" />
        </Field>
        <Field label="Folder path" hint="Server-relative path inside the site (e.g. /Shared Documents/KB)">
          <Input value={local.folderPath} onChange={(_, d) => update({ folderPath: d.value })} placeholder="/Shared Documents/KB" />
        </Field>
        <Field label="Activity log list (SharePoint list name)" hint="A list in the same site used to record every action.">
          <Input value={local.logListName} onChange={(_, d) => update({ logListName: d.value })} placeholder="KB Loader Log" />
        </Field>
      </div>
      {error && <MessageBar intent="error" style={{ marginTop: 12 }}><MessageBarBody>{error}</MessageBarBody></MessageBar>}
      <div className={s.actions}>
        <Button appearance="primary" onClick={onScan} disabled={scanning || !local.siteUrl || !local.folderPath}>
          {scanning ? <Spinner size="tiny" /> : 'Scan folder'}
        </Button>
      </div>
    </Card>
  );
}
