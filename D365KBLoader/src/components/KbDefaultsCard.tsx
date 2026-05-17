import { useEffect, useState } from 'react';
import {
  Card, Field, Dropdown, Option, Switch, makeStyles, tokens, Text, Spinner, Button,
} from '@fluentui/react-components';
import { LocalLanguage24Filled, TagMultiple24Regular, Edit24Regular, FolderSearch20Regular } from '@fluentui/react-icons';
import type { KbConfig, KbLanguage, KbSubject, DuplicateAction } from '../types';
import type { KbLoaderService } from '../services/KbLoaderService';
import { BrowseSubjectDialog } from './BrowseSubjectDialog';

const useStyles = makeStyles({
  card: {
    maxWidth: '760px',
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
  title: { fontSize: tokens.fontSizeBase500, fontWeight: tokens.fontWeightSemibold },
  sub: { color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 },
  body: {
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalXL} ${tokens.spacingVerticalXL}`,
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: tokens.spacingVerticalL,
    rowGap: tokens.spacingVerticalL,
  },
  full: { gridColumn: '1 / span 2' },
  subjRow: { display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS },
});

export interface KbDefaultsCardProps {
  service: KbLoaderService;
  config: KbConfig;
  onChange: (c: KbConfig) => void;
}

export function KbDefaultsCard({ service, config, onChange }: KbDefaultsCardProps) {
  const s = useStyles();
  const [langs, setLangs] = useState<KbLanguage[]>([]);
  const [langsLoading, setLangsLoading] = useState(false);
  const [subjectOpen, setSubjectOpen] = useState(false);
  const [selectedSubject, setSelectedSubject] = useState<KbSubject | undefined>();

  useEffect(() => {
    let alive = true;
    setLangsLoading(true);
    service.listLanguages().then(l => { if (alive) setLangs(l); }).catch(() => undefined).finally(() => setLangsLoading(false));
    return () => { alive = false; };
  }, [service]);

  function update(patch: Partial<KbConfig>) {
    onChange({ ...config, ...patch });
  }

  const langName = langs.find(l => l.id === config.defaultLanguageId)?.displayName;

  return (
    <Card className={s.card}>
      <div className={s.accent} />
      <div className={s.header}>
        <span className={s.iconWrap}><Edit24Regular /></span>
        <div>
          <div className={s.title}>Article defaults</div>
          <div className={s.sub}>Applied to every article you load — per-article overrides are available in Review.</div>
        </div>
      </div>
      <div className={s.body}>
        <Field label="Default language" hint="Sets languagelocaleid on each new article.">
          {langsLoading ? <Spinner size="tiny" /> : (
            <Dropdown
              value={langName ?? ''}
              selectedOptions={config.defaultLanguageId ? [config.defaultLanguageId] : []}
              onOptionSelect={(_, d) => update({ defaultLanguageId: d.optionValue })}
              placeholder="Select a language"
            >
              {langs.map(l => (
                <Option key={l.id} value={l.id} text={l.displayName}>
                  <LocalLanguage24Filled style={{ marginRight: 8 }} />
                  {l.displayName} ({l.code})
                </Option>
              ))}
            </Dropdown>
          )}
        </Field>

        <Field label="Default subject / category" hint="Browse the subject tree from Dataverse.">
          <div className={s.subjRow}>
            <Text style={{ flex: 1 }}>
              {selectedSubject ? selectedSubject.path : config.defaultSubjectId ? `(id: ${config.defaultSubjectId})` : <em style={{ color: tokens.colorNeutralForeground3 }}>none</em>}
            </Text>
            <Button icon={<FolderSearch20Regular />} onClick={() => setSubjectOpen(true)}>Browse…</Button>
          </div>
        </Field>

        <Field label="Publish on load" hint="Off = create as Draft for reviewer approval; on = publish immediately.">
          <Switch
            checked={!!config.publishOnLoad}
            label={config.publishOnLoad ? 'Publish immediately' : 'Save as Draft'}
            onChange={(_, d) => update({ publishOnLoad: d.checked })}
          />
        </Field>

        <Field label="If a matching article already exists" hint="Match is by exact title in the target KB.">
          <Dropdown
            value={duplicateLabel(config.duplicateAction)}
            selectedOptions={[config.duplicateAction ?? 'skip']}
            onOptionSelect={(_, d) => update({ duplicateAction: d.optionValue as DuplicateAction })}
          >
            <Option value="skip" text="Skip — leave the existing article">Skip — leave the existing article</Option>
            <Option value="update-existing" text="Update existing — overwrite content and title">Update existing — overwrite content and title</Option>
            <Option value="create-new" text="Create new (allow duplicates)">Create new (allow duplicates)</Option>
          </Dropdown>
        </Field>

        <Field className={s.full} label="Source-folder behavior" hint="Recurse into sub-folders and skip files that succeeded in a prior run.">
          <div style={{ display: 'flex', gap: tokens.spacingHorizontalL, flexWrap: 'wrap' }}>
            <Switch
              checked={!!config.recursive}
              label="Include sub-folders"
              onChange={(_, d) => update({ recursive: d.checked })}
            />
            <Switch
              checked={!!config.incremental}
              label="Incremental (skip files loaded successfully before)"
              onChange={(_, d) => update({ incremental: d.checked })}
            />
          </div>
        </Field>
      </div>

      <BrowseSubjectDialog
        open={subjectOpen}
        service={service}
        onClose={() => setSubjectOpen(false)}
        onPick={subj => {
          setSelectedSubject(subj);
          update({ defaultSubjectId: subj.id });
          setSubjectOpen(false);
        }}
      />
    </Card>
  );
}

function duplicateLabel(a?: DuplicateAction): string {
  switch (a) {
    case 'update-existing': return 'Update existing — overwrite content and title';
    case 'create-new': return 'Create new (allow duplicates)';
    default: return 'Skip — leave the existing article';
  }
}
