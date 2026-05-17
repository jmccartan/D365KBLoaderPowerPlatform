import { useEffect, useState } from 'react';
import {
  Dialog, DialogSurface, DialogTitle, DialogBody, DialogActions, DialogContent,
  Button, Spinner, Text, makeStyles, tokens, MessageBar, MessageBarBody,
} from '@fluentui/react-components';
import { TagMultiple24Filled, ChevronRight20Regular, FolderArrowUp24Regular, Home20Regular } from '@fluentui/react-icons';
import type { KbLoaderService } from '../services/KbLoaderService';
import type { KbSubject } from '../types';

const useStyles = makeStyles({
  toolbar: { display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS, marginBottom: tokens.spacingVerticalS, flexWrap: 'wrap' },
  crumbs: {
    display: 'flex', alignItems: 'center', flexWrap: 'wrap',
    gap: tokens.spacingHorizontalXS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
    flex: 1, minWidth: '200px',
  },
  list: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS, maxHeight: '420px', overflow: 'auto', marginTop: tokens.spacingVerticalM },
  row: {
    display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalM,
    padding: tokens.spacingHorizontalM, borderRadius: tokens.borderRadiusMedium, cursor: 'pointer',
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  icon: {
    width: '36px', height: '36px', borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorBrandBackground2, color: tokens.colorBrandForeground1,
    display: 'inline-flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0,
  },
  empty: { padding: tokens.spacingHorizontalL, textAlign: 'center', color: tokens.colorNeutralForeground3 },
});

export interface BrowseSubjectDialogProps {
  open: boolean;
  service: KbLoaderService;
  onPick: (subject: KbSubject) => void;
  onClose: () => void;
}

export function BrowseSubjectDialog({ open, service, onPick, onClose }: BrowseSubjectDialogProps) {
  const s = useStyles();
  const [stack, setStack] = useState<KbSubject[]>([]);   // breadcrumb stack
  const [children, setChildren] = useState<KbSubject[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>();
  const [selected, setSelected] = useState<KbSubject | undefined>();

  async function load(parent: KbSubject | undefined) {
    setLoading(true);
    setError(undefined);
    try {
      setChildren(await service.listSubjects(parent?.id));
    } catch (e: any) {
      setError(String(e?.message ?? e));
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    if (open) {
      setStack([]);
      setSelected(undefined);
      load(undefined);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  const current = stack[stack.length - 1];

  function enter(subj: KbSubject) {
    setStack(prev => [...prev, subj]);
    setSelected(subj);
    load(subj);
  }
  function goUp() {
    setStack(prev => prev.slice(0, -1));
    setSelected(stack[stack.length - 2]);
    load(stack[stack.length - 2]);
  }
  function goRoot() {
    setStack([]);
    setSelected(undefined);
    load(undefined);
  }

  return (
    <Dialog open={open} onOpenChange={(_, d) => { if (!d.open) onClose(); }} modalType="modal">
      <DialogSurface style={{ maxWidth: '720px' }}>
        <DialogBody>
          <DialogTitle>Choose a subject (category)</DialogTitle>
          <DialogContent>
            <div className={s.toolbar}>
              <div className={s.crumbs}>
                <Button size="small" appearance="subtle" icon={<Home20Regular />} onClick={goRoot}>Root</Button>
                {stack.map((subj, i) => (
                  <span key={subj.id} style={{ display: 'inline-flex', alignItems: 'center' }}>
                    <ChevronRight20Regular style={{ color: tokens.colorNeutralForeground3 }} />
                    <Button size="small" appearance="subtle" onClick={() => {
                      setStack(stack.slice(0, i + 1));
                      setSelected(subj);
                      load(subj);
                    }}>{subj.name}</Button>
                  </span>
                ))}
              </div>
              <Button icon={<FolderArrowUp24Regular />} appearance="subtle" onClick={goUp} disabled={stack.length === 0}>Up</Button>
            </div>
            {error && <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar>}
            <div className={s.list}>
              {loading && <div className={s.empty}><Spinner size="small" label="Loading subjects…" /></div>}
              {!loading && children.length === 0 && (
                <div className={s.empty}>
                  No sub-categories.
                  {current && ' Click "Use this category" to pick this level.'}
                </div>
              )}
              {!loading && children.map(c => (
                <div key={c.id} className={s.row} onDoubleClick={() => onPick(c)} onClick={() => { setSelected(c); if (c.hasChildren !== false) enter(c); }}>
                  <span className={s.icon}><TagMultiple24Filled /></span>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <Text weight="semibold" truncate block>{c.name}</Text>
                    <Text size={200} truncate block style={{ color: tokens.colorNeutralForeground3 }}>{c.path}</Text>
                  </div>
                  <ChevronRight20Regular style={{ color: tokens.colorNeutralForeground3 }} />
                </div>
              ))}
            </div>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose}>Cancel</Button>
            <Button
              appearance="primary"
              disabled={!current && !selected}
              onClick={() => {
                const pick = selected ?? current;
                if (pick) onPick(pick);
              }}
            >
              Use this category
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
