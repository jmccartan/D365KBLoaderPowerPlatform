import { useMemo, useState } from 'react';
import {
  Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle,
  Field, Input, Menu, MenuDivider, MenuItem, MenuList, MenuPopover, MenuTrigger,
  makeStyles, tokens, Text,
} from '@fluentui/react-components';
import { Delete20Regular, Save20Regular } from '@fluentui/react-icons';
import type { SavedScanProfile } from '../types';

const useStyles = makeStyles({
  wrap: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: tokens.spacingHorizontalM,
    flexWrap: 'wrap',
  },
  meta: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXXS,
  },
  title: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
  },
  sub: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  manageList: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    marginTop: tokens.spacingVerticalM,
  },
  manageRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: tokens.spacingHorizontalM,
    padding: tokens.spacingHorizontalM,
    borderTopWidth: '1px',
    borderRightWidth: '1px',
    borderBottomWidth: '1px',
    borderLeftWidth: '1px',
    borderTopStyle: 'solid',
    borderRightStyle: 'solid',
    borderBottomStyle: 'solid',
    borderLeftStyle: 'solid',
    borderTopColor: tokens.colorNeutralStroke2,
    borderRightColor: tokens.colorNeutralStroke2,
    borderBottomColor: tokens.colorNeutralStroke2,
    borderLeftColor: tokens.colorNeutralStroke2,
    borderRadius: tokens.borderRadiusLarge,
  },
});

export interface ProfilesBarProps {
  profiles: SavedScanProfile[];
  selectedProfileId?: string;
  loading?: boolean;
  onApply: (profile: SavedScanProfile) => void;
  onSave: (name: string) => Promise<void>;
  onDelete: (id: string) => Promise<void>;
}

export function ProfilesBar({ profiles, selectedProfileId, loading = false, onApply, onSave, onDelete }: ProfilesBarProps) {
  const s = useStyles();
  const [saveOpen, setSaveOpen] = useState(false);
  const [manageOpen, setManageOpen] = useState(false);
  const [saving, setSaving] = useState(false);
  const [deletingId, setDeletingId] = useState<string | undefined>();
  const [name, setName] = useState('');

  const selected = useMemo(
    () => profiles.find(profile => profile.id === selectedProfileId),
    [profiles, selectedProfileId],
  );

  async function handleSave() {
    if (!name.trim()) {
      return;
    }
    setSaving(true);
    try {
      await onSave(name.trim());
      setName('');
      setSaveOpen(false);
    } finally {
      setSaving(false);
    }
  }

  async function handleDelete(id: string) {
    setDeletingId(id);
    try {
      await onDelete(id);
    } finally {
      setDeletingId(undefined);
    }
  }

  return (
    <div className={s.wrap}>
      <div className={s.meta}>
        <span className={s.title}>Profiles</span>
        <Text className={s.sub}>
          {selected ? `Active: ${selected.name}` : 'Save a repeatable scan setup for demos, dry-runs, and production loads.'}
        </Text>
      </div>
      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <Button appearance="secondary" disabled={loading} icon={<Save20Regular />}>
            {selected ? `Profiles: ${selected.name}` : 'Profiles'}
          </Button>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            {profiles.length === 0 && <MenuItem disabled>No saved profiles yet</MenuItem>}
            {profiles.map(profile => (
              <MenuItem key={profile.id} onClick={() => onApply(profile)}>
                {profile.name}
              </MenuItem>
            ))}
            <MenuDivider />
            <MenuItem onClick={() => setSaveOpen(true)}>Save current as…</MenuItem>
            <MenuItem onClick={() => setManageOpen(true)}>Manage</MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>

      <Dialog open={saveOpen} onOpenChange={(_, data) => { if (!data.open) setSaveOpen(false); }}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Save current profile</DialogTitle>
            <DialogContent>
              <Field label="Profile name" required>
                <Input value={name} onChange={(_, data) => setName(data.value)} placeholder="Production – English KB" />
              </Field>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setSaveOpen(false)} disabled={saving}>Cancel</Button>
              <Button appearance="primary" onClick={handleSave} disabled={saving || !name.trim()}>
                {saving ? 'Saving…' : 'Save profile'}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      <Dialog open={manageOpen} onOpenChange={(_, data) => { if (!data.open) setManageOpen(false); }}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Manage saved profiles</DialogTitle>
            <DialogContent>
              <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                Delete stale profiles here. Use the dropdown menu to apply one.
              </Text>
              <div className={s.manageList}>
                {profiles.length === 0 && <Text>No saved profiles yet.</Text>}
                {profiles.map(profile => (
                  <div key={profile.id} className={s.manageRow}>
                    <div>
                      <Text weight="semibold" block>{profile.name}</Text>
                      <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                        {profile.environmentId ? `Environment: ${profile.environmentId}` : 'No environment saved'}
                      </Text>
                    </div>
                    <Button
                      appearance="subtle"
                      icon={<Delete20Regular />}
                      onClick={() => handleDelete(profile.id)}
                      disabled={deletingId === profile.id}
                    >
                      {deletingId === profile.id ? 'Deleting…' : 'Delete'}
                    </Button>
                  </div>
                ))}
              </div>
            </DialogContent>
            <DialogActions>
              <Button appearance="primary" onClick={() => setManageOpen(false)}>Done</Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
}
