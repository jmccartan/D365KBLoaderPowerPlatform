import { useState } from 'react';
import {
  Dialog, DialogSurface, DialogTitle, DialogBody, DialogActions, DialogContent,
  Button, Text, makeStyles, tokens, Tab, TabList, Textarea, Spinner, MessageBar, MessageBarBody, MessageBarTitle, Badge,
} from '@fluentui/react-components';
import {
  Sparkle24Filled, CheckmarkCircle20Filled, DismissCircle20Filled, Eye20Regular, Code20Regular, ArrowSync20Regular,
} from '@fluentui/react-icons';
import type { ArticleSuggestion } from '../types';

const useStyles = makeStyles({
  surface: {
    maxWidth: '900px',
    width: '95vw',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
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
  },
  loading: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingHorizontalXXL,
  },
  changes: {
    margin: 0,
    paddingLeft: tokens.spacingHorizontalL,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
  },
  preview: {
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
    padding: tokens.spacingHorizontalL,
    maxHeight: '420px',
    overflow: 'auto',
    backgroundColor: tokens.colorNeutralBackground1,
    marginTop: tokens.spacingVerticalS,
  },
  raw: {
    fontFamily: 'Consolas, Cascadia Code, monospace',
    fontSize: '12.5px',
    minHeight: '320px',
    marginTop: tokens.spacingVerticalS,
  },
  proposedTitle: {
    padding: tokens.spacingHorizontalM,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
    fontWeight: tokens.fontWeightSemibold,
    marginTop: tokens.spacingVerticalS,
  },
});

export interface SuggestEditsDialogProps {
  open: boolean;
  loading: boolean;
  error?: string;
  suggestion?: ArticleSuggestion;
  currentTitle: string;
  onAccept: (suggestion: ArticleSuggestion) => void;
  onDecline: () => void;
  onRegenerate: () => void;
}

export function SuggestEditsDialog({
  open, loading, error, suggestion, currentTitle, onAccept, onDecline, onRegenerate,
}: SuggestEditsDialogProps) {
  const s = useStyles();
  const [tab, setTab] = useState<'preview' | 'html'>('preview');

  const titleChanged = suggestion?.title && suggestion.title !== currentTitle;

  return (
    <Dialog open={open} onOpenChange={(_, d) => { if (!d.open) onDecline(); }} modalType="modal">
      <DialogSurface className={s.surface}>
        <DialogBody>
          <DialogTitle>
            <div className={s.header}>
              <span className={s.icon}><Sparkle24Filled /></span>
              <span>Copilot suggestions</span>
              {suggestion && suggestion.changes.length > 0 && (
                <Badge appearance="tint" color="brand">
                  {suggestion.changes.length} change{suggestion.changes.length === 1 ? '' : 's'}
                </Badge>
              )}
            </div>
          </DialogTitle>
          <DialogContent>
            {loading && (
              <div className={s.loading}>
                <Spinner size="large" />
                <Text>Reviewing the article…</Text>
              </div>
            )}

            {!loading && error && (
              <MessageBar intent="error">
                <MessageBarBody>
                  <MessageBarTitle>Could not generate suggestions</MessageBarTitle>
                  {error}
                </MessageBarBody>
              </MessageBar>
            )}

            {!loading && !error && suggestion && (
              <>
                <MessageBar intent={suggestion.changes.length === 0 ? 'success' : 'info'}>
                  <MessageBarBody>
                    <MessageBarTitle>Summary</MessageBarTitle>
                    {suggestion.summary}
                  </MessageBarBody>
                </MessageBar>

                {titleChanged && (
                  <div className={s.proposedTitle}>
                    Proposed title: {suggestion.title}
                  </div>
                )}

                {suggestion.changes.length > 0 && (
                  <div style={{ marginTop: tokens.spacingVerticalM }}>
                    <Text weight="semibold" block style={{ marginBottom: tokens.spacingVerticalXS }}>
                      Proposed changes
                    </Text>
                    <ul className={s.changes}>
                      {suggestion.changes.map((c, i) => (
                        <li key={i}><Text size={300}>{c}</Text></li>
                      ))}
                    </ul>
                  </div>
                )}

                <div style={{ marginTop: tokens.spacingVerticalL }}>
                  <TabList
                    selectedValue={tab}
                    onTabSelect={(_, d) => setTab(d.value as 'preview' | 'html')}
                    appearance="subtle-circular"
                  >
                    <Tab icon={<Eye20Regular />} value="preview">Preview</Tab>
                    <Tab icon={<Code20Regular />} value="html">Proposed HTML</Tab>
                  </TabList>
                  {tab === 'preview' && (
                    <div className={s.preview} dangerouslySetInnerHTML={{ __html: suggestion.html }} />
                  )}
                  {tab === 'html' && (
                    <Textarea
                      value={suggestion.html}
                      readOnly
                      resize="vertical"
                      rows={16}
                      className={s.raw}
                    />
                  )}
                </div>
              </>
            )}
          </DialogContent>
          <DialogActions>
            <Button
              appearance="subtle"
              icon={<ArrowSync20Regular />}
              onClick={onRegenerate}
              disabled={loading}
            >
              Regenerate
            </Button>
            <Button
              appearance="secondary"
              icon={<DismissCircle20Filled />}
              onClick={onDecline}
              disabled={loading}
            >
              Decline
            </Button>
            <Button
              appearance="primary"
              icon={<CheckmarkCircle20Filled />}
              onClick={() => suggestion && onAccept(suggestion)}
              disabled={loading || !suggestion}
            >
              Accept changes
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
