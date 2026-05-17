import { useMemo, useState } from 'react';
import {
  Card, Text, ProgressBar, makeStyles, tokens,
  Table, TableHeader, TableRow, TableHeaderCell, TableBody, TableCell, Badge, Button, Divider, Spinner,
  MessageBar, MessageBarBody, MessageBarTitle, MessageBarActions,
  Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input, Textarea,
} from '@fluentui/react-components';
import {
  DocumentTable20Filled, DocumentBulletListMultiple24Filled,
  CheckmarkCircle24Filled, DismissCircle24Filled, History24Regular, Mail20Regular,
} from '@fluentui/react-icons';
import type { LogEntry, ReportResult } from '../types';

const useStyles = makeStyles({
  wrap: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalL },
  card: {
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
    padding: `${tokens.spacingVerticalL} ${tokens.spacingHorizontalXL}`,
    flexWrap: 'wrap',
  },
  headerText: { display: 'flex', flexDirection: 'column', flex: 1 },
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
  stats: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))',
    gap: tokens.spacingHorizontalL,
  },
  stat: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: tokens.spacingHorizontalL,
    borderRadius: tokens.borderRadiusLarge,
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
    backgroundColor: tokens.colorNeutralBackground1,
  },
  statIcon: {
    width: '44px',
    height: '44px',
    borderRadius: tokens.borderRadiusLarge,
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: 0,
  },
  statBrand: {
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
  },
  statGreen: {
    backgroundColor: tokens.colorPaletteGreenBackground2,
    color: tokens.colorPaletteGreenForeground1,
  },
  statRed: {
    backgroundColor: tokens.colorPaletteRedBackground2,
    color: tokens.colorPaletteRedForeground1,
  },
  statValue: {
    fontSize: '28px',
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: '32px',
    color: tokens.colorNeutralForeground1,
  },
  statLabel: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  progressWrap: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
  },
  progressTop: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
  },
  actionRow: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    flexWrap: 'wrap',
  },
});

interface EmailStatus {
  kind: 'success' | 'error';
  message: string;
}

export interface ProgressPanelProps {
  done: number;
  total: number;
  errors: number;
  log: LogEntry[];
  onSaveReport: () => void;
  reportSaving: boolean;
  reportResult?: ReportResult;
  reportError?: string;
  canEmailReport?: boolean;
  emailSending?: boolean;
  emailStatus?: EmailStatus;
  onEmailReport?: (to: string[], subject: string, html: string) => Promise<void>;
}

export function ProgressPanel({
  done, total, errors, log,
  onSaveReport, reportSaving, reportResult, reportError,
  canEmailReport, emailSending = false, emailStatus, onEmailReport,
}: ProgressPanelProps) {
  const s = useStyles();
  const pct = total === 0 ? 0 : done / total;
  const successes = Math.max(0, done - errors);
  const [emailOpen, setEmailOpen] = useState(false);
  const [to, setTo] = useState('');
  const [subject, setSubject] = useState('');
  const [message, setMessage] = useState('');

  const defaultSubject = useMemo(
    () => `KB Loader Run Report — ${new Date().toLocaleDateString()}`,
    [],
  );
  const defaultMessage = useMemo(
    () => `<p>The KB Loader run has completed.</p><ul><li><strong>Total selected:</strong> ${total}</li><li><strong>Completed:</strong> ${done}</li><li><strong>Loaded successfully:</strong> ${successes}</li><li><strong>Errors:</strong> ${errors}</li></ul><p>${reportResult ? `Latest report: <strong>${reportResult.fileName}</strong>.</p>` : 'The attached workbook contains the full activity log.</p>'}`,
    [done, errors, reportResult, successes, total],
  );

  function openEmailDialog() {
    setSubject(defaultSubject);
    setMessage(defaultMessage);
    setEmailOpen(true);
  }

  async function sendEmail() {
    if (!onEmailReport) {
      return;
    }
    const recipients = to.split(',').map(value => value.trim()).filter(Boolean);
    await onEmailReport(recipients, subject.trim(), message);
    setEmailOpen(false);
    setTo('');
  }

  return (
    <div className={s.wrap}>
      <Card className={s.card}>
        <div className={s.accent} />
        <div className={s.header}>
          <div className={s.headerText}>
            <span className={s.title}>Load progress</span>
            <span className={s.sub}>{done} of {total} processed</span>
          </div>
        </div>
        <div className={s.body}>
          <div className={s.stats}>
            <div className={s.stat}>
              <span className={`${s.statIcon} ${s.statBrand}`}><DocumentBulletListMultiple24Filled /></span>
              <div>
                <div className={s.statValue}>{total}</div>
                <div className={s.statLabel}>Selected for load</div>
              </div>
            </div>
            <div className={s.stat}>
              <span className={`${s.statIcon} ${s.statGreen}`}><CheckmarkCircle24Filled /></span>
              <div>
                <div className={s.statValue}>{successes}</div>
                <div className={s.statLabel}>Successfully loaded</div>
              </div>
            </div>
            <div className={s.stat}>
              <span className={`${s.statIcon} ${s.statRed}`}><DismissCircle24Filled /></span>
              <div>
                <div className={s.statValue}>{errors}</div>
                <div className={s.statLabel}>Errors</div>
              </div>
            </div>
          </div>
          <div className={s.progressWrap}>
            <div className={s.progressTop}>
              <Text weight="semibold">Overall progress</Text>
              <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                {Math.round(pct * 100)}%
              </Text>
            </div>
            <ProgressBar value={pct} thickness="large" />
          </div>
        </div>
      </Card>

      <Card className={s.card}>
        <div className={s.accent} />
        <div className={s.header}>
          <History24Regular style={{ color: tokens.colorBrandForeground1 }} />
          <div className={s.headerText}>
            <span className={s.title}>Activity log</span>
            <span className={s.sub}>
              {log.length} entries · auto-saved as an Excel report after each load
            </span>
          </div>
          <div className={s.actionRow}>
            <Button
              appearance="primary"
              icon={reportSaving ? <Spinner size="tiny" /> : <DocumentTable20Filled />}
              onClick={onSaveReport}
              disabled={reportSaving || log.length === 0}
            >
              {reportSaving ? 'Saving…' : 'Save Excel report'}
            </Button>
            <Button
              appearance="secondary"
              icon={emailSending ? <Spinner size="tiny" /> : <Mail20Regular />}
              onClick={openEmailDialog}
              disabled={!canEmailReport || log.length === 0 || emailSending}
            >
              {emailSending ? 'Sending…' : 'Email report…'}
            </Button>
          </div>
        </div>
        {reportResult && (
          <MessageBar intent="success" style={{ margin: `0 ${tokens.spacingHorizontalXL} ${tokens.spacingVerticalM}` }}>
            <MessageBarBody>
              <MessageBarTitle>Report saved</MessageBarTitle>
              <code>{reportResult.fileName}</code> {reportResult.downloaded ? 'downloaded to your browser' : `→ ${reportResult.location}`}
            </MessageBarBody>
          </MessageBar>
        )}
        {emailStatus && (
          <MessageBar intent={emailStatus.kind === 'success' ? 'success' : 'error'} style={{ margin: `0 ${tokens.spacingHorizontalXL} ${tokens.spacingVerticalM}` }}>
            <MessageBarBody>
              <MessageBarTitle>{emailStatus.kind === 'success' ? 'Email sent' : 'Could not send email'}</MessageBarTitle>
              {emailStatus.message}
            </MessageBarBody>
          </MessageBar>
        )}
        {reportError && (
          <MessageBar intent="error" style={{ margin: `0 ${tokens.spacingHorizontalXL} ${tokens.spacingVerticalM}` }}>
            <MessageBarBody>
              <MessageBarTitle>Could not save report</MessageBarTitle>
              {reportError}
            </MessageBarBody>
            <MessageBarActions>
              <Button appearance="transparent" onClick={onSaveReport}>Retry</Button>
            </MessageBarActions>
          </MessageBar>
        )}
        <Divider />
        <Table size="small">
          <TableHeader>
            <TableRow>
              <TableHeaderCell>Time</TableHeaderCell>
              <TableHeaderCell>File</TableHeaderCell>
              <TableHeaderCell>Action</TableHeaderCell>
              <TableHeaderCell>Status</TableHeaderCell>
              <TableHeaderCell>Message</TableHeaderCell>
            </TableRow>
          </TableHeader>
          <TableBody>
            {log.map((entry, index) => (
              <TableRow key={entry.id ?? index}>
                <TableCell>{new Date(entry.timestamp).toLocaleString()}</TableCell>
                <TableCell>{entry.fileName}</TableCell>
                <TableCell><Text style={{ textTransform: 'capitalize' }}>{entry.action}</Text></TableCell>
                <TableCell>
                  <Badge
                    appearance="tint"
                    color={entry.status === 'success' ? 'success' : entry.status === 'error' ? 'danger' : 'informative'}
                    size="small"
                  >
                    {entry.status}
                  </Badge>
                </TableCell>
                <TableCell>{entry.message}</TableCell>
              </TableRow>
            ))}
            {log.length === 0 && (
              <TableRow>
                <TableCell colSpan={5}>
                  <Text italic style={{ color: tokens.colorNeutralForeground3 }}>
                    No activity yet.
                  </Text>
                </TableCell>
              </TableRow>
            )}
          </TableBody>
        </Table>
      </Card>

      <Dialog open={emailOpen} onOpenChange={(_, data) => { if (!data.open) setEmailOpen(false); }}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Email the run report</DialogTitle>
            <DialogContent>
              <Field label="To" hint="Comma-separated email addresses" required>
                <Input value={to} onChange={(_, data) => setTo(data.value)} placeholder="owner@contoso.com, kb-team@contoso.com" />
              </Field>
              <Field label="Subject" required>
                <Input value={subject} onChange={(_, data) => setSubject(data.value)} />
              </Field>
              <Field label="Message (HTML)" required>
                <Textarea value={message} onChange={(_, data) => setMessage(data.value)} resize="vertical" rows={10} />
              </Field>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setEmailOpen(false)} disabled={emailSending}>Cancel</Button>
              <Button appearance="primary" onClick={sendEmail} disabled={emailSending || !to.trim() || !subject.trim() || !message.trim()}>
                {emailSending ? 'Sending…' : 'Send'}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
}
