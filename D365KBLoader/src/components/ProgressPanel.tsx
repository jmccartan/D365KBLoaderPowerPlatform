import {
  Card, Text, ProgressBar, makeStyles, tokens,
  Table, TableHeader, TableRow, TableHeaderCell, TableBody, TableCell, Badge, Button, Divider, Spinner,
  MessageBar, MessageBarBody, MessageBarTitle, MessageBarActions
} from '@fluentui/react-components';
import {
  DocumentTable20Filled, DocumentBulletListMultiple24Filled,
  CheckmarkCircle24Filled, DismissCircle24Filled, History24Regular
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
    border: `1px solid ${tokens.colorNeutralStroke2}`,
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
});

export interface ProgressPanelProps {
  done: number;
  total: number;
  errors: number;
  log: LogEntry[];
  onSaveReport: () => void;
  reportSaving: boolean;
  reportResult?: ReportResult;
  reportError?: string;
}

export function ProgressPanel({
  done, total, errors, log,
  onSaveReport, reportSaving, reportResult, reportError,
}: ProgressPanelProps) {
  const s = useStyles();
  const pct = total === 0 ? 0 : done / total;
  const successes = Math.max(0, done - errors);

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
          <Button
            appearance="primary"
            icon={reportSaving ? <Spinner size="tiny" /> : <DocumentTable20Filled />}
            onClick={onSaveReport}
            disabled={reportSaving || log.length === 0}
          >
            {reportSaving ? 'Saving…' : 'Save Excel report'}
          </Button>
        </div>
        {reportResult && (
          <MessageBar intent="success" style={{ margin: `0 ${tokens.spacingHorizontalXL} ${tokens.spacingVerticalM}` }}>
            <MessageBarBody>
              <MessageBarTitle>Report saved</MessageBarTitle>
              <code>{reportResult.fileName}</code> {reportResult.downloaded ? 'downloaded to your browser' : `→ ${reportResult.location}`}
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
            {log.map((l, i) => (
              <TableRow key={l.id ?? i}>
                <TableCell>{new Date(l.timestamp).toLocaleString()}</TableCell>
                <TableCell>{l.fileName}</TableCell>
                <TableCell><Text style={{ textTransform: 'capitalize' }}>{l.action}</Text></TableCell>
                <TableCell>
                  <Badge
                    appearance="tint"
                    color={l.status === 'success' ? 'success' : l.status === 'error' ? 'danger' : 'informative'}
                    size="small"
                  >
                    {l.status}
                  </Badge>
                </TableCell>
                <TableCell>{l.message}</TableCell>
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
    </div>
  );
}
