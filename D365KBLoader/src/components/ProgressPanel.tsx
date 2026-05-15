import {
  Card, CardHeader, Text, ProgressBar, makeStyles, tokens,
  Table, TableHeader, TableRow, TableHeaderCell, TableBody, TableCell, Badge, Button
} from '@fluentui/react-components';
import { ArrowClockwise24Regular } from '@fluentui/react-icons';
import type { LogEntry } from '../types';

const useStyles = makeStyles({
  card: { padding: tokens.spacingHorizontalL },
  bar: { marginBottom: tokens.spacingVerticalL }
});

export interface ProgressPanelProps {
  done: number;
  total: number;
  errors: number;
  log: LogEntry[];
  onRefresh?: () => void;
}

export function ProgressPanel({ done, total, errors, log, onRefresh }: ProgressPanelProps) {
  const s = useStyles();
  const pct = total === 0 ? 0 : done / total;
  return (
    <Card className={s.card}>
      <CardHeader
        header={<Text weight="semibold" size={500}>3. Load progress &amp; activity log</Text>}
        description={<Text size={200}>{done} / {total} processed · {errors} errors</Text>}
        action={onRefresh && <Button icon={<ArrowClockwise24Regular />} onClick={onRefresh} appearance="subtle">Refresh log</Button>}
      />
      <div className={s.bar}><ProgressBar value={pct} thickness="large" /></div>
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
              <TableCell>{l.action}</TableCell>
              <TableCell>
                <Badge color={l.status === 'success' ? 'success' : l.status === 'error' ? 'danger' : 'informative'} size="small">
                  {l.status}
                </Badge>
              </TableCell>
              <TableCell>{l.message}</TableCell>
            </TableRow>
          ))}
          {log.length === 0 && (
            <TableRow><TableCell colSpan={5}><Text italic>No activity yet.</Text></TableCell></TableRow>
          )}
        </TableBody>
      </Table>
    </Card>
  );
}
