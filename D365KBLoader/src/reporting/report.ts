import ExcelJS from 'exceljs';
import type { LogEntry, KbConfig } from '../types';

/**
 * Build a formatted .xlsx report of a KB-load run.
 * Returns the file bytes plus a suggested file name.
 */
export async function buildReportWorkbook(
  config: KbConfig,
  log: LogEntry[],
): Promise<{ buffer: ArrayBuffer; fileName: string }> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'D365 KB Loader';
  wb.created = new Date();

  // --- Summary sheet ---
  const summary = wb.addWorksheet('Summary');
  summary.columns = [
    { key: 'k', width: 28 },
    { key: 'v', width: 70 },
  ];

  const counts = countByAction(log);
  const startedAt = log.length
    ? new Date(Math.min(...log.map(l => new Date(l.timestamp).getTime())))
    : new Date();
  const finishedAt = log.length
    ? new Date(Math.max(...log.map(l => new Date(l.timestamp).getTime())))
    : new Date();

  const meta: Array<[string, string | number]> = [
    ['D365 KB Loader — Run Report', ''],
    ['', ''],
    ['Site URL', config.siteUrl || ''],
    ['Folder path', config.folderPath || ''],
    ['Started', startedAt.toLocaleString()],
    ['Finished', finishedAt.toLocaleString()],
    ['', ''],
    ['Files processed', counts.process.total],
    ['  • succeeded', counts.process.success],
    ['  • failed', counts.process.error],
    ['Articles loaded', counts.load.total],
    ['  • succeeded', counts.load.success],
    ['  • failed', counts.load.error],
    ['Files skipped (unsupported)', counts.skip.total],
  ];
  meta.forEach(([k, v], i) => {
    const row = summary.addRow({ k, v });
    if (i === 0) {
      row.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
      row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1278D2' } };
      row.height = 28;
      summary.mergeCells(`A${row.number}:B${row.number}`);
      row.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    } else if (k && !String(k).startsWith(' ')) {
      row.getCell(1).font = { bold: true };
    }
  });

  // --- Activity log sheet ---
  const sheet = wb.addWorksheet('Activity log', { views: [{ state: 'frozen', ySplit: 1 }] });
  sheet.columns = [
    { header: 'Timestamp', key: 'timestamp', width: 22 },
    { header: 'File', key: 'fileName', width: 44 },
    { header: 'Action', key: 'action', width: 12 },
    { header: 'Status', key: 'status', width: 12 },
    { header: 'Message', key: 'message', width: 60 },
    { header: 'Source path', key: 'sourcePath', width: 50 },
    { header: 'Knowledge article id', key: 'knowledgeArticleId', width: 40 },
  ];

  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B3A6F' } };
  headerRow.alignment = { vertical: 'middle' };
  headerRow.height = 22;

  for (const e of log) {
    const row = sheet.addRow({
      timestamp: new Date(e.timestamp).toLocaleString(),
      fileName: e.fileName,
      action: e.action,
      status: e.status,
      message: e.message,
      sourcePath: e.sourcePath ?? '',
      knowledgeArticleId: e.knowledgeArticleId ?? '',
    });
    row.alignment = { vertical: 'top', wrapText: true };
    const statusCell = row.getCell('status');
    const colors = statusFill(e.status);
    statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.bg } };
    statusCell.font = { color: { argb: colors.fg }, bold: true };
    statusCell.alignment = { horizontal: 'center', vertical: 'middle' };
  }

  sheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: sheet.columns.length },
  };

  const buffer = await wb.xlsx.writeBuffer();
  const fileName = `KB-Loader-Report-${stamp(new Date())}.xlsx`;
  return { buffer: buffer as ArrayBuffer, fileName };
}

function statusFill(status: LogEntry['status']): { bg: string; fg: string } {
  switch (status) {
    case 'success': return { bg: 'FFD7F0E1', fg: 'FF0E7C3F' };
    case 'error':   return { bg: 'FFF8D7D7', fg: 'FFB02A37' };
    default:        return { bg: 'FFE1EEFB', fg: 'FF0B3A6F' };
  }
}

function countByAction(log: LogEntry[]) {
  const acc = {
    process: { total: 0, success: 0, error: 0 },
    load:    { total: 0, success: 0, error: 0 },
    skip:    { total: 0, success: 0, error: 0 },
    scan:    { total: 0, success: 0, error: 0 },
  };
  for (const e of log) {
    const bucket = acc[e.action as keyof typeof acc];
    if (!bucket) continue;
    bucket.total++;
    if (e.status === 'success') bucket.success++;
    else if (e.status === 'error') bucket.error++;
  }
  return acc;
}

function stamp(d: Date): string {
  const pad = (n: number) => n.toString().padStart(2, '0');
  return (
    d.getFullYear().toString() +
    pad(d.getMonth() + 1) +
    pad(d.getDate()) +
    '-' +
    pad(d.getHours()) +
    pad(d.getMinutes()) +
    pad(d.getSeconds())
  );
}

/** Trigger a browser download for the given bytes. */
export function downloadBlob(buffer: ArrayBuffer, fileName: string) {
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
