import { stripHtmlIfPresent, IBreachListRow } from '../../dataBreachForm/services/SharePointService';
import { FullItem } from './reportViewerTypes';
import { formatDate, getRef, getReportFieldRows } from './reportViewerUtils';

function escapeCsvCell(val: string): string {
  if (!val) return '';
  const str = String(val);
  if (str.indexOf(',') !== -1 || str.indexOf('"') !== -1 || str.indexOf('\n') !== -1) {
    return `"${str.replace(/"/g, '""')}"`;
  }
  return str;
}

function downloadText(csv: string, filename: string): void {
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export function printReport(): void {
  const styleId = '__breach-print-style';
  let style = document.getElementById(styleId) as HTMLStyleElement | null;
  if (!style) {
    style = document.createElement('style');
    style.id = styleId;
    document.head.appendChild(style);
  }
  style.innerHTML = [
    '@media print {',
    '  body * { visibility: hidden !important; }',
    '  #breach-report-detail, #breach-report-detail * { visibility: visible !important; }',
    // position: fixed here would clip the printout to a single page — fixed/
    // absolute content isn't paginated by print engines, it's just pinned to
    // the first page. Static positioning lets it flow across as many pages
    // as the content needs.
    '  #breach-report-detail { width: 100%; padding: 20px; }',
    '  #breach-report-detail .no-print { display: none !important; }',
    // These sit outside #breach-report-detail (phase banner, back links, the
    // test-mode warning) — collapse them with display:none rather than
    // relying on visibility:hidden, so they don't leave blank space above
    // the report now that it's back in normal document flow.
    '  .govuk-phase-banner, .govuk-back-link, .govuk-warning-text { display: none !important; }',
    '  a::after { content: none !important; }',
    '}',
  ].join('\n');
  window.print();
  window.addEventListener('afterprint', function handler() {
    const el = document.getElementById(styleId);
    if (el && el.parentNode) { el.parentNode.removeChild(el); }
    window.removeEventListener('afterprint', handler);
  });
}

export function downloadSingleCsv(item: FullItem): void {
  const rows = getReportFieldRows(item);
  const csv = ['Field,Value', ...rows.map(function(r) { return escapeCsvCell(r[0]) + ',' + escapeCsvCell(stripHtmlIfPresent(r[1])); })].join('\n');
  const ref = getRef(item.Id).replace(/[^a-zA-Z0-9-]/g, '-');
  downloadText(csv, `${ref}.csv`);
}

export function downloadCsv(rows: IBreachListRow[]): void {
  const headers = [
    'Reference', 'Status', 'Submitted by', 'Date submitted',
    'Directorate', 'Reported to ICO', 'Approx individuals affected',
    'Impact to data subjects', 'Breach categories', 'Nature of breach', 'Date discovered',
  ];
  const body = rows.map(i => [
    escapeCsvCell(getRef(i.Id)),
    escapeCsvCell(i.Status || 'Open'),
    escapeCsvCell(i.Author ? i.Author.Title : ''),
    escapeCsvCell(formatDate(i.Created)),
    escapeCsvCell(i.Directorate_x0020_of_x0020_repor || ''),
    escapeCsvCell(i.Reported_x0020_to_x0020_ICO || ''),
    escapeCsvCell(i.Approx_x0020_individuals_x0020_a || ''),
    escapeCsvCell(i.Impact_x0020_to_x0020_Data_x0020 || ''),
    escapeCsvCell(i.Breach_x0020_categories || ''),
    escapeCsvCell(i.Categorise_x0020_nature_x0020_of || ''),
    escapeCsvCell(i.Date_x0020_breach_x0020_discover ? formatDate(i.Date_x0020_breach_x0020_discover) : ''),
  ].join(','));
  const csv = [headers.join(','), ...body].join('\n');
  downloadText(csv, `data-breaches-${new Date().toISOString().slice(0, 10)}.csv`);
}
