import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { ICommLogEntry } from '../../dataBreachForm/services/SharePointService';
import { IEmailAttachment } from '../components/IReportViewerProps';
import { TEAM_NAME } from '../../dataBreachForm/config';

// Directions where sentBy/sentByName is a staff member (reviewer/admin)
// acting on behalf of the team, not a named individual on the report giving
// their own response — 'Reply' is the one direction where sentBy genuinely
// is the other party (submitter/person responsible), so it's deliberately
// excluded here and left showing their own identity.
const STAFF_DIRECTIONS = ['Outbound', 'CaseClosed', 'CaseReopened', 'StatusChange'];

// Shows which team member actually actioned the entry, not just the shared
// team mailbox — matters for accountability since the person who requested
// further information isn't necessarily who later reviews the response and
// closes the case (e.g. they're on leave in the meantime). Falls back to
// the raw email for entries logged before sentByName existed.
export function staffAttribution(entry: ICommLogEntry): string {
  return `${entry.sentByName || entry.sentBy} (on behalf of ${TEAM_NAME})`;
}

export interface IReportSection {
  title: string;
  rows: Array<[string, string]>;
}

function formatDateTime(iso: string): string {
  if (!iso) return '';
  return new Date(iso).toLocaleString('en-GB', {
    day: '2-digit', month: '2-digit', year: 'numeric',
    hour: '2-digit', minute: '2-digit',
  });
}

// GOV.UK palette used across the app's own summary-list styling — kept in
// sync by eye with the colours in ReportViewer.tsx/govuk-frontend, not
// pulled from a shared constant (there isn't one — the app styles via CSS
// classes, this is the one place colours need to be literal RGB triples
// for jsPDF).
const GOVUK_TEXT: [number, number, number] = [11, 12, 12]; // #0b0c0c
const GOVUK_GREY: [number, number, number] = [80, 90, 95]; // #505a5f
const GOVUK_BLUE: [number, number, number] = [29, 112, 184]; // #1d70b8
const GOVUK_PANEL_GREY: [number, number, number] = [243, 242, 241]; // #f3f2f1
const GOVUK_BORDER: [number, number, number] = [177, 180, 182]; // #b1b4b6
const GOVUK_LIGHT_BLUE: [number, number, number] = [222, 234, 246];

// jsPDF's standard "helvetica" font only supports WinAnsiEncoding (cp1252) —
// it has no glyph for U+2192 (→). Feeding it an unsupported code point
// doesn't just render a missing-glyph box, it throws autoTable's width
// calculation off badly enough that the whole cell falls back to spacing
// every character apart onto its own line, which is why the comm log used
// to render as "N i g e l . A r t h u r @ ...". Kept as separate From/To
// rows instead (see below) rather than reintroducing any arrow character.

// Friendly labels for comm-log entry headers, mirroring the tags shown
// on-screen in ReportViewer.tsx (getStatusChangeStyle / the Info
// requested / Reply received tags) so the PDF reads the same way.
const DIRECTION_LABELS: { [key: string]: string } = {
  Outbound: 'Request for More Information',
  Reply: 'User Information Submitted',
  CaseClosed: 'Case Closed',
  CaseReopened: 'Case Reopened',
};

const STATUS_CHANGE_LABELS: { [status: string]: string } = {
  Closed: 'Case Closed',
  Open: 'Case Reopened',
  Pending: 'Marked Pending',
};

function describeCommLogEntry(entry: ICommLogEntry): string {
  if (entry.direction === 'StatusChange') {
    return STATUS_CHANGE_LABELS[entry.subject] || 'Status Change';
  }
  return DIRECTION_LABELS[entry.direction] || entry.direction;
}

// Builds the "final PDF summary" attached to case-closed emails, styled to
// resemble the app's own GOV.UK summary-list report page (bordered
// label/value rows grouped under section headers) rather than being a
// wall of plain text — sections/rows are passed in already grouped to
// match exactly what ReportViewer.tsx's <Section> components show on
// screen (see getReportFieldSections there), so this file stays decoupled
// from the FullItem type and just lays out whatever it's given.
export function buildReportPdfBase64(ref: string, status: string, sections: IReportSection[], commLog: ICommLogEntry[]): IEmailAttachment {
  const doc = new jsPDF({ unit: 'pt', format: 'a4' });
  const marginX = 40;
  const pageWidth = doc.internal.pageSize.getWidth();
  const usableWidth = pageWidth - marginX * 2;
  const labelColWidth = 170;

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(18);
  doc.setTextColor(GOVUK_TEXT[0], GOVUK_TEXT[1], GOVUK_TEXT[2]);
  doc.text('Data Breach Report', marginX, 50);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(11);
  doc.setTextColor(GOVUK_GREY[0], GOVUK_GREY[1], GOVUK_GREY[2]);
  doc.text(`${ref}    Status: ${status || 'Open'}`, marginX, 68);

  let startY = 86;

  const sectionTableOptions = {
    theme: 'plain' as const,
    margin: { left: marginX, right: marginX },
    styles: {
      font: 'helvetica', fontSize: 9, textColor: GOVUK_TEXT,
      cellPadding: { top: 5, bottom: 5, left: 8, right: 8 },
      lineColor: GOVUK_BORDER, lineWidth: 0.75, overflow: 'linebreak' as const,
    },
    headStyles: {
      fillColor: GOVUK_PANEL_GREY, textColor: GOVUK_TEXT,
      fontStyle: 'bold' as const, fontSize: 9, halign: 'left' as const,
    },
    columnStyles: {
      0: { fontStyle: 'bold' as const, textColor: GOVUK_GREY, cellWidth: labelColWidth },
      1: { cellWidth: usableWidth - labelColWidth },
    },
  };

  sections.forEach(function(section) {
    // Skip fields nobody filled in — same idea as infoTable() in
    // EmailTemplates.ts, which drops blank rows from the emailed summary.
    const rows = section.rows.filter(function(r) { return !!(r[1] && r[1].trim()); });
    if (rows.length === 0) return;
    autoTable(doc, {
      ...sectionTableOptions,
      startY,
      head: [[{ content: section.title.toUpperCase(), colSpan: 2 }]],
      body: rows,
    });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    startY = (doc as any).lastAutoTable.finalY + 16;
  });

  autoTable(doc, {
    ...sectionTableOptions,
    startY,
    head: [[{ content: 'COMMUNICATION LOG', colSpan: 2 }]],
    body: commLog.length === 0 ? [[{ content: 'No communications have been logged for this report.', colSpan: 2 }]] : [],
  });
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  startY = (doc as any).lastAutoTable.finalY + 12;

  commLog.forEach(function(entry) {
    const metaRows: Array<[string, string]> = [['When', formatDateTime(entry.timestamp)]];
    if (entry.sentBy) {
      // Staff-driven entries show which reviewer/admin actually actioned
      // them, "on behalf of" the team — see staffAttribution. 'Reply' is
      // excluded (see STAFF_DIRECTIONS) since sentBy there is genuinely the
      // submitter/person responsible's own identity, not a staff member's.
      const from = STAFF_DIRECTIONS.indexOf(entry.direction) !== -1 ? staffAttribution(entry) : entry.sentBy;
      metaRows.push(['From', from]);
    }
    // CaseClosed notifies several people at once (submitter, person
    // responsible, line manager, shared inbox), so sentTo is a comma-joined
    // list rather than one address — one address per line reads far better
    // than a single long wrapped line, and is far less likely to straddle a
    // page break awkwardly. Every other direction only ever has one address,
    // so this is a no-op for those.
    if (entry.sentTo) metaRows.push(['To', entry.sentTo.split(',').map(function(s) { return s.trim(); }).filter(Boolean).join('\n')]);
    // Subject is just "Data Breach Report <ref> - ..." restating the
    // direction/reference already shown by the header above and the
    // report-level heading — redundant in the PDF, unlike in an actual
    // email where it's the first thing shown.
    if (entry.message) metaRows.push(['Message', entry.message]);
    autoTable(doc, {
      ...sectionTableOptions,
      startY,
      head: [[{ content: describeCommLogEntry(entry), colSpan: 2 }]],
      body: metaRows,
      headStyles: { ...sectionTableOptions.headStyles, fillColor: GOVUK_LIGHT_BLUE, textColor: GOVUK_BLUE },
      // Without these, a row whose wrapped content doesn't fit in the
      // remaining space on the current page gets split mid-row: the bottom
      // half of the wrapped text is pushed onto the next page underneath a
      // re-printed head row (autoTable's default is to repeat the head on
      // every page). That's what produced the "Case Reopened" header
      // appearing twice with "Team)" and "son?" (a mid-word split of a
      // longer message) orphaned beneath the second one. pageBreak: 'avoid'
      // keeps the whole (short) entry table together, moving it to the next
      // page as a unit instead of splitting through the middle of a cell.
      pageBreak: 'avoid',
      rowPageBreak: 'avoid',
    });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    startY = (doc as any).lastAutoTable.finalY + 12;
  });

  const dataUri = doc.output('datauristring');
  const base64 = dataUri.substring(dataUri.indexOf('base64,') + 'base64,'.length);
  return { filename: `${ref}.pdf`, contentType: 'application/pdf', base64 };
}
