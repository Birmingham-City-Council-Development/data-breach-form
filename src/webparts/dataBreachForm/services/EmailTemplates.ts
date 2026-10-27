// ── EmailTemplates.ts ─────────────────────────────────────────────────────────
//
// HTML email templates for data breach notifications.
// Edit the content inside each build*Email function to change wording.
// Edit the helper functions (wrap, infoTable, button, etc.) to change styling.
// ─────────────────────────────────────────────────────────────────────────────

import { ADMIN_NOTIFICATION_EMAIL, ORG_NAME } from '../config';

export interface IBreachEmailData {
  ref: string;
  itemId: number;
  siteUrl: string;
  officerName: string;
  officerEmail: string;
  responsibleName: string;
  responsibleEmail: string;
  lineManager: string;
  lineManagerEmail: string;
  officerDirectorate: string;
  officerDivision: string;
  dateOccurred: string;
  timeOccurred: string;
  dateDiscovered: string;
  timeDiscovered: string;
  breachDetails: string;
  approxIndividuals: string;
  impactToSubjects: string;
  natureOfBreach: string;
  thirdPartyInvolved: string;
  isSelfResponsible: boolean;
}

export interface ICommEmailData {
  ref: string;
  itemId: number;
  siteUrl: string;
  recipientName: string;
  message: string;
}

export type BreachRole = 'submitter' | 'responsible' | 'lineManager';

export interface ICommunicationNoticeEmailData {
  ref: string;
  itemId: number;
  siteUrl: string;
  recipientName: string;
  recipientRoles: BreachRole[];
  fromName: string;
  toName: string;
  message: string;
  direction: 'request' | 'reply';
}

export interface IClosedEmailData {
  ref: string;
  recipientName: string;
  reason?: string;
  reportedToIco?: string;
  eLearningRequired?: string;
  proceduresReviewRequired?: string;
}

// ── Private helpers ───────────────────────────────────────────────────────────

function esc(s: string): string {
  return (s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// Full HTML document wrapper with branded header and footer. options lets
// callers override the banner/footer copy that otherwise assumes there's a
// link back into the report below — needed for emails (like the case-closed
// one) that don't carry one.
function wrap(title: string, content: string, options?: { noticeText?: string; footerText?: string }): string {
  const noticeText = options && options.noticeText;
  const footerText = options && options.footerText;
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>${esc(title)}</title>
</head>
<body style="margin:0;padding:0;background-color:#f3f2f1;font-family:Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" role="presentation" style="background-color:#f3f2f1;">
  <tr>
    <td align="center" style="padding:24px 16px;">
      <table width="100%" cellpadding="0" cellspacing="0" role="presentation" style="max-width:620px;">

        <!-- Yellow accent bar -->
        <tr>
          <td height="6" bgcolor="#ffb700" style="background-color:#ffb700;font-size:0;line-height:0;">&nbsp;</td>
        </tr>

        <!-- Header -->
        <tr>
          <td bgcolor="#003078" style="background-color:#003078;padding:20px 28px 24px;">
            ${ORG_NAME ? `<p style="margin:0 0 6px 0;font-family:Arial,sans-serif;font-size:12px;font-weight:bold;
                      color:#7ab3db;text-transform:uppercase;letter-spacing:1px;">
              ${esc(ORG_NAME)}
            </p>` : ''}
            <p style="margin:0;font-family:Arial,sans-serif;font-size:22px;font-weight:bold;
                      color:#ffffff;line-height:1.3;">
              ${esc(title)}
            </p>
          </td>
        </tr>

        <!-- Do-not-reply notice — placed above the fold, before any content,
             since a footer note is too easy to skim past. Wording avoids
             claiming the mailbox is unmonitored (it's a real shared mailbox,
             just not one that feeds back into the report) and instead says
             plainly what to do instead. -->
        <tr>
          <td bgcolor="#fff3cd" style="background-color:#fff3cd;padding:14px 28px;border-bottom:2px solid #ffb700;">
            <p style="margin:0;font-family:Arial,sans-serif;font-size:13px;font-weight:bold;color:#594d00;line-height:1.5;">
              ${esc(noticeText || 'Do not reply to this email — replies are not added to the report. Use the link below to respond via the online report.')}
            </p>
          </td>
        </tr>

        <!-- Body -->
        <tr>
          <td bgcolor="#ffffff" style="background-color:#ffffff;padding:28px 28px 8px;">
            ${content}
          </td>
        </tr>

        <!-- Footer -->
        <tr>
          <td bgcolor="#ffffff" style="background-color:#ffffff;border-top:1px solid #b1b4b6;padding:16px 28px 24px;">
            <p style="margin:0;font-family:Arial,sans-serif;font-size:12px;color:#505a5f;line-height:1.6;">
              This is an automated notification from ${ORG_NAME ? esc(ORG_NAME) + '&#39;s' : 'the'} Data Breach Reporting System.<br>
              ${esc(footerText || 'Please do not reply to this email — use the button above to view this report and communicate with the team.')}
            </p>
          </td>
        </tr>

      </table>
    </td>
  </tr>
</table>
</body>
</html>`;
}

// Labelled key-value info box with blue left border
function infoTable(rows: Array<[string, string]>): string {
  const cells = rows
    .filter(function(r) { return !!(r[1] && r[1].trim()); })
    .map(function(r) {
      return `<tr>
        <td style="font-weight:bold;color:#505a5f;width:170px;vertical-align:top;padding:5px 14px 5px 0;font-size:13px;">${esc(r[0])}</td>
        <td style="color:#0b0c0c;vertical-align:top;padding:5px 0;font-size:13px;">${esc(r[1])}</td>
      </tr>`;
    })
    .join('');
  return `<table width="100%" cellpadding="0" cellspacing="0"
    style="border-left:4px solid #1d70b8;background-color:#f3f2f1;margin:8px 0 22px;">
    <tr><td style="padding:14px 18px;">
      <table width="100%" cellpadding="0" cellspacing="0">${cells}</table>
    </td></tr>
  </table>`;
}

// Clickable call-to-action button
function ctaButton(label: string, url: string): string {
  return `<table cellpadding="0" cellspacing="0" role="presentation" style="margin:22px 0 8px;">
    <tr>
      <td bgcolor="#1d70b8" style="background-color:#1d70b8;padding:13px 24px;">
        <a href="${esc(url)}"
           style="font-family:Arial,sans-serif;font-size:14px;font-weight:bold;
                  color:#ffffff;text-decoration:none;white-space:nowrap;">
          ${esc(label)} &#8594;
        </a>
      </td>
    </tr>
  </table>`;
}

// Small section heading
function sectionHeading(text: string): string {
  return `<p style="margin:22px 0 2px;font-size:11px;color:#505a5f;
                    text-transform:uppercase;letter-spacing:0.9px;font-weight:bold;">${esc(text)}</p>`;
}

// Standard paragraph
function para(text: string): string {
  return `<p style="margin:0 0 14px;font-size:14px;color:#0b0c0c;line-height:1.6;">${text}</p>`;
}

// Yellow warning / info panel
function noticePanel(html: string): string {
  return `<div style="border-left:4px solid #ffb700;background-color:#fff3cd;
                      padding:12px 16px;margin:4px 0 18px;">
    <p style="margin:0;font-size:14px;color:#0b0c0c;line-height:1.6;">${html}</p>
  </div>`;
}

// Quoted message block (blue left border)
function messageBlock(text: string): string {
  return `<div style="background-color:#f3f2f1;border-left:4px solid #1d70b8;padding:14px 18px;
               margin:8px 0 22px;font-size:14px;color:#0b0c0c;line-height:1.6;
               white-space:pre-wrap;">${esc(text)}</div>`;
}

// ── Submission email templates ────────────────────────────────────────────────

// 1. Sent to the officer who submitted the report
export function buildSubmitterEmail(d: IBreachEmailData): string {
  const reportUrl = `${d.siteUrl}/SitePages/Report-A-Data-Breach.aspx?reportId=${d.itemId}`;
  const responsibleNote = d.isSelfResponsible
    ? noticePanel(
        `You have also been identified as the <strong>person responsible</strong> for this breach. ` +
        `The data breach notification team may contact you for further information — please be prepared to ` +
        `discuss the circumstances of the breach and any actions you have taken.`
      )
    : '';

  return wrap('Data Breach Report Submitted', [
    para(`Dear ${esc(d.officerName)},`),
    para(
      `Thank you for submitting a data breach report. Your report has been received ` +
      `and is now being reviewed by the data breach notification team.`
    ),
    responsibleNote,
    infoTable([
      ['Reference',        d.ref],
      ['Date of breach',   d.dateOccurred + d.timeOccurred],
      ['Date discovered',  d.dateDiscovered + d.timeDiscovered],
    ]),
    sectionHeading('Details of the breach'),
    messageBlock(d.breachDetails || 'Not provided'),
    para(
      `You can track the progress of your report, view any communications from the team, ` +
      `and submit additional information using the link below. Please keep your reference ` +
      `number to hand.`
    ),
    ctaButton('Track your report', reportUrl),
  ].join(''));
}

// 2. Sent to the person identified as responsible for the breach
export function buildResponsibleEmail(d: IBreachEmailData): string {
  const reportUrl = `${d.siteUrl}/SitePages/Report-A-Data-Breach.aspx?reportId=${d.itemId}`;
  return wrap('Data Breach — You Have Been Identified as the Person Responsible', [
    para(`Dear ${esc(d.responsibleName)},`),
    para(
      `A data breach report has been submitted in which you have been identified as ` +
      `the <strong>person responsible</strong> for the incident.`
    ),
    infoTable([
      ['Reference',        d.ref],
      ['Date of breach',   d.dateOccurred + d.timeOccurred],
      ['Date discovered',  d.dateDiscovered + d.timeDiscovered],
    ]),
    sectionHeading('Details of the breach'),
    messageBlock(d.breachDetails || 'Not provided'),
    noticePanel(
      `The data breach notification team may contact you for further information regarding this incident. ` +
      `Please be prepared to discuss the circumstances of the breach and any actions you have taken. ` +
      `Do not discuss this matter with others until contacted by the team.`
    ),
    para(`You can view the full report and communicate with the data breach notification team using the link below.`),
    ctaButton('View your report', reportUrl),
  ].join(''));
}

// 3. Sent to the line manager of the person responsible
export function buildLineManagerEmail(d: IBreachEmailData): string {
  const reportUrl = `${d.siteUrl}/SitePages/Report-A-Data-Breach.aspx?reportId=${d.itemId}`;
  return wrap('Data Breach Notification', [
    para(`Dear ${esc(d.lineManager)},`),
    para(
      `Please be advised that a data breach report has been submitted involving a member ` +
      `of staff you manage, <strong>${esc(d.responsibleName)}</strong>.`
    ),
    infoTable([
      ['Reference',        d.ref],
      ['Employee',         d.responsibleName],
      ['Date of breach',   d.dateOccurred + d.timeOccurred],
      ['Date discovered',  d.dateDiscovered + d.timeDiscovered],
    ]),
    sectionHeading('Details of the breach'),
    messageBlock(d.breachDetails || 'Not provided'),
    para(
      `No further action is required from you at this time. However, please be aware that ` +
      `you may be contacted by the data breach notification team as part of their review process.`
    ),
    ctaButton('View full report', reportUrl),
  ].join(''));
}

// 4. Sent to the admin / shared mailbox on every new submission
export function buildAdminEmail(d: IBreachEmailData): string {
  const reportUrl = `${d.siteUrl}/SitePages/Report-A-Data-Breach.aspx?reportId=${d.itemId}`;
  const directorate = d.officerDivision
    ? `${d.officerDirectorate} — ${d.officerDivision}`
    : d.officerDirectorate;

  return wrap('New Data Breach Report Received', [
    para(`A new data breach report has been received and requires review.`),

    sectionHeading('People Involved'),
    infoTable([
      ['Reporting officer',   `${d.officerName} (${d.officerEmail})`],
      ['Person responsible',  `${d.responsibleName} (${d.responsibleEmail})`],
      ['Line manager',        `${d.lineManager} (${d.lineManagerEmail})`],
      ['Directorate',         directorate],
    ]),

    sectionHeading('Key Information'),
    infoTable([
      ['Date of breach',          d.dateOccurred + d.timeOccurred],
      ['Date discovered',         d.dateDiscovered + d.timeDiscovered],
      ['Approx. individuals',     d.approxIndividuals],
      ['Impact to data subjects', d.impactToSubjects],
      ['Nature of breach',        d.natureOfBreach],
      ['Third party involved',    d.thirdPartyInvolved],
    ]),

    sectionHeading('Breach Details'),
    `<div style="background-color:#f3f2f1;border-left:4px solid #b1b4b6;padding:14px 18px;
                 margin:8px 0 22px;font-size:14px;color:#0b0c0c;line-height:1.6;
                 white-space:pre-wrap;">${esc(d.breachDetails || 'Not provided')}</div>`,

    ctaButton('View full report', reportUrl),
  ].join(''));
}

// ── Communication log email templates ─────────────────────────────────────────

// Sent by an admin to the report submitter requesting more information
export function buildRequestInfoEmail(d: ICommEmailData): string {
  const reportUrl = `${d.siteUrl}/SitePages/Report-A-Data-Breach.aspx?reportId=${d.itemId}`;
  return wrap('Additional Information Required', [
    para(`Dear ${esc(d.recipientName)},`),
    para(
      `The data breach notification team requires some additional information regarding ` +
      `a report you are involved in.`
    ),
    infoTable([['Reference', d.ref]]),
    sectionHeading('Information Required'),
    messageBlock(d.message),
    para(`Please use the button below to view your report and submit your response.`),
    ctaButton('Submit more information', reportUrl),
  ].join(''));
}

export interface IUserReplyEmailData {
  ref: string;
  itemId: number;
  siteUrl: string;
  // The person actually typing and sending this reply — always the person
  // responsible in practice, since that's the only role the "Submit more
  // information" action is shown to (see isCurrentUserRecipient in
  // ReportViewer.tsx). Kept as its own name+email pair (not reused from
  // responsibleName below) so the two stay visually distinct even though
  // they're the same person today, in case that ever changes.
  submittedByName: string;
  submittedByEmail: string;
  // The report's named person responsible, shown alongside submittedBy so
  // whoever's reading this in the shared inbox can immediately confirm the
  // response actually came from the right person, not just trust the name.
  responsibleName: string;
  message: string;
}

// Sent by the person responsible back to the data breach notification team
// with additional information. Previously showed the report's original
// Author (submitter) as "Submitted by" regardless of who actually sent the
// reply — misleading whenever the person responsible differs from whoever
// filed the report, which is the common case.
export function buildUserReplyEmail(d: IUserReplyEmailData): string {
  const reportUrl = `${d.siteUrl}/SitePages/Report-A-Data-Breach.aspx?reportId=${d.itemId}`;
  return wrap('Response Received', [
    para(`A response has been submitted for the following data breach report.`),
    infoTable([
      ['Reference',            d.ref],
      ['Person responsible',   d.responsibleName],
      ['Response submitted by', d.submittedByEmail ? `${d.submittedByName} (${d.submittedByEmail})` : d.submittedByName],
    ]),
    sectionHeading('Response'),
    messageBlock(d.message),
    ctaButton('View full report', reportUrl),
  ].join(''));
}

const BREACH_ROLE_LABELS: Record<BreachRole, string> = {
  submitter: 'the submitter',
  responsible: 'the person responsible',
  lineManager: 'the line manager',
};

// e.g. ["responsible"] -> "the person responsible"; ["submitter","lineManager"]
// -> "the submitter and the line manager" — a person can hold more than one
// role on the same report, so this states all of them rather than picking one.
function describeBreachRoles(roles: BreachRole[]): string {
  const labels = roles.map(function(r) { return BREACH_ROLE_LABELS[r]; });
  if (labels.length <= 1) return labels[0] || 'someone named on this report';
  return `${labels.slice(0, -1).join(', ')} and ${labels[labels.length - 1]}`;
}

// Sent to everyone named on a report (submitter, person responsible, line
// manager) other than whoever directly sent/received a given message, so
// they stay aware of the back-and-forth even when it isn't addressed to them.
// Always states exactly which role(s) put the recipient on this notice.
export function buildCommunicationNoticeEmail(d: ICommunicationNoticeEmailData): string {
  const reportUrl = `${d.siteUrl}/SitePages/Report-A-Data-Breach.aspx?reportId=${d.itemId}`;

  // A submitter who isn't also the person responsible doesn't need the
  // back-and-forth detail on a request for more information — just the
  // reassurance that the team has gone to the right person, not them.
  const isSubmitterOnlyRequest = d.direction === 'request' &&
    d.recipientRoles.length === 1 && d.recipientRoles[0] === 'submitter';

  if (isSubmitterOnlyRequest) {
    return wrap('Data Breach Report — Update', [
      para(`Dear ${esc(d.recipientName)},`),
      para(
        `You're receiving this as the submitter of the report below. The data breach notification team ` +
        `has reached out to ${esc(d.toName)}, the person responsible for this report, requesting ` +
        `additional information. No action is needed from you at this time.`
      ),
      infoTable([['Reference', d.ref]]),
      para(`You can view the full report using the link below.`),
      ctaButton('View report', reportUrl),
    ].join(''));
  }

  // The person responsible is never a recipient here — they're a primary
  // party (they get the real request/reply emails directly), not a
  // bystander — so every notice reaching this point is safely "no action
  // needed" regardless of direction.
  const roleDescription = describeBreachRoles(d.recipientRoles);
  const intro = d.direction === 'request'
    ? `The data breach notification team has requested further information on the report below. ` +
      `You're receiving this because you're ${roleDescription} for this report — no action ` +
      `is needed from you unless you're contacted directly.`
    : `A response has been submitted on the report below. You're receiving this because ` +
      `you're ${roleDescription} for this report — no action is needed from you unless ` +
      `you're contacted directly.`;
  return wrap('Data Breach Report — Communication Update', [
    para(`Dear ${esc(d.recipientName)},`),
    para(intro),
    infoTable([
      ['Reference', d.ref],
      ['From',      d.fromName],
      ['To',        d.toName],
    ]),
    sectionHeading('Message'),
    messageBlock(d.message),
    para(`You can view the full report and communication history using the link below.`),
    ctaButton('View full report', reportUrl),
  ].join(''));
}

// Sent to everyone named on a report (submitter, person responsible, line
// manager) plus the shared inbox when the data breach notification team
// reopens a closed case — same fixed "broadcast to everyone" recipient set
// as buildClosedEmail, since reopening restores everyone's access to the
// report (see isClosed in syncItemAccess), not just the person responsible.
// d.message is optional — a reopen doesn't have to come with a note, but
// when the team knows up front what's needed (the common case), including
// it here saves a separate "request more information" round trip.
export function buildReopenedEmail(d: ICommEmailData): string {
  const reportUrl = `${d.siteUrl}/SitePages/Report-A-Data-Breach.aspx?reportId=${d.itemId}`;
  const hasMessage = !!(d.message && d.message.trim());
  return wrap('Data Breach Report Reopened', [
    para(`Dear ${esc(d.recipientName)},`),
    para(
      `The data breach notification team has reopened the data breach report below for further review.`
    ),
    infoTable([['Reference', d.ref]]),
    hasMessage ? sectionHeading('Additional Information Required') : '',
    hasMessage ? messageBlock(d.message) : para('No further action is required from you at this time.'),
    para(`Please use the button below to view the report${hasMessage ? ' and submit any additional information requested' : ''}.`),
    ctaButton('View report', reportUrl),
  ].join(''));
}

// Sent to the report submitter when the data breach notification team closes their case.
// Deliberately carries no link back into the report — access to a closed
// report is removed (see syncItemAccess's isClosed handling), so a link here
// would just lead to the "no longer accessible" error page.
export function buildClosedEmail(d: IClosedEmailData): string {
  const hasPanelRecommendations = !!(d.reportedToIco || d.eLearningRequired || d.proceduresReviewRequired || (d.reason && d.reason.trim()));
  return wrap('Data Breach Report Closed', [
    para(`Dear ${esc(d.recipientName)},`),
    para(
      `The data breach notification team has completed its review of the reported breach ` +
      `below, and it has now been closed. Please see if there has been any further action ` +
      `that is required from you by viewing the panel recommendations below.`
    ),
    infoTable([['Reference', d.ref]]),
    hasPanelRecommendations ? sectionHeading('Panel Recommendations') : '',
    hasPanelRecommendations ? infoTable([
      ['Reported to ICO',              d.reportedToIco || ''],
      ['Staff e-Learning training',    d.eLearningRequired || ''],
      ['Procedures review required',   d.proceduresReviewRequired || ''],
    ]) : '',
    d.reason && d.reason.trim() ? messageBlock(d.reason) : '',
    noticePanel(
      `This report is now closed and access to it has been removed — it can no longer be viewed online. ` +
      `A summary of this case has been attached to this email as a PDF; please keep this for your records.`
    ),
    para(
      `If you believe this is a mistake, please contact us at ` +
      `<a href="mailto:${esc(ADMIN_NOTIFICATION_EMAIL)}" style="color:#1d70b8;">${esc(ADMIN_NOTIFICATION_EMAIL)}</a>.`
    ),
  ].join(''), {
    noticeText: 'This report is now closed and no longer accessible online — see below for how to get in touch if you have questions.',
    footerText: 'This report is now closed and no longer accessible online.',
  });
}
