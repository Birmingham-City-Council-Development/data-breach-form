// Deployment settings. Every value here is injected at build time from .env
// by gulpfile.js (webpack DefinePlugin) — see .env.example for what each one
// means. Edit .env, not this file, then rebuild.

// Comma-separated .env value → trimmed, non-empty list.
function list(value: string | undefined, fallback: string[]): string[] {
  const items = (value || '').split(',').map(s => s.trim()).filter(Boolean);
  return items.length ? items : fallback;
}

// Organisation name shown in email headers/footers. Leave blank to omit.
export const ORG_NAME: string = process.env.ORG_NAME || '';

// Name of the team that handles breach reports, used for "on behalf of"
// attributions and reply notices.
export const TEAM_NAME: string = process.env.TEAM_NAME || 'Data Breach Notification Team';

// Team named in the pre-filled "Reason for decision" wording. Falls back to TEAM_NAME.
export const REVIEW_TEAM_NAME: string = process.env.REVIEW_TEAM_NAME || TEAM_NAME;

export const SP_LIST_NAME: string = process.env.SP_LIST_NAME || 'Data Breaches';

// SharePoint groups whose members can see and edit all breach reports.
// Exact match, case-sensitive.
export const SP_ADMIN_GROUPS: string[] = list(process.env.SP_ADMIN_GROUPS, ['Data Breach Owners']);

// SharePoint groups whose members can see (Read only) every breach report,
// without needing full admin (Contribute) rights. Panel Recommendations
// edits for this group go through a separate elevated endpoint rather than
// direct item permissions — see appendCommLogEntryElevated's sibling for
// that field.
export const SP_REVIEWER_GROUPS: string[] = list(process.env.SP_REVIEWER_GROUPS, ['Data Breach Reviewers']);

// The central mailbox that handles all incoming/outgoing report correspondence.
// Fixed per deployment (build-time), not property-pane configurable — this is
// a team mailbox, not a per-page setting like the elevated endpoint keys.
export const ADMIN_NOTIFICATION_EMAIL: string = process.env.ADMIN_NOTIFICATION_EMAIL || '';

// Linked from the TEST_MODE warning banner, pointing testers at the real service.
export const LIVE_SERVICE_URL: string = process.env.LIVE_SERVICE_URL || '';

// Per-file upload limit on the submission form, in MB. Keep the APIM
// request-size cap above this plus ~35% for base64 overhead.
export const MAX_ATTACHMENT_MB: number = Number(process.env.MAX_ATTACHMENT_MB) || 20;

// File extensions the submission form accepts (lowercase, no dots). The
// upload-attachment Logic App should enforce the same list server-side.
export const ALLOWED_ATTACHMENT_EXTENSIONS: string[] = list(process.env.ALLOWED_ATTACHMENT_EXTENSIONS, [
  'pdf', 'doc', 'docx', 'odt', 'rtf', 'txt',
  'xls', 'xlsx', 'ods', 'csv',
  'ppt', 'pptx',
  'msg', 'eml',
  'png', 'jpg', 'jpeg', 'gif',
]).map(ext => ext.toLowerCase().replace(/^\./, ''));

// Default APIM URLs for the elevated Logic App operations. A value set in a
// web part's property pane overrides these. API keys are deliberately not
// here — they stay in the property pane so they never get baked into the .sppkg.
export const DEFAULT_ENDPOINT_URLS = {
  sendMail: process.env.SENDMAIL_URL || '',
  appendCommLog: process.env.APPEND_COMM_LOG_URL || '',
  updatePanelRecommendations: process.env.UPDATE_PANEL_RECOMMENDATIONS_URL || '',
  uploadAttachment: process.env.UPLOAD_ATTACHMENT_URL || '',
};
