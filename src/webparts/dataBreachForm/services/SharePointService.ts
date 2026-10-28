import { spfi, SPFI } from '@pnp/sp';
import { SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/attachments';
import '@pnp/sp/site-groups';
import '@pnp/sp/site-users';
import '@pnp/sp/search';
import '@pnp/sp/security';
import { parseBinderWithErrorCheck, HttpRequestError } from '@pnp/queryable';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IBreachListItem } from './IBreachListItem';
import { IFormState } from '../components/IFormState';
import { SP_LIST_NAME, SP_ADMIN_GROUPS, SP_REVIEWER_GROUPS } from '../config';

let _sp: SPFI;

export function initSharePoint(context: WebPartContext): void {
  _sp = spfi().using(SPFx(context));
}

function formatDate(day: string, month: string, year: string): string | null {
  if (!day || !month || !year) return null;
  const d = parseInt(day, 10);
  const m = parseInt(month, 10);
  const y = parseInt(year, 10);
  if (isNaN(d) || isNaN(m) || isNaN(y)) return null;
  return new Date(y, m - 1, d).toISOString();
}

// Maps the React form state to SharePoint list column internal names.
export function mapFormToListItem(form: IFormState): IBreachListItem {
  return {
    Title:                              '',
    // Deliberately left blank at submission — whether a breach is reported
    // to the ICO is an admin decision made during Panel recommendations
    // review, not something the submitter answers.
    Reported_x0020_to_x0020_ICO:        '',

    Details_x0020_of_x0020_the_x0020:  form.officerName,
    Email_x0020_of_x0020_the_x0020_r:  form.officerEmail,
    Directorate_x0020_of_x0020_repor:  form.officerDirectorate,
    Division_x0020_of_x0020_reportin:  form.officerDivision,

    Name_x0020_of_x0020_person_x0020:  form.responsibleName,
    Email_x0020_address_x0020_of_x00:  form.responsibleEmail,
    Has_x0020_person_x0020_responsib:  form.responsibleTraining,
    Directorate_x0020_of_x0020_perso:  form.responsibleDirectorate,
    Division_x0020_of_x0020_person_x:  form.responsibleDivision,
    Line_x0020_manager_x0020_of_x002:  form.lineManager,
    Email_x0020_address_x0020_of_x000: form.lineManagerEmail,

    Details_x0020_of_x0020_the_x00200: form.breachDetails,
    Outline_x0020_of_x0020_data_x002: form.dataBreachedOutline,
    Approx_x0020_individuals_x0020_a:  form.approxIndividuals,
    Categorise_x0020_nature_x0020_of:  form.categoriseNature.join(';'),
    Categorise_x0020_nature_x0020_ot:  form.categoriseNatureOther,
    Categories_x0020_of_x0020_person:  form.categoriesPersonalData.join(';'),
    Personal_x0020_data_x0020_other:   form.categoriesPersonalDataOther,
    Categories_x0020_of_x0020_indivi:  form.categoriesIndividuals.join(';'),
    Individuals_x0020_other:           form.categoriesIndividualsOther,
    Impact_x0020_to_x0020_Data_x0020:  form.impactToSubjects,
    Breach_x0020_categories:           form.breachCategories.join(';'),
    Date_x0020_breach_x0020_discover:  formatDate(form.dateDiscoveredDay, form.dateDiscoveredMonth, form.dateDiscoveredYear),
    Time_x0020_breach_x0020_discover:  form.timeDiscovered,
    Date_x0020_breach_x0020_happened:  formatDate(form.dateHappenedDay, form.dateHappenedMonth, form.dateHappenedYear),
    Time_x0020_breach_x0020_happened:  form.timeHappened,

    Name_x0020_of_x0020_third_x0020_:  form.thirdPartyName,
    Address_x0020_of_x0020_third_x00:  form.thirdPartyAddress,
    Contact_x0020_for_x0020_third_x0:  form.thirdPartyContact,
    Email_x0020_for_x0020_third_x002:  form.thirdPartyEmail,
    Telephone_x0020_for_x0020_third_:  form.thirdPartyTelephone,

    Likely_x0020_consequences_x0020_:  form.likelyConsequences,
    Consequences_x0020_not_x0020_yet:  form.consequencesExplanation,
    Steps_x0020_taken_x0020_for_x002:  form.stepsTaken,
    Evidence_x0020_to_x0020_prevent_:  form.evidencePrevent,

    Have_x0020_affected_x0020_indivi:  form.individualsInformed,
    Details_x0020_of_x0020_informing:  form.individualsInformedDetails,
    Does_x0020_breach_x0020_relate_x:  form.thirdPartyAgreement,
    What_x0020_information_x0020_sha:  form.infoShared,
    Have_x0020_other_x0020_organisat:  form.otherOrgsInformed,
    Details_x0020_of_x0020_informing0: form.otherOrgsDetails,
    Has_x0020_media_x0020_covered_x0:  form.mediaCovered,
    Media_x0020_coverage_x0020_of_x0:  form.mediaDetails,
    Any_x0020_other_x0020_info:        form.anyOtherInfo,

    Panel_x0020_Recommendations:       form.panelRecommendations,
    Notification_x0020_to_x0020_ICO_:  form.notificationToIco,
    // Left blank at submission, same as Reported to ICO above — these are
    // admin decisions made during Panel recommendations review, not
    // something the submitter answers.
    Staff_x0020_e_x002d_Learning_x00:  '',
    Review_x0020_of_x0020_Procedures:  '',
    Reason_x0020_for_x0020_Decision:   form.reasonForDecision,
  };
}

// Accounts in many SharePoint/AD groups (typically admin/IT accounts) can make
// currentUser.groups() noticeably slower — and a stalled request here must
// never leave callers waiting forever, since it gates page/report loading.
const ADMIN_CHECK_TIMEOUT_MS = 10000;

// If the real check is still running when the timeout wins the race, it is
// NOT cancelled — it keeps going in the background. Pass onLateResult to get
// notified when it eventually settles, so callers can self-correct the UI
// (e.g. reveal admin-only features) without requiring a page refresh.
export async function checkIsAdmin(onLateResult?: (isAdmin: boolean) => void): Promise<boolean> {
  if (!_sp) return false;
  const check = (async (): Promise<boolean> => {
    try {
      const user = await _sp.web.currentUser();
      if (user.IsSiteAdmin) return true;
      const groups = await _sp.web.currentUser.groups();
      for (let i = 0; i < groups.length; i++) {
        if (SP_ADMIN_GROUPS.indexOf(groups[i].Title) !== -1) return true;
      }
      return false;
    } catch (e) {
      return false;
    }
  })();

  let timedOut = false;
  const timeout = new Promise<boolean>(function(resolve) {
    setTimeout(function() {
      timedOut = true;
      console.warn(`[SharePointService] checkIsAdmin did not resolve within ${ADMIN_CHECK_TIMEOUT_MS}ms — treating as non-admin until the real result arrives.`);
      resolve(false);
    }, ADMIN_CHECK_TIMEOUT_MS);
  });

  check.then(function(result) {
    if (timedOut) {
      console.info(`[SharePointService] checkIsAdmin resolved late with isAdmin=${result}.`);
      if (onLateResult) onLateResult(result);
    }
  }).catch(function() { /* check() already swallows its own errors */ });

  return Promise.race([check, timeout]);
}

// Same shape as checkIsAdmin, but for the reviewer tier (client-side
// reviewers: see every report, edit only Panel Recommendations — see
// syncItemAccess and updatePanelRecommendationsElevated). Deliberately
// doesn't also return true for admins — the two are mutually exclusive in
// the UI (an admin sees the full admin panel, not the narrower reviewer
// view), so callers should check isAdmin first and only fall back to this.
export async function checkIsReviewer(onLateResult?: (isReviewer: boolean) => void): Promise<boolean> {
  if (!_sp) return false;
  const check = (async (): Promise<boolean> => {
    try {
      const groups = await _sp.web.currentUser.groups();
      for (let i = 0; i < groups.length; i++) {
        if (SP_REVIEWER_GROUPS.indexOf(groups[i].Title) !== -1) return true;
      }
      return false;
    } catch (e) {
      return false;
    }
  })();

  let timedOut = false;
  const timeout = new Promise<boolean>(function(resolve) {
    setTimeout(function() {
      timedOut = true;
      console.warn(`[SharePointService] checkIsReviewer did not resolve within ${ADMIN_CHECK_TIMEOUT_MS}ms — treating as non-reviewer until the real result arrives.`);
      resolve(false);
    }, ADMIN_CHECK_TIMEOUT_MS);
  });

  check.then(function(result) {
    if (timedOut) {
      console.info(`[SharePointService] checkIsReviewer resolved late with isReviewer=${result}.`);
      if (onLateResult) onLateResult(result);
    }
  }).catch(function() { /* check() already swallows its own errors */ });

  return Promise.race([check, timeout]);
}

// SharePoint claims logins look like "i:0#.f|membership|user@domain.com" —
// the actual account identifier is everything after the last "|".
function stripClaimsPrefix(loginName: string): string {
  const idx = loginName.lastIndexOf('|');
  return idx !== -1 ? loginName.slice(idx + 1) : loginName;
}

export async function getCurrentUserEmail(): Promise<string> {
  if (!_sp) return '';
  try {
    const user = await _sp.web.currentUser();
    // SharePoint populates Email from the account's mailbox — accounts with
    // no mailbox provisioned (e.g. a cloud-only .onmicrosoft.com admin/test
    // account) can have this blank even though the account itself is valid.
    // Fall back to LoginName (the claims login, always present) and then
    // display name, so "sent by"/"closed by" stamps never end up blank.
    const loginName = user.LoginName as string;
    return (user.Email as string) || (loginName ? stripClaimsPrefix(loginName) : '') || (user.Title as string) || '';
  } catch (e) {
    return '';
  }
}

// Display name for comm log "closed by"/"requested by" stamps — Title here
// is the same User Information List value that can go stale/claims-encoded
// for the same reason documented on cleanCurrentUserName (profile not yet
// synced), so it's run through the same cleanup rather than trusted as-is.
export async function getCurrentUserName(): Promise<string> {
  if (!_sp) return '';
  try {
    const user = await _sp.web.currentUser();
    const loginName = user.LoginName as string;
    const email = (user.Email as string) || (loginName ? stripClaimsPrefix(loginName) : '');
    return cleanCurrentUserName(user.Title as string, email);
  } catch (e) {
    return '';
  }
}

export interface IBreachListRow {
  Id: number;
  Title: string;
  Data_x0020_Breach_x0020_ID: string;
  Status: string;
  Created: string;
  Author: { Title: string; EMail: string };
  Directorate_x0020_of_x0020_repor: string;
  Reported_x0020_to_x0020_ICO: string;
  Approx_x0020_individuals_x0020_a: string;
  Impact_x0020_to_x0020_Data_x0020: string;
  Breach_x0020_categories: string;
  Categorise_x0020_nature_x0020_of: string;
  Date_x0020_breach_x0020_discover: string;
  [key: string]: unknown;
}

// SharePoint REST caps an unpaged request at a default page size (100 items),
// so a plain query silently truncates once the list grows past that — the
// unfiltered admin view hits this far sooner than a per-user filtered view.
// Page through results by Id (ascending) until a short page signals the end.
const BREACH_ITEMS_PAGE_SIZE = 500;

export async function getBreachItems(adminMode: boolean, currentUserEmail: string): Promise<IBreachListRow[]> {
  if (!_sp) throw new Error('SharePoint service not initialised.');
  const list = _sp.web.lists.getByTitle(SP_LIST_NAME);
  const select = [
    'Id', 'Title', 'Data_x0020_Breach_x0020_ID', 'Status', 'Created',
    'Author/Title', 'Author/EMail',
    'Directorate_x0020_of_x0020_repor',
    'Reported_x0020_to_x0020_ICO',
    'Approx_x0020_individuals_x0020_a',
    'Impact_x0020_to_x0020_Data_x0020',
    'Breach_x0020_categories',
    'Categorise_x0020_nature_x0020_of',
    'Date_x0020_breach_x0020_discover',
  ];
  const expand = ['Author'];

  // Non-admins see reports where they're the submitter, the person
  // responsible, or the line manager — anyone named in the report, not just
  // whoever created it — so it shows up in their own "View Reports" list
  // rather than only being reachable via a direct link from a notification.
  const involvedFilter = `(Author/EMail eq '${currentUserEmail}'` +
    ` or Email_x0020_address_x0020_of_x00 eq '${currentUserEmail}'` +
    ` or Email_x0020_address_x0020_of_x000 eq '${currentUserEmail}')`;

  const all: IBreachListRow[] = [];
  let afterId = 0;
  for (;;) {
    const idFilter = `Id gt ${afterId}`;
    const filter = adminMode ? idFilter : `${idFilter} and ${involvedFilter}`;
    const page = await list.items
      .select(...select)
      .expand(...expand)
      .filter(filter)
      .orderBy('Id', true)
      .top(BREACH_ITEMS_PAGE_SIZE)() as IBreachListRow[];
    if (!page.length) break;
    all.push(...page);
    afterId = page[page.length - 1].Id;
    if (page.length < BREACH_ITEMS_PAGE_SIZE) break;
  }
  return all;
}

export async function getBreachItemById(id: number): Promise<IBreachListItem & { Id: number; Author: { Title: string; EMail: string }; Created: string; Data_x0020_Breach_x0020_ID: string; Status: string }> {
  if (!_sp) throw new Error('SharePoint service not initialised.');
  return _sp.web.lists.getByTitle(SP_LIST_NAME).items.getById(id)
    .select('*', 'Author/Title', 'Author/EMail')
    .expand('Author')() as Promise<IBreachListItem & { Id: number; Author: { Title: string; EMail: string }; Created: string; Data_x0020_Breach_x0020_ID: string; Status: string }>;
}

export interface IAttachmentInfo {
  FileName: string;
  ServerRelativeUrl: string;
}

export interface IAttachmentUpload {
  filename: string;
  contentType: string;
  // Raw base64 content, no "data:...;base64," prefix — same convention as
  // IEmailAttachment in IReportViewerProps.ts.
  contentBytes: string;
}

// Submitters only hold Add Items on the list, not Edit (see syncItemAccess's
// doc comment and createBreachRecord below) — even immediately after
// creating their own item, before any permission sync has run, a direct
// PnP item.attachmentFiles.add() 403s, because attaching to an item is an
// edit. Routes through an elevated APIM-fronted Logic App instead, which
// verifies the caller is the item's Author server-side before writing via
// its own elevated SharePoint connection — same pattern as
// appendCommLogEntryElevated.
export async function uploadAttachmentElevated(
  apiUrl: string,
  apiKey: string,
  itemId: number,
  callerEmail: string,
  attachment: IAttachmentUpload
): Promise<void> {
  const res = await fetch(apiUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Ocp-Apim-Subscription-Key': apiKey },
    body: JSON.stringify({ itemId, callerEmail, ...attachment }),
  });
  if (!res.ok) {
    throw new Error(`Failed to upload attachment "${attachment.filename}" (${res.status}).`);
  }
}

// Read-only listing for the report viewer's Attachments section. Named
// individuals/reviewers only hold Read on the item, which is sufficient —
// attachment listing/download doesn't need an elevated Logic App endpoint.
export async function getAttachmentsForItem(itemId: number): Promise<IAttachmentInfo[]> {
  if (!_sp) throw new Error('SharePoint service not initialised.');
  const files = await _sp.web.lists.getByTitle(SP_LIST_NAME).items.getById(itemId).attachmentFiles();
  return files.map(f => ({ FileName: f.FileName, ServerRelativeUrl: f.ServerRelativeUrl }));
}

const UPDATE_MAX_ATTEMPTS = 3;

// Current ETag for the item, used to make updateBreachRecord's write
// conditional. Only the header is needed — select('Id') keeps the read
// itself cheap.
async function getItemEtag(id: number): Promise<string | null> {
  try {
    return await (_sp.web.lists.getByTitle(SP_LIST_NAME).items.getById(id)
      .select('Id')
      .using(parseBinderWithErrorCheck(async (r: Response) => {
        await r.json();
        return r.headers.get('etag');
      }))() as Promise<string | null>);
  } catch {
    return null;
  }
}

// Conditions the write on the item's current ETag (same approach as
// appendCommLogEntry) so a genuine conflict — e.g. the comm log write and an
// auto status change landing on the same item close together, or two admins
// editing the same field around the same time — is detected deterministically
// via SharePoint's If-Match check, rather than relying on it happening to
// also surface as a 409 Save Conflict, which isn't guaranteed for every kind
// of collision. Safe to retry verbatim: every caller here writes a fixed set
// of new values rather than something computed from the current state, so a
// retry doesn't need to recompute anything, just re-attempt with a fresh ETag.
export async function updateBreachRecord(id: number, fields: Partial<IBreachListItem>): Promise<void> {
  if (!_sp) throw new Error('SharePoint service not initialised.');
  const item = _sp.web.lists.getByTitle(SP_LIST_NAME).items.getById(id);
  for (let attempt = 1; attempt <= UPDATE_MAX_ATTEMPTS; attempt++) {
    const etag = await getItemEtag(id);
    try {
      await item.update(fields, etag || '*');
      return;
    } catch (err) {
      const isConflict = err instanceof HttpRequestError && (err.status === 412 || err.status === 409);
      if (isConflict && attempt < UPDATE_MAX_ATTEMPTS) {
        await delay(150 + Math.floor(Math.random() * 250));
        continue;
      }
      throw err;
    }
  }
}

// SharePoint People Search's PreferredName managed property is sourced from
// the User Profile Service — if a person's profile hasn't finished
// syncing/being crawled yet (common for newer or rarely-logged-in accounts),
// it falls back to the raw claims-encoded account name instead of a real
// display name (e.g. "0#.f|membership|name@domain.com"). Title is a second,
// independently-populated managed property that's often still correct when
// PreferredName isn't, so it's tried first as a fallback; if both look
// claims-encoded, a readable name is derived from the email's local part
// instead — a raw claims string should never reach the form.
function looksClaimsEncoded(s: string): boolean {
  return s.indexOf('|') !== -1;
}
function nameFromEmail(email: string): string {
  const local = (email.split('@')[0] || '').replace(/[._]+/g, ' ').trim();
  return local.split(' ').filter(Boolean).map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
}
function cleanDisplayName(preferredName: string | undefined, title: string | undefined, email: string): string {
  if (preferredName && !looksClaimsEncoded(preferredName)) return preferredName;
  if (title && !looksClaimsEncoded(title)) return title;
  return nameFromEmail(email);
}
// Same claims-encoded fallback as cleanDisplayName, for callers that only
// have a single name value to check (e.g. SPFx's pageContext.user.displayName,
// sourced from the site's User Information List rather than People Search —
// a different source that can go stale/claims-encoded for the same
// profile-not-synced-yet reason, but was never routed through
// cleanDisplayName since it doesn't come from searchUsers).
export function cleanCurrentUserName(displayName: string | undefined, email: string): string {
  if (displayName && !looksClaimsEncoded(displayName)) return displayName;
  return nameFromEmail(email);
}

export async function searchUsers(query: string): Promise<{ displayName: string; email: string }[]> {
  if (!_sp || query.trim().length < 2) return [];
  try {
    const results = await _sp.search({
      Querytext: `${query.trim()}*`,
      SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
      SelectProperties: ['PreferredName', 'WorkEmail', 'Title'],
      RowLimit: 8,
    });
    type PeopleResult = { PreferredName?: string; WorkEmail?: string; Title?: string };
    const rows = results.PrimarySearchResults as unknown as PeopleResult[];
    return (rows || [])
      .filter(r => !!r.WorkEmail)
      .map(r => ({
        displayName: cleanDisplayName(r.PreferredName, r.Title, r.WorkEmail as string),
        email: r.WorkEmail || '',
      }));
  } catch {
    return [];
  }
}

// ── Communication log ────────────────────────────────────────────────────────
//
// Stored as a JSON array in the CommunicationLog column on the main breach
// list item. Each entry is one message in the thread (outbound, reply, or a
// status-change stamp). All entries are now written by the webpart itself —
// replies used to also arrive via a Logic App watching a mailbox for actual
// email replies, but that flow has been disabled, so every entry going
// forward is already clean plain text typed into the app.

export interface ICommLogEntry {
  direction: 'Outbound' | 'Reply' | 'StatusChange' | 'CaseClosed' | 'CaseReopened';
  sentBy: string;
  // Display name of the staff member behind sentBy, captured at the time a
  // staff-driven entry (Outbound/StatusChange/CaseClosed/CaseReopened) is
  // written — so the comm log and PDF can show "<name> on behalf of Data
  // Breach Notification Team" instead of either the individual's bare email
  // or the shared team address with no way to tell which reviewer/admin
  // actually acted. Not set on Reply entries — sentBy there is already the
  // named individual's own identity, shown as-is. Optional because entries
  // logged before this field existed won't have it; display code falls back
  // to sentBy (the email) in that case.
  sentByName?: string;
  sentTo: string;
  subject: string;
  message: string;
  timestamp: string;
}

export function stripHtmlIfPresent(s: string): string {
  // Some list columns are configured as rich text, so SharePoint wraps saved
  // values in <div class="ExternalClassXXXX">...</div> — strip that markup so
  // raw HTML doesn't leak into the UI or exports.
  const trimmed = (s || '').trim();
  if (trimmed.slice(0, 20).indexOf('<') === -1) return trimmed;
  return trimmed
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#\d+;/g, '')
    .replace(/[ \t]{2,}/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

export function parseCommLog(raw: string | null | undefined): ICommLogEntry[] {
  if (!raw || !raw.trim()) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed as ICommLogEntry[] : [];
  } catch {
    return [];
  }
}

export async function saveCommLog(id: number, entries: ICommLogEntry[]): Promise<void> {
  if (!_sp) return;
  await _sp.web.lists.getByTitle(SP_LIST_NAME).items.getById(id).update({
    CommunicationLog: JSON.stringify(entries),
  });
}

const COMM_LOG_MAX_ATTEMPTS = 3;

function delay(ms: number): Promise<void> {
  return new Promise(function(resolve) { setTimeout(resolve, ms); });
}

// Reads CommunicationLog together with the item's current ETag. The ETag is
// what lets the write in appendCommLogEntry be made conditional — SharePoint
// rejects the update (412) if the item changed since this read, rather than
// silently accepting it and clobbering whatever the other write added.
async function readCommLogWithEtag(id: number): Promise<{ log: string | null | undefined; etag: string | null }> {
  return _sp.web.lists.getByTitle(SP_LIST_NAME).items.getById(id)
    .select('CommunicationLog')
    .using(parseBinderWithErrorCheck(async (r: Response) => {
      const json = await r.json();
      return { log: json.CommunicationLog as string | null | undefined, etag: r.headers.get('etag') };
    }))() as Promise<{ log: string | null | undefined; etag: string | null }>;
}

// Fetches the latest CommunicationLog, appends the entry, and saves —
// conditioned on the item's ETag so a concurrent write (e.g. two admins
// acting on the same report within moments of each other) is detected and
// retried, instead of one silently overwriting the other's entry with no
// error on either side. Falls back to an unconditional write only if ETag
// capture itself fails, so a quirk in that mechanism degrades to the old
// best-effort behaviour rather than blocking the append entirely.
//
// extraFields lets a caller fold something like a Status change into this
// same request (e.g. auto-marking a case Pending when a message is sent) —
// writing it as a separate .update() call would race this one for the same
// item and trip SharePoint's save-conflict check.
export async function appendCommLogEntry(id: number, entry: ICommLogEntry, extraFields?: Partial<IBreachListItem>): Promise<ICommLogEntry[]> {
  if (!_sp) return [entry];
  const item = _sp.web.lists.getByTitle(SP_LIST_NAME).items.getById(id);

  for (let attempt = 1; attempt <= COMM_LOG_MAX_ATTEMPTS; attempt++) {
    let log: string | null | undefined;
    let etag: string | null = null;
    try {
      const read = await readCommLogWithEtag(id);
      log = read.log;
      etag = read.etag;
    } catch {
      const fresh = await item.select('CommunicationLog')() as { CommunicationLog?: string | null };
      log = fresh.CommunicationLog;
    }

    const updated = parseCommLog(log).concat(entry);

    try {
      await item.update({ CommunicationLog: JSON.stringify(updated), ...extraFields }, etag || '*');
      return updated;
    } catch (err) {
      const isConflict = err instanceof HttpRequestError && (err.status === 412 || err.status === 409);
      if (isConflict && attempt < COMM_LOG_MAX_ATTEMPTS) {
        // Someone else updated this item since we read it — back off briefly
        // and retry with a fresh read rather than overwriting their change.
        await delay(150 + Math.floor(Math.random() * 250));
        continue;
      }
      throw err;
    }
  }
  // Unreachable in practice — the loop above always returns or throws — but
  // keeps this a valid async function under strict return-type checking.
  throw new Error('Failed to save communication log entry.');
}

// Named individuals (submitter, person responsible) only hold Read on their
// own report item — not Edit — so they can't use appendCommLogEntry's direct
// item.update() (that's reserved for admins, who do have Edit rights at the
// item level). Instead this calls an elevated APIM-fronted Logic App that
// re-verifies the caller is actually the Author or person responsible for
// this specific item (fetched fresh server-side, not trusted from the
// request) before appending on their behalf — see
// logic-apps/Data_Breach_Append_Comm_Log for the workflow definition.
//
// The workflow only confirms the save; it deliberately doesn't return the
// log, so the endpoint can't be used to read another report's
// correspondence. The updated log is read back here under the caller's own
// SharePoint permissions instead — anyone allowed to append can already read
// the item.
export async function appendCommLogEntryElevated(
  apiUrl: string,
  apiKey: string,
  itemId: number,
  callerEmail: string,
  entry: ICommLogEntry
): Promise<ICommLogEntry[]> {
  const res = await fetch(apiUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Ocp-Apim-Subscription-Key': apiKey },
    body: JSON.stringify({ itemId, callerEmail, entry }),
  });
  if (!res.ok) {
    throw new Error(`Failed to save communication log entry (${res.status}).`);
  }
  // Earlier versions of the workflow returned the full log. Use it if it's
  // there, so this works whichever version is deployed.
  const body: unknown = await res.json().catch(function() { return undefined; });
  if (Array.isArray(body)) return body as ICommLogEntry[];
  return readCommLog(itemId);
}

// Reads a report's communication log under the current user's own permissions.
async function readCommLog(itemId: number): Promise<ICommLogEntry[]> {
  const item = await _sp.web.lists.getByTitle(SP_LIST_NAME).items.getById(itemId).select('CommunicationLog')();
  return parseCommLog(item.CommunicationLog as string);
}

// Named after the Panel recommendations section, but also covers Status —
// reviewers use the same elevated endpoint for both, since both are just a
// straightforward field overwrite (unlike CommunicationLog, which needs
// read-modify-write append semantics via appendCommLogEntryElevated).
export interface IPanelRecommendationsFields {
  Reported_x0020_to_x0020_ICO?: string;
  Staff_x0020_e_x002d_Learning_x00?: string;
  Review_x0020_of_x0020_Procedures?: string;
  Reason_x0020_for_x0020_Decision?: string;
  Status?: string;
}

// Reviewers only hold Read at the item level — see syncItemAccess — so
// editing the "Panel recommendations" section (Reported to ICO, e-Learning
// required, Procedures review required, Reason for decision) goes through
// this elevated endpoint instead of a direct item.update(). The Logic App
// verifies the caller is actually in the Reviewers or Owners SharePoint
// group server-side before writing, since this isn't scoped to a specific
// item the caller is named on (unlike appendCommLogEntryElevated) —
// reviewers can edit this section on any report.
export async function updatePanelRecommendationsElevated(
  apiUrl: string,
  apiKey: string,
  itemId: number,
  callerEmail: string,
  fields: IPanelRecommendationsFields
): Promise<void> {
  const res = await fetch(apiUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Ocp-Apim-Subscription-Key': apiKey },
    body: JSON.stringify({ itemId, callerEmail, fields }),
  });
  if (!res.ok) {
    throw new Error(`Failed to save Panel Recommendations (${res.status}).`);
  }
}


// ─────────────────────────────────────────────────────────────────────────────

function extractSpErrorMessage(err: unknown): string {
  if (err instanceof Error) return err.message;
  if (err && typeof err === 'object') {
    const e = err as Record<string, unknown>;
    // @pnp/sp wraps OData errors under .data['odata.error'].message.value
    if (e.data && typeof e.data === 'object') {
      const data = e.data as Record<string, unknown>;
      const odataErr = data['odata.error'] as Record<string, unknown> | undefined;
      if (odataErr && typeof odataErr.message === 'object' && odataErr.message !== null) {
        const msgObj = odataErr.message as Record<string, unknown>;
        if (typeof msgObj.value === 'string') return msgObj.value;
      }
    }
    if (typeof e.message === 'string') return e.message;
    try { return JSON.stringify(err); } catch { /* fall through */ }
  }
  return 'An unexpected error occurred. Check the browser console for details.';
}

// ── Per-item access control ─────────────────────────────────────────────────
//
// Makes "only people named on a report can open it" actually true at the
// SharePoint permission level, not just how the app's UI happens to filter
// the list. Each report's item-level permissions are broken from the list's
// inherited (broad) permissions and set to exactly: the admin group(s), plus
// whoever is currently named as submitter/reporting officer/person
// responsible/line manager. This only has any effect once the base list
// permission for regular staff no longer grants list-wide Read — as long as
// that broad grant remains, these narrower per-item grants are moot, since
// broader access always wins. See syncItemAccess's caller sites for when
// this gets (re-)applied.

// Cache the in-flight promise, not the resolved value — concurrent callers
// then share one request instead of racing to assign the same module-level
// variable after their own separate awaits.
let _readRoleDefIdPromise: Promise<number> | null = null;
let _adminRoleDefIdPromise: Promise<number> | null = null;
let _adminGroupIds: number[] | null = null;
let _reviewerGroupIds: number[] | null = null;

function getReadRoleDefId(): Promise<number> {
  if (!_readRoleDefIdPromise) {
    _readRoleDefIdPromise = _sp.web.roleDefinitions.getByName('Read')().then(function(roleDef) { return roleDef.Id; });
  }
  return _readRoleDefIdPromise;
}

// Full Contribute (includes Delete) for the admin group — unlike named
// individuals, admins may need to remove a report submitted in error. Named
// individuals never need Edit at the item level at all: writes to the
// communication log go through appendCommLogEntryElevated instead, which
// uses the elevated Logic App connection rather than the caller's own
// permissions — so Read is enough for everyone except admins.
function getAdminRoleDefId(): Promise<number> {
  if (!_adminRoleDefIdPromise) {
    _adminRoleDefIdPromise = _sp.web.roleDefinitions.getByName('Contribute')().then(function(roleDef) { return roleDef.Id; });
  }
  return _adminRoleDefIdPromise;
}

async function resolveGroupIds(groupNames: string[], label: string): Promise<number[]> {
  const ids: number[] = [];
  for (let i = 0; i < groupNames.length; i++) {
    try {
      const group = await _sp.web.siteGroups.getByName(groupNames[i])();
      ids.push(group.Id);
    } catch (e) {
      // Group not found, or this account can't resolve it — skip rather than
      // fail the whole sync, so a config typo doesn't lock everyone out.
      console.error(`[SharePointService] Could not resolve ${label} group "${groupNames[i]}" for item permissions:`, e);
    }
  }
  return ids;
}

async function getReviewerGroupIds(): Promise<number[]> {
  if (_reviewerGroupIds !== null) return _reviewerGroupIds;
  _reviewerGroupIds = await resolveGroupIds(SP_REVIEWER_GROUPS, 'reviewer');
  return _reviewerGroupIds;
}

async function getAdminGroupIds(): Promise<number[]> {
  if (_adminGroupIds !== null) return _adminGroupIds;
  const ids: number[] = await resolveGroupIds(SP_ADMIN_GROUPS, 'admin');
  _adminGroupIds = ids;
  return ids;
}

// Grants item-level access to exactly two tiers: the admin group(s)
// ("Data Breach Owners" — developers/full control) at Contribute (includes
// Delete — admins may need to remove a report submitted in error), and
// everyone else at Read — the given involvedEmails (submitter, reporting
// officer, person responsible, line manager) plus the reviewer group(s)
// ("Data Breach Reviewers" — client-side reviewers, see every report). None
// of them need Edit at the item level: named-individual replies go through
// appendCommLogEntryElevated, and reviewer edits to Panel Recommendations go
// through a similar elevated endpoint — both write under the Logic App's own
// elevated connection rather than the caller's permissions. Nobody else gets
// access, even if they can otherwise browse the site. Revokes anyone
// currently granted who isn't in either set anymore (e.g. after the person
// responsible is reassigned), so access stays accurate as those fields
// change, not just additive over time. Best-effort throughout: a failure to
// resolve or grant one person logs and continues rather than aborting, since
// the caller (report submission or a section save) must still succeed even
// if a permissions sync hiccups.
//
// isClosed drops involvedEmails from the grant entirely once a report is
// Closed — only the admin group and the reviewer group keep access, so a
// closed report doesn't sit forever with unique permissions naming a
// submitter/officer/responsible/line-manager who no longer has a live reason
// to hold access. The revoke loop below already strips anyone not in the
// resulting keep set, so passing isClosed=true is enough to pull those grants
// back on the very next sync — no separate removal path needed. The
// server-side permission-sync Logic App (external, not tracked in this repo)
// mirrors this same isClosed rule on its own polling trigger, since that's
// the only path guaranteed to run for a Reviewer-driven close (Reviewers
// don't hold Manage Permissions, so this client-side call is a no-op 403 for
// them — see the breakRoleInheritance catch below).
export async function syncItemAccess(
  itemId: number,
  involvedEmails: Array<string | undefined>,
  isClosed: boolean = false
): Promise<void> {
  if (!_sp) return;
  const item = _sp.web.lists.getByTitle(SP_LIST_NAME).items.getById(itemId);

  try {
    await item.breakRoleInheritance(false, true);
  } catch (e) {
    // A 403 here means the current account doesn't have Manage Permissions
    // on this list — the single most likely reason this whole feature would
    // silently not work (e.g. for a regular submitter's own session). Any
    // other error is more likely "already unique," which is harmless and
    // expected on a re-sync, so only 403 is surfaced as a clear warning.
    if (e instanceof HttpRequestError && e.status === 403) {
      console.warn(`[SharePointService] Access denied breaking permission inheritance on item ${itemId} — this account likely lacks Manage Permissions on the list, so item-level access grants will not take effect.`, e);
    }
  }

  let readRoleDefId: number;
  let adminRoleDefId: number;
  try {
    [readRoleDefId, adminRoleDefId] = await Promise.all([getReadRoleDefId(), getAdminRoleDefId()]);
  } catch (e) {
    console.error('[SharePointService] Could not resolve a permission level — aborting item permission sync:', e);
    return;
  }
  const adminGroupIds = await getAdminGroupIds();
  const reviewerGroupIds = await getReviewerGroupIds();

  const seen: { [emailLower: string]: boolean } = {};
  const resolveIds = async function(emails: Array<string | undefined>): Promise<number[]> {
    const ids: number[] = [];
    for (let i = 0; i < emails.length; i++) {
      const email = (emails[i] || '').trim();
      if (!email || seen[email.toLowerCase()]) continue;
      seen[email.toLowerCase()] = true;
      try {
        const user = await _sp.web.ensureUser(email);
        ids.push(user.Id);
      } catch (e) {
        console.error(`[SharePointService] Could not resolve user for item permissions: ${email}`, e);
      }
    }
    return ids;
  };

  const readOnlyIds = (isClosed ? [] : await resolveIds(involvedEmails)).concat(reviewerGroupIds);

  const keepIds: { [id: number]: boolean } = {};
  adminGroupIds.forEach(function(id) { keepIds[id] = true; });
  readOnlyIds.forEach(function(id) { keepIds[id] = true; });

  try {
    const current = await item.roleAssignments.select('PrincipalId')() as Array<{ PrincipalId: number }>;
    for (let i = 0; i < current.length; i++) {
      const principalId = current[i].PrincipalId;
      if (keepIds[principalId]) continue;
      // Revoked principals could currently hold either level — remove both;
      // whichever one they didn't hold is expected to no-op.
      try { await item.roleAssignments.remove(principalId, readRoleDefId); } catch { /* not held at Read */ }
      try { await item.roleAssignments.remove(principalId, adminRoleDefId); } catch { /* not held at Contribute */ }
    }
  } catch (e) {
    console.error(`[SharePointService] Failed to read current permissions on item ${itemId}:`, e);
  }

  let grantedCount = 0;
  const grant = async function(ids: number[], roleDefId: number): Promise<void> {
    for (let i = 0; i < ids.length; i++) {
      try {
        await item.roleAssignments.add(ids[i], roleDefId);
        grantedCount++;
      } catch (e) {
        console.error(`[SharePointService] Failed to grant principal ${ids[i]} access to item ${itemId}:`, e);
      }
    }
  };
  await grant(adminGroupIds, adminRoleDefId);
  await grant(readOnlyIds, readRoleDefId);

  // Explicit success signal — without this, "nothing in the console" is
  // ambiguous between "it worked" and "it hasn't finished yet."
  console.info(`[SharePointService] Item permission sync complete for report ${itemId}: ${grantedCount} principals granted (${adminGroupIds.length} admin group(s) at Contribute, ${readOnlyIds.length} named individual(s)/reviewer group(s) at Read).`);
}

export async function createBreachRecord(form: IFormState): Promise<number> {
  if (!_sp) throw new Error('SharePoint service not initialised. Call initSharePoint() first.');
  const item = mapFormToListItem(form);
  try {
    const result = await _sp.web.lists.getByTitle(SP_LIST_NAME).items.add(item);
    const newId = result.Id as number;
    try {
      const created = await getBreachItemById(newId);
      await syncItemAccess(newId, [
        created.Author ? created.Author.EMail : undefined,
        form.officerEmail,
        form.responsibleEmail,
        form.lineManagerEmail,
      ]);
    } catch (permErr) {
      // The report itself was created successfully — a permissions sync
      // failure shouldn't be reported as a failed submission. Worst case,
      // the named individuals can't see it yet until this is retried (e.g.
      // via the next edit), which is a safe-by-default failure direction.
      console.error(`[SharePointService] Failed to set item-level permissions for report ${newId}:`, permErr);
    }
    return newId;
  } catch (err) {
    console.error('SharePoint write error:', err);
    throw new Error(extractSpErrorMessage(err));
  }
}

