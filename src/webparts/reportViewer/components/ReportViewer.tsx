import * as React from 'react';
import type { IReportViewerProps, IEmailAttachment } from './IReportViewerProps';
import { buildReportPdfBase64, IReportSection } from '../services/ReportPdf';
import {
  checkIsAdmin,
  checkIsReviewer,
  getCurrentUserEmail,
  getCurrentUserName,
  getBreachItems,
  getBreachItemById,
  updateBreachRecord,
  parseCommLog,
  appendCommLogEntry,
  stripHtmlIfPresent,
  syncItemAccess,
  getAttachmentsForItem,
  IBreachListRow,
  ICommLogEntry,
  IPanelRecommendationsFields,
  IAttachmentInfo,
} from '../../dataBreachForm/services/SharePointService';
import { IBreachListItem } from '../../dataBreachForm/services/IBreachListItem';
import { buildRequestInfoEmail, buildUserReplyEmail, buildClosedEmail, buildReopenedEmail, buildCommunicationNoticeEmail, BreachRole } from '../../dataBreachForm/services/EmailTemplates';
import { ADMIN_NOTIFICATION_EMAIL, TEAM_NAME } from '../../dataBreachForm/config';
import { FullItem, EditingState } from './reportViewerTypes';
import { formatDate, getRef, YES_NO, DEFAULT_REASON_FOR_DECISION, getReportFieldSections } from './reportViewerUtils';
import { printReport, downloadSingleCsv } from './reportExport';
import { Field, Section } from './SummaryList';
import ReportList, { useReportListControls } from './ReportList';
import { CloseCaseDialog, ReopenCaseDialog } from './StatusDialogs';
import CommLogEntries from './CommLogEntries';
import SubmittedReport from './SubmittedReport';

// ── Main component ───────────────────────────────────────────────────────────

const ReportViewer: React.FC<IReportViewerProps> = ({ spContext, sharedMailbox = '', sendEmailFromMailbox, targetItemId, onTargetItemOpened, testMode = false, appendCommLogEntryElevated, updatePanelRecommendationsElevated }) => {
  const [isAdmin, setIsAdmin] = React.useState(false);
  const [isReviewer, setIsReviewer] = React.useState(false);
  const [currentUserEmail, setCurrentUserEmail] = React.useState('');
  const [currentUserName, setCurrentUserName] = React.useState('');
  const [items, setItems] = React.useState<IBreachListRow[]>([]);
  const [selectedItem, setSelectedItem] = React.useState<FullItem | null>(null);
  const [loading, setLoading] = React.useState(true);
  const [loadingDetail, setLoadingDetail] = React.useState(false);
  const [detailError, setDetailError] = React.useState('');
  const [editingSection, setEditingSection] = React.useState<string | null>(null);
  const [editingState, setEditingState] = React.useState<EditingState>({});
  const [saving, setSaving] = React.useState(false);
  const [saveError, setSaveError] = React.useState('');
  // Surfaced whenever a message/notice sent fine but writing it to the
  // Communication Log itself failed — deliberately separate from saveError/
  // sendError/userReplyError, which all live inside panels that get closed
  // right before the log write is attempted, so an error set on those would
  // never actually be visible to the user.
  const [commLogWarning, setCommLogWarning] = React.useState('');
  // Set when "Mark as Closed" is blocked by incomplete Panel recommendations
  // — draws the eye to that section (highlighted border) instead of leaving
  // people to work out from the error text alone where to go fix it.
  const [highlightPanelRecommendations, setHighlightPanelRecommendations] = React.useState(false);
  const panelRecommendationsRef = React.useRef<HTMLDivElement>(null);
  // "Mark as Closed" opens this confirmation instead of closing immediately —
  // closing sends a final notification email to the submitter and (per
  // current design) can't be undone through the UI, since Reopen is disabled.
  const [showCloseConfirm, setShowCloseConfirm] = React.useState(false);
  // "Reopen" opens this confirmation instead of reopening immediately, same
  // reasoning as showCloseConfirm — reopening sends its own notification
  // email and the optional message field needs somewhere to be typed first.
  const [showReopenConfirm, setShowReopenConfirm] = React.useState(false);
  const [reopenMessage, setReopenMessage] = React.useState('');
  // Search/sort/paging for the list view — kept here so it survives opening a report.
  const listControls = useReportListControls(items);
  const [showRequestInfo, setShowRequestInfo] = React.useState(false);
  const [requestInfoText, setRequestInfoText] = React.useState('');
  const [sending, setSending] = React.useState(false);
  const [sendSuccess, setSendSuccess] = React.useState(false);
  const [sendError, setSendError] = React.useState('');
  const [commLog, setCommLog] = React.useState<ICommLogEntry[]>([]);
  const [refreshingLog, setRefreshingLog] = React.useState(false);
  const [attachments, setAttachments] = React.useState<IAttachmentInfo[]>([]);
  const [showUserReply, setShowUserReply] = React.useState(false);
  const [userReplyText, setUserReplyText] = React.useState('');
  const [userReplySending, setUserReplySending] = React.useState(false);
  const [userReplySent, setUserReplySent] = React.useState(false);
  const [userReplyError, setUserReplyError] = React.useState('');

  // Requests for more information go to the person responsible for the
  // breach, not the submitter — they're often different people, and the
  // person responsible is who's actually being asked to explain/clarify.
  // Falls back to the submitter if no responsible-person email is on record
  // (which also covers the case where they're the same person).
  const getRequestInfoRecipient = (): { name: string; email: string } => {
    if (!selectedItem) return { name: 'Colleague', email: '' };
    const responsibleName  = selectedItem.Name_x0020_of_x0020_person_x0020;
    const responsibleEmail = selectedItem.Email_x0020_address_x0020_of_x00;
    if (responsibleEmail && responsibleEmail.trim()) {
      return { name: responsibleName || 'Colleague', email: responsibleEmail };
    }
    return {
      name: selectedItem.Author ? selectedItem.Author.Title : 'Colleague',
      email: selectedItem.Author ? selectedItem.Author.EMail : '',
    };
  };

  // True only if the signed-in user is the person responsible — the reply
  // action is specifically their job, not the submitter's or line manager's,
  // even though both are named on the report. If a submitter happens to
  // also be the responsible person, this still returns true via the email
  // match — it's the role that matters, not which field they came from.
  // Deliberately independent of isAdmin: an admin who also happens to be the
  // responsible person for this specific report must still be able to see
  // and use the reply flow for it.
  const isCurrentUserRecipient = (): boolean => {
    if (!selectedItem || !currentUserEmail) return false;
    const email = currentUserEmail.toLowerCase();
    const responsibleEmail = (selectedItem.Email_x0020_address_x0020_of_x00 || '').toLowerCase();
    return !!responsibleEmail && email === responsibleEmail;
  };

  // Broader than isCurrentUserRecipient: also includes the submitter and
  // line manager, who should be able to view the report and its
  // communication history since they're named in it, but don't get the
  // "reply" action — that's the responsible person's job specifically.
  const isCurrentUserInvolved = (): boolean => {
    if (isCurrentUserRecipient()) return true;
    if (!selectedItem || !currentUserEmail) return false;
    const email = currentUserEmail.toLowerCase();
    const authorEmail = selectedItem.Author ? selectedItem.Author.EMail.toLowerCase() : '';
    const lineManagerEmail = (selectedItem.Email_x0020_address_x0020_of_x000 || '').toLowerCase();
    return (!!authorEmail && email === authorEmail) || (!!lineManagerEmail && email === lineManagerEmail);
  };

  // Bystanders to keep informed of back-and-forth communication — submitter
  // and line manager, but deliberately NOT the person responsible. They're a
  // primary party to the case, not a bystander: they either directly receive
  // the request or directly send the reply via the real request/reply
  // emails, so a passive "FYI, something happened" copy on top of that is
  // redundant at best and wrongly implies "no action needed" at worst.
  // Normally deduplicated by email, tracking every role a person holds (the
  // same person is often more than one of these) so notices can state
  // exactly why each recipient is getting one. In test mode, roles are
  // deliberately NOT merged even when they share an email — realistically
  // they'd always be different people, but while testing with one mailbox
  // across both roles, sending each tailored variant separately lets you
  // review both rather than getting one merged email.
  const getInvolvedPeople = (): Array<{ name: string; email: string; roles: BreachRole[] }> => {
    if (!selectedItem) return [];
    const candidates: Array<{ name: string; email: string; role: BreachRole }> = [
      { name: selectedItem.Author ? selectedItem.Author.Title : '', email: selectedItem.Author ? selectedItem.Author.EMail : '', role: 'submitter' },
      { name: selectedItem.Line_x0020_manager_x0020_of_x002 || '', email: selectedItem.Email_x0020_address_x0020_of_x000 || '', role: 'lineManager' },
    ];
    if (testMode) {
      return candidates
        .filter(function(c) { return !!c.email.trim(); })
        .map(function(c) { return { name: c.name, email: c.email, roles: [c.role] }; });
    }
    const result: Array<{ name: string; email: string; roles: BreachRole[] }> = [];
    candidates.forEach(function(c) {
      const key = c.email.toLowerCase().trim();
      if (!key) return;
      const existing = result.filter(function(r) { return r.email.toLowerCase().trim() === key; })[0];
      if (existing) {
        existing.roles.push(c.role);
      } else {
        result.push({ name: c.name, email: c.email, roles: [c.role] });
      }
    });
    return result;
  };

  // Keeps everyone named on the report in the loop whenever a message goes
  // back and forth, even when they aren't the direct sender/recipient of
  // that particular message. Best-effort and fire-and-forget: these are
  // informational copies, so a failure here shouldn't block or error out the
  // actual request/reply that already sent successfully.
  const sendCommunicationNotices = (opts: { excludeEmail: string; fromName: string; toName: string; message: string; direction: 'request' | 'reply' }): void => {
    if (!sendEmailFromMailbox || !selectedItem) return;
    const ref = getRef(selectedItem.Id);
    const itemId = selectedItem.Id;
    const siteUrl = spContext.pageContext.web.absoluteUrl;
    const exclude = opts.excludeEmail.toLowerCase().trim();
    getInvolvedPeople()
      .filter(function(p) { return testMode || p.email.toLowerCase().trim() !== exclude; })
      .forEach(function(p) {
        // Tag the subject with the role in test mode so several tailored
        // copies landing in the same test mailbox are easy to tell apart.
        const roleTag = testMode ? ` [${p.roles.join('+')}]` : '';
        const subject = `Data Breach Report ${ref} - Communication Update${roleTag}`;
        const body = buildCommunicationNoticeEmail({
          ref, itemId, siteUrl, recipientName: p.name || 'Colleague', recipientRoles: p.roles,
          fromName: opts.fromName, toName: opts.toName, message: opts.message, direction: opts.direction,
        });
        sendEmailFromMailbox!(p.email, subject, body).catch(function(err: Error) {
          console.error(`[ReportViewer] Failed to send communication notice to ${p.email}:`, err);
        });
      });
  };

  const buildEmailContent = (): { to: string; subject: string; body: string } => {
    const ref       = selectedItem ? getRef(selectedItem.Id) : '';
    const itemId    = selectedItem ? selectedItem.Id : 0;
    const recipient = getRequestInfoRecipient();
    const siteUrl   = spContext.pageContext.web.absoluteUrl;
    // This is the email most likely to prompt a reflexive "just hit reply" —
    // it's a direct ask for information — so the subject itself flags not to.
    const subject   = `Data Breach Report ${ref} - Additional Information Required (do not reply to this email)`;
    const body      = buildRequestInfoEmail({ ref, itemId, siteUrl, recipientName: recipient.name, message: requestInfoText });
    return { to: recipient.email, subject, body };
  };

  const buildEmailPreview = (): string => {
    const name    = getRequestInfoRecipient().name;
    const siteUrl = spContext.pageContext.web.absoluteUrl;
    return [
      `Dear ${name},`,
      '',
      'We require some more information regarding a data breach report you\'re involved in:',
      '',
      requestInfoText || '(your message will appear here)',
      '',
      'Please reply to this email with any additional information you can provide.',
      '',
      'You can view your previous data breach reports, or submit a new one here:',
      siteUrl,
    ].join('\n');
  };

  const handleSendEmail = (): void => {
    if (!requestInfoText.trim()) return;
    setSending(true);
    setSendError('');
    setCommLogWarning('');
    const { to, subject, body } = buildEmailContent();
    const msgBody = requestInfoText;

    const logEntry = (): void => {
      const newEntry: ICommLogEntry = {
        direction: 'Outbound',
        sentBy: currentUserEmail,
        sentByName: currentUserName,
        sentTo: to,
        subject,
        message: msgBody,
        timestamp: new Date().toISOString(),
      };
      setSending(false);
      setSendSuccess(true);
      setShowRequestInfo(false);
      setRequestInfoText('');
      if (selectedItem) {
        const itemId = selectedItem.Id;
        // Requesting information implies the case is now awaiting a
        // response, so move it out of Open automatically. Leaves
        // Pending/Closed cases alone — closed cases should be reopened
        // explicitly, not silently by this. Admins get this folded into the
        // same request as the comm log write so the two can't race each
        // other; reviewers only hold Read at the item level, so it has to be
        // two separate elevated calls instead — status first, then the log
        // entry — same sequencing sendClosedNotification already relies on.
        const markPending = selectedItem.Status !== 'Pending' && selectedItem.Status !== 'Closed';
        const useElevatedRequestInfo = !isAdmin && isReviewer && appendCommLogEntryElevated && updatePanelRecommendationsElevated;
        const savePromise: Promise<ICommLogEntry[]> = useElevatedRequestInfo
          ? (markPending ? updatePanelRecommendationsElevated(itemId, currentUserEmail, { Status: 'Pending' }) : Promise.resolve())
              .then(function() { return appendCommLogEntryElevated(itemId, currentUserEmail, newEntry); })
          : appendCommLogEntry(itemId, newEntry, markPending ? { Status: 'Pending' } as Partial<IBreachListItem> : undefined);
        savePromise
          .then(function(updated) {
            setCommLog(updated);
            if (markPending) {
              setSelectedItem(function(prev) { return prev && prev.Id === itemId ? { ...prev, Status: 'Pending' } : prev; });
              setItems(function(prev) {
                return prev.map(function(r) { return r.Id === itemId ? { ...r, Status: 'Pending' } : r; });
              });
            }
          })
          .catch(function(err: Error) {
            console.error('[ReportViewer] Failed to save communication log entry:', err);
            setCommLog(commLog.concat(newEntry));
            setCommLogWarning('Your message was sent, but we could not confirm it was saved to this report\'s communication log. Refresh the page and check the log below — if it\'s still missing, try sending again.');
          });
      }
      sendCommunicationNotices({
        excludeEmail: to,
        fromName: sharedMailbox || TEAM_NAME,
        toName: getRequestInfoRecipient().name,
        message: msgBody,
        direction: 'request',
      });
    };

    if (sendEmailFromMailbox) {
      sendEmailFromMailbox(to, subject, body).then(logEntry).catch(function(err: Error) {
        setSending(false);
        setSendError(err && err.message ? err.message : 'Failed to send. Please try again.');
      });
    } else {
      setSending(false);
      setSendError('Email sending is not configured. Please contact your administrator.');
    }
  };

  const refreshCommLog = (): void => {
    if (!selectedItem || refreshingLog) return;
    setRefreshingLog(true);
    getBreachItemById(selectedItem.Id)
      .then(function(item) {
        setCommLog(parseCommLog(item.CommunicationLog));
        setRefreshingLog(false);
      })
      .catch(function(err: Error) {
        console.error('[ReportViewer] Failed to refresh communication log:', err);
        setRefreshingLog(false);
      });
  };

  const handleUserReply = (): void => {
    if (!userReplyText.trim() || !selectedItem) return;
    setUserReplySending(true);
    setUserReplyError('');
    setCommLogWarning('');
    const ref = getRef(selectedItem.Id);
    let lastOutboundSentBy = '';
    for (let i = commLog.length - 1; i >= 0; i--) {
      // CaseReopened counts as an outbound message too — sendReopenedNotification
      // sends it with the same From-reviewer/To-responsible shape as a plain
      // request for information (see the comment there), so a reply should
      // route back to whichever reviewer reopened the case, same as it would
      // for a regular request.
      if ((commLog[i].direction === 'Outbound' || commLog[i].direction === 'CaseReopened') && commLog[i].sentBy) {
        lastOutboundSentBy = commLog[i].sentBy;
        break;
      }
    }
    // Falls back to the central team mailbox (not the property-pane
    // sharedMailbox, which is display-only and may be blank) when nobody's
    // emailed this submitter yet,
    // so there's always somewhere for the reply to land.
    const sendTo = lastOutboundSentBy || ADMIN_NOTIFICATION_EMAIL;
    const siteUrl = spContext.pageContext.web.absoluteUrl;
    // The person actually sending this reply is always the person
    // responsible (isCurrentUserRecipient gates the "Submit more
    // information" action to them specifically) — not selectedItem.Author,
    // who is whoever originally filed the report and is frequently a
    // different person. Falls back to their email if the responsible-person
    // name field is somehow blank.
    const responsibleName = selectedItem.Name_x0020_of_x0020_person_x0020 || currentUserEmail;
    const emailBody = buildUserReplyEmail({
      ref,
      itemId: selectedItem.Id,
      siteUrl,
      submittedByName: responsibleName,
      submittedByEmail: currentUserEmail,
      responsibleName,
      message: userReplyText,
    });
    const sendFn = sendEmailFromMailbox;
    if (!sendFn) {
      setUserReplySending(false);
      setUserReplyError('Email sending is not configured. Your reply has not been sent — please contact your administrator.');
      return;
    }
    // Matches buildUserReplyEmail's own heading rather than the old "RE: ..."
    // reply-style subject — this was never a real email reply thread (see
    // the do-not-reply notice baked into every one of these templates), so
    // "RE:" implied a threading relationship that doesn't exist.
    const subject = `Response Received - Data Breach Report ${ref}`;
    sendFn(sendTo, subject, emailBody)
      .then(function() {
        const newEntry: ICommLogEntry = {
          direction: 'Reply',
          sentBy: currentUserEmail,
          sentTo: sendTo,
          subject,
          message: userReplyText,
          timestamp: new Date().toISOString(),
        };
        setUserReplySending(false);
        setUserReplySent(true);
        setShowUserReply(false);
        setUserReplyText('');
        if (selectedItem) {
          const save = appendCommLogEntryElevated
            ? appendCommLogEntryElevated(selectedItem.Id, currentUserEmail, newEntry)
            : appendCommLogEntry(selectedItem.Id, newEntry);
          save
            .then(function(updated) { setCommLog(updated); })
            .catch(function(err: Error) {
              console.error('[ReportViewer] Failed to save communication log entry:', err);
              setCommLog(commLog.concat(newEntry));
              setCommLogWarning('Your reply was sent, but we could not confirm it was saved to this report\'s communication log. Refresh the page and check the log below — if it\'s still missing, try sending again.');
            });
        }
        sendCommunicationNotices({
          excludeEmail: currentUserEmail,
          fromName: responsibleName,
          toName: TEAM_NAME,
          message: userReplyText,
          direction: 'reply',
        });
        // sendTo is whichever reviewer/admin last emailed this submitter
        // (their personal address, not the shared inbox), so the central
        // team mailbox otherwise never learns a reply came in unless that
        // one person happens to be watching for it. Mirrors how
        // sendClosedNotification always includes ADMIN_NOTIFICATION_EMAIL
        // as a fixed recipient. Skipped when sendTo already is the shared
        // mailbox (the lastOutboundSentBy fallback case) to avoid a
        // duplicate copy of the same email.
        if (sendTo.toLowerCase().trim() !== ADMIN_NOTIFICATION_EMAIL.toLowerCase().trim()) {
          sendFn(ADMIN_NOTIFICATION_EMAIL, subject, emailBody).catch(function(err: Error) {
            console.error('[ReportViewer] Failed to notify data breach notification mailbox of user reply:', err);
          });
        }
      })
      .catch(function(err: Error) {
        setUserReplySending(false);
        setUserReplyError((err && err.message) ? err.message : 'Failed to send reply. Please try again.');
      });
  };

  React.useEffect(() => {
    // getCurrentUserEmail() is a single lightweight call and always resolves
    // well before checkIsAdmin()'s/checkIsReviewer()'s timeout could fire, so
    // by the time a late result arrives, resolvedEmail is guaranteed to be
    // populated. Reviewers get the same unfiltered listing as admins — both
    // see every report, they just differ in what they can edit.
    let resolvedEmail = '';
    let resolvedAdmin = false;
    let resolvedReviewer = false;
    const refreshList = function(): void {
      getBreachItems(resolvedAdmin || resolvedReviewer, resolvedEmail)
        .then(function(rows) { setItems(rows); })
        .catch(function() { /* keep whatever list is already showing */ });
    };
    const applyLateAdminResult = function(admin: boolean): void {
      resolvedAdmin = admin;
      setIsAdmin(admin);
      refreshList();
    };
    const applyLateReviewerResult = function(reviewer: boolean): void {
      resolvedReviewer = reviewer;
      setIsReviewer(reviewer);
      refreshList();
    };

    Promise.all([checkIsAdmin(applyLateAdminResult), checkIsReviewer(applyLateReviewerResult), getCurrentUserEmail(), getCurrentUserName()])
      .then(function(results) {
        const admin = results[0];
        const reviewer = results[1];
        const email = results[2];
        const name = results[3];
        resolvedEmail = email;
        resolvedAdmin = admin;
        resolvedReviewer = reviewer;
        setIsAdmin(admin);
        setIsReviewer(reviewer);
        setCurrentUserEmail(email);
        setCurrentUserName(name);
        return getBreachItems(admin || reviewer, email);
      })
      .then(function(rows) {
        setItems(rows);
        setLoading(false);
        if (targetItemId) {
          openItem(targetItemId);
          // Consume the target so a later unmount/remount (e.g. switching
          // between Home and View Reports) lands on the list, not the report.
          if (onTargetItemOpened) onTargetItemOpened();
        }
      })
      .catch(function() { setLoading(false); });
  }, []);

  const openItem = (id: number): void => {
    setLoadingDetail(true);
    setSelectedItem(null);
    setDetailError('');
    setEditingSection(null);
    setEditingState({});
    setSaveError('');
    setSendSuccess(false);
    setShowRequestInfo(false);
    setRequestInfoText('');
    setCommLog([]);
    setShowUserReply(false);
    setUserReplyText('');
    setUserReplySending(false);
    setUserReplySent(false);
    setUserReplyError('');
    setAttachments([]);
    getAttachmentsForItem(id)
      .then(setAttachments)
      .catch(function(err: Error) { console.error('[ReportViewer] Failed to load attachments:', err); });
    getBreachItemById(id)
      .then(function(item) {
        setSelectedItem(item);
        setCommLog(parseCommLog(item.CommunicationLog));
        setLoadingDetail(false);

        // Regular submitters' own sessions can't set item permissions (no
        // Manage Permissions right), so the sync at submission time silently
        // can't take effect. Admins are notified of every new report by
        // email and typically open it soon after — re-running the sync here
        // (under the admin's own, more privileged session) narrows the
        // access gap from "possibly not until the case is closed" down to
        // "as soon as an admin checks the new report."
        if (isAdmin) {
          syncItemAccess(item.Id, [
            item.Author ? item.Author.EMail : undefined,
            item.Email_x0020_of_x0020_the_x0020_r,
            item.Email_x0020_address_x0020_of_x00,
            item.Email_x0020_address_x0020_of_x000,
          ], item.Status === 'Closed').catch(function(err: Error) {
            console.error('[ReportViewer] Failed to sync item permissions on open:', err);
          });
        }
      })
      .catch(function(err: Error) {
        console.error('[ReportViewer] Failed to load report detail:', err);
        setLoadingDetail(false);
        setDetailError(`This report has been closed, accessing this report is now unavailable — please check your inbox for the closure email with the report summary PDF. If you believe this is a mistake please contact us at ${ADMIN_NOTIFICATION_EMAIL}.`);
      });
  };

  const startEdit = (sectionKey: string): void => {
    const seed: EditingState = {};
    if (sectionKey === 'review' && !f('Reason_x0020_for_x0020_Decision')) {
      seed.Reason_x0020_for_x0020_Decision = DEFAULT_REASON_FOR_DECISION;
    }
    setEditingState(seed);
    setEditingSection(sectionKey);
    setSaveError('');
  };

  const cancelEdit = (): void => {
    setEditingSection(null);
    setEditingState({});
    setSaveError('');
    setHighlightPanelRecommendations(false);
  };

  const onChange = (key: keyof EditingState, value: string): void => {
    setEditingState(function(prev) { return { ...prev, [key]: value }; });
  };

  // Status-change/case-closed comm log stamps have no extraFields, so unlike
  // handleSendEmail's admin-only appendCommLogEntry call (which sets Status
  // atomically alongside the log write), these can route through the
  // reviewer-safe elevated path too — reviewers now trigger status changes
  // via setQuickStatus and need this to succeed without Edit rights.
  const saveCommLogEntryForCurrentUser = (itemId: number, entry: ICommLogEntry): Promise<ICommLogEntry[]> => {
    if (!isAdmin && appendCommLogEntryElevated) {
      return appendCommLogEntryElevated(itemId, currentUserEmail, entry);
    }
    return appendCommLogEntry(itemId, entry);
  };

  // Emails everyone involved (submitter, person responsible, line manager,
  // and the shared/admin inbox) and logs one "Case Closed" comm log entry
  // recording all of them, when a case is marked Closed — best-effort: the
  // status change itself has already been saved, so a failure here
  // shouldn't be reported as a failed save. Unlike getInvolvedPeople (used
  // for the request/reply "bystander" notices, which deliberately excludes
  // the person responsible since they already get a direct request/reply
  // email on that interaction), a case-closed notice is the only
  // communication these three ever receive about the outcome — so all
  // three (plus the shared inbox) are genuinely recipients here, not just
  // bystanders getting an FYI copy of something else.
  // log is passed in explicitly rather than read from the commLog state
  // variable — this is called from inside setQuickStatus's promise chain,
  // immediately after setCommLog(updatedLog) queues a state update for the
  // StatusChange stamp that was just appended. React state updates aren't
  // applied synchronously, so the commLog closure this function would
  // otherwise capture is still the pre-update value from whatever render
  // was current when "Mark as Closed" was clicked — one message behind,
  // always missing at least that stamp (and, on a report with a fast
  // back-and-forth right before closing, sometimes more). Bug caused the
  // PDF attached to the closure email to silently drop the most recent
  // comm log entries.
  const sendClosedNotification = (item: FullItem, log: ICommLogEntry[]): void => {
    if (!sendEmailFromMailbox) return;
    const ref = getRef(item.Id);
    const reason = item.Reason_x0020_for_x0020_Decision || '';
    const panelFields = {
      reportedToIco: item.Reported_x0020_to_x0020_ICO || '',
      eLearningRequired: item.Staff_x0020_e_x002d_Learning_x00 || '',
      proceduresReviewRequired: item.Review_x0020_of_x0020_Procedures || '',
    };

    const candidates: Array<{ name: string; email: string }> = [
      { name: item.Author ? item.Author.Title : '', email: item.Author ? item.Author.EMail : '' },
      { name: item.Name_x0020_of_x0020_person_x0020 || '', email: item.Email_x0020_address_x0020_of_x00 || '' },
      { name: item.Line_x0020_manager_x0020_of_x002 || '', email: item.Email_x0020_address_x0020_of_x000 || '' },
      // The shared/admin inbox gets a copy of every closure too, same as it
      // gets the "New Data Breach Report Received" notice on submission —
      // it's a fixed team mailbox, always included regardless of who else
      // is (or isn't) named on the report.
      { name: '', email: ADMIN_NOTIFICATION_EMAIL },
    ];
    // Deduplicate by email — the same person often holds more than one of
    // these roles (e.g. submitter and person responsible are frequently the
    // same person), and they should only get one copy of the notice.
    const recipients: Array<{ name: string; email: string }> = [];
    candidates.forEach(function(c) {
      const key = c.email.toLowerCase().trim();
      if (!key) return;
      if (recipients.filter(function(r) { return r.email.toLowerCase().trim() === key; }).length === 0) {
        recipients.push({ name: c.name, email: c.email });
      }
    });
    if (recipients.length === 0) return;

    const subject = `Data Breach Report ${ref} - Case Closed`;
    const message = reason ? `This case has been closed.\n\nReason for decision: ${reason}` : 'This case has been closed.';

    // Built once and reused for every recipient — same PDF, not regenerated
    // per person. Best-effort: a PDF-generation failure shouldn't block the
    // closure notice itself, so this just falls back to sending without an
    // attachment rather than failing the whole close.
    let attachment: IEmailAttachment | undefined;
    try {
      // Stripped the same way the CSV export strips its values — field
      // content can carry HTML from rich-text SharePoint columns.
      const sections: IReportSection[] = getReportFieldSections(item).map(function(s) {
        return { title: s.title, rows: s.rows.map(function(r) { return [r[0], stripHtmlIfPresent(r[1])] as [string, string]; }) };
      });
      attachment = buildReportPdfBase64(ref, item.Status || 'Open', sections, log);
    } catch (err) {
      console.error('[ReportViewer] Failed to build case-closed PDF summary — sending without attachment:', err);
    }

    Promise.all(recipients.map(function(r) {
      const body = buildClosedEmail({ ref, reason, recipientName: r.name || 'Colleague', ...panelFields });
      return sendEmailFromMailbox!(r.email, subject, body, attachment).catch(function(err: Error) {
        // Swallowed per-recipient so one bad address doesn't stop the
        // others from being notified or block the comm log write below.
        console.error(`[ReportViewer] Failed to send case-closed notification to ${r.email}:`, err);
      });
    })).then(function() {
      const newEntry: ICommLogEntry = {
        direction: 'CaseClosed',
        sentBy: currentUserEmail,
        sentByName: currentUserName,
        sentTo: recipients.map(function(r) { return r.email; }).join(', '),
        subject,
        message,
        timestamp: new Date().toISOString(),
      };
      saveCommLogEntryForCurrentUser(item.Id, newEntry)
        .then(function(updated) { setCommLog(updated); })
        .catch(function(err: Error) {
          console.error('[ReportViewer] Failed to save communication log entry:', err);
          setCommLog(function(prev) { return prev.concat(newEntry); });
          setCommLogWarning('The case-closed email was sent, but we could not confirm it was saved to this report\'s communication log. Refresh the page to check.');
        });
    }).catch(function(err: Error) {
      // Each per-recipient send already swallows its own error above, so
      // this only catches something going wrong in the .then() itself.
      console.error('[ReportViewer] sendClosedNotification failed unexpectedly:', err);
    });
  };

  // Broadcasts to everyone named on the report — submitter, person
  // responsible, line manager — plus the shared admin inbox, same recipient
  // set as sendClosedNotification and matching buildReopenedEmail's own doc
  // comment ("same fixed broadcast to everyone... since reopening restores
  // everyone's access to the report, not just the person responsible").
  // message is optional free text typed into the reopen-confirm modal (e.g.
  // what additional information is needed). Best-effort, same as
  // sendClosedNotification — the status change itself has already saved by
  // the time this runs, so a failure here shouldn't be reported as a failed
  // status change.
  const sendReopenedNotification = (item: FullItem, message: string): void => {
    if (!sendEmailFromMailbox) return;
    const ref = getRef(item.Id);
    const siteUrl = spContext.pageContext.web.absoluteUrl;

    const candidates: Array<{ name: string; email: string }> = [
      { name: item.Author ? item.Author.Title : '', email: item.Author ? item.Author.EMail : '' },
      { name: item.Name_x0020_of_x0020_person_x0020 || '', email: item.Email_x0020_address_x0020_of_x00 || '' },
      { name: item.Line_x0020_manager_x0020_of_x002 || '', email: item.Email_x0020_address_x0020_of_x000 || '' },
      { name: '', email: ADMIN_NOTIFICATION_EMAIL },
    ];
    // Deduplicate by email — the same person often holds more than one of
    // these roles, and they should only get one copy of the notice.
    const recipients: Array<{ name: string; email: string }> = [];
    candidates.forEach(function(c) {
      const key = c.email.toLowerCase().trim();
      if (!key) return;
      if (recipients.filter(function(r) { return r.email.toLowerCase().trim() === key; }).length === 0) {
        recipients.push({ name: c.name, email: c.email });
      }
    });
    if (recipients.length === 0) return;

    const subject = `Data Breach Report ${ref} - Case Reopened`;
    const trimmedMessage = message.trim();

    Promise.all(recipients.map(function(r) {
      const body = buildReopenedEmail({ ref, itemId: item.Id, siteUrl, recipientName: r.name || 'Colleague', message: trimmedMessage });
      return sendEmailFromMailbox!(r.email, subject, body).catch(function(err: Error) {
        // Swallowed per-recipient so one bad address doesn't stop the
        // others from being notified or block the comm log write below.
        console.error(`[ReportViewer] Failed to send case-reopened notification to ${r.email}:`, err);
      });
    })).then(function() {
      const newEntry: ICommLogEntry = {
        direction: 'CaseReopened',
        sentBy: currentUserEmail,
        sentByName: currentUserName,
        sentTo: recipients.map(function(r) { return r.email; }).join(', '),
        subject,
        message: trimmedMessage || 'This case has been reopened for further review.',
        timestamp: new Date().toISOString(),
      };
      saveCommLogEntryForCurrentUser(item.Id, newEntry)
        .then(function(updated) { setCommLog(updated); })
        .catch(function(err: Error) {
          console.error('[ReportViewer] Failed to save communication log entry:', err);
          setCommLog(function(prev) { return prev.concat(newEntry); });
          setCommLogWarning('The case-reopened email was sent, but we could not confirm it was saved to this report\'s communication log. Refresh the page to check.');
        });
    }).catch(function(err: Error) {
      console.error('[ReportViewer] sendReopenedNotification failed unexpectedly:', err);
    });
  };

  // Status is only ever changed via setQuickStatus (the quick-action buttons)
  // now — there's no Status control in any editable section — so this never
  // needs to check for or trigger a close notification itself.
  const saveSection = (): void => {
    if (!selectedItem) return;
    setSaving(true);
    setSaveError('');
    // Reviewers (not admins) editing Panel recommendations only hold Read at
    // the item level, so this specific save routes through the elevated
    // endpoint instead of the direct item.update() every other section uses.
    const useElevatedPanelSave = editingSection === 'review' && !isAdmin && isReviewer && updatePanelRecommendationsElevated;
    const savePromise = useElevatedPanelSave
      ? updatePanelRecommendationsElevated(selectedItem.Id, currentUserEmail, editingState as IPanelRecommendationsFields)
      : updateBreachRecord(selectedItem.Id, editingState as Partial<IBreachListItem>);
    savePromise
      .then(function() {
        const updated = { ...selectedItem, ...editingState } as FullItem;
        setSelectedItem(updated);
        setItems(function(prev) {
          return prev.map(function(r) {
            if (r.Id !== selectedItem.Id) return r;
            return { ...r, Title: updated.Title || r.Title };
          });
        });
        setEditingSection(null);
        setEditingState({});
        setSaving(false);
        setHighlightPanelRecommendations(false);

        // Re-sync item-level permissions in case this edit changed who's
        // named on the report (e.g. reassigning the person responsible) —
        // fire-and-forget so a permissions hiccup doesn't hold up the save
        // itself, which has already succeeded by this point.
        syncItemAccess(updated.Id, [
          updated.Author ? updated.Author.EMail : undefined,
          updated.Email_x0020_of_x0020_the_x0020_r,
          updated.Email_x0020_address_x0020_of_x00,
          updated.Email_x0020_address_x0020_of_x000,
        ], updated.Status === 'Closed').catch(function(err: Error) {
          console.error('[ReportViewer] Failed to sync item permissions after save:', err);
        });
      })
      .catch(function() {
        setSaveError('Failed to save. Please try again.');
        setSaving(false);
      });
  };

  // Panel recommendations must be explicitly answered before a case can be
  // closed. Checking the Yes/No fields alone isn't enough to prove someone
  // actually reviewed them: "Reported to ICO" is set by the submitter at
  // submission time (so it's never blank), and the SharePoint columns behind
  // e-Learning/Procedures may carry their own default value regardless of
  // whether an admin ever opened this section. Reason for decision is the
  // one field nothing else can populate — it's only ever written when an
  // admin explicitly opens Panel recommendations and saves — so requiring it
  // too is what actually proves a deliberate decision was made, not just
  // that the fields happen to hold some value.
  const getMissingPanelRecommendationFields = (): string[] => {
    if (!selectedItem) return [];
    const missing: string[] = [];
    if (!selectedItem.Reported_x0020_to_x0020_ICO) missing.push('Reported to ICO');
    if (!selectedItem.Staff_x0020_e_x002d_Learning_x00) missing.push('Staff e-Learning training');
    if (!selectedItem.Review_x0020_of_x0020_Procedures) missing.push('Procedures review required');
    if (!selectedItem.Reason_x0020_for_x0020_Decision) missing.push('Reason for decision');
    return missing;
  };

  const setQuickStatus = (newStatus: string, reopenMessage?: string): void => {
    if (!selectedItem) return;
    if (newStatus === 'Closed') {
      const missing = getMissingPanelRecommendationFields();
      if (missing.length > 0) {
        // Open the section for editing and bring it into view, rather than
        // leaving the admin/reviewer to hunt for it after reading the error
        // text. startEdit clears saveError internally, so it must run first
        // — otherwise it would wipe out the message set right after it.
        startEdit('review');
        setSaveError(`Before closing this case, complete the following in Panel recommendations: ${missing.join(', ')}.`);
        setHighlightPanelRecommendations(true);
        if (panelRecommendationsRef.current) {
          panelRecommendationsRef.current.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
        return;
      }
    }
    const previousStatus = selectedItem.Status;
    // Reopening clears the old closing rationale — it no longer applies once
    // the case is back under review, and clearing it forces a genuinely
    // fresh decision (not a stale reason left over from before) if it's
    // closed again, consistent with the mandatory-decision check above.
    const isReopening = newStatus === 'Open' && previousStatus === 'Closed';
    const statusUpdate: Partial<IBreachListItem> = { Status: newStatus } as Partial<IBreachListItem>;
    if (isReopening) {
      statusUpdate.Reason_x0020_for_x0020_Decision = '';
    }
    setSaving(true);
    setSaveError('');
    setCommLogWarning('');
    // Reviewers only hold Read at the item level — route the Status write
    // through the same elevated endpoint Panel recommendations uses rather
    // than the direct item.update() admins use.
    const statusSavePromise = (!isAdmin && isReviewer && updatePanelRecommendationsElevated)
      ? updatePanelRecommendationsElevated(selectedItem.Id, currentUserEmail, statusUpdate as IPanelRecommendationsFields)
      : updateBreachRecord(selectedItem.Id, statusUpdate);
    statusSavePromise
      .then(function() {
        const updated = { ...selectedItem, ...statusUpdate } as FullItem;
        setSelectedItem(updated);
        setItems(function(prev) {
          return prev.map(function(r) { return r.Id === selectedItem.Id ? { ...r, Status: newStatus } : r; });
        });
        setSaving(false);

        // Re-sync item-level permissions immediately on a status change —
        // closing drops the named individuals (submitter/officer/
        // responsible/line manager) down to no access so a closed report
        // doesn't sit forever with unique permissions naming people who no
        // longer need it; reopening restores them. Best-effort/fire-and-
        // forget like the other syncItemAccess call sites — a Reviewer-
        // triggered close 403s here (Reviewers don't hold Manage
        // Permissions) and relies on the external permission-sync Logic
        // App's own poll to catch up instead.
        syncItemAccess(selectedItem.Id, [
          selectedItem.Author ? selectedItem.Author.EMail : undefined,
          selectedItem.Email_x0020_of_x0020_the_x0020_r,
          selectedItem.Email_x0020_address_x0020_of_x00,
          selectedItem.Email_x0020_address_x0020_of_x000,
        ], newStatus === 'Closed').catch(function(err: Error) {
          console.error('[ReportViewer] Failed to sync item permissions after status change:', err);
        });

        // Stamp who made the change into the comm log — appendCommLogEntry
        // does its own read-modify-write, so this is deliberately awaited
        // before sendClosedNotification's own log write (the "case closed"
        // email notice) to avoid the two racing and one clobbering the other.
        const statusEntry: ICommLogEntry = {
          direction: 'StatusChange',
          sentBy: currentUserEmail,
          sentByName: currentUserName,
          sentTo: '',
          // Carries the resulting status so the header can look up "who
          // closed this and when" without having to parse it back out of
          // the free-text message below.
          subject: newStatus,
          message: `Status changed from ${previousStatus || 'Open'} to ${newStatus}.`,
          timestamp: new Date().toISOString(),
        };
        saveCommLogEntryForCurrentUser(selectedItem.Id, statusEntry)
          .then(function(updatedLog) { setCommLog(updatedLog); return updatedLog; })
          .catch(function(err: Error) {
            console.error('[ReportViewer] Failed to save communication log entry:', err);
            const fallbackLog = commLog.concat(statusEntry);
            setCommLog(fallbackLog);
            setCommLogWarning('The status change was saved, but we could not confirm it was recorded in this report\'s communication log. Refresh the page to check.');
            return fallbackLog;
          })
          .then(function(latestCommLog) {
            if (newStatus === 'Closed' && previousStatus !== 'Closed') {
              sendClosedNotification(updated, latestCommLog);
            }
            if (isReopening) {
              sendReopenedNotification(updated, reopenMessage || '');
            }
          });
      })
      .catch(function() {
        setSaveError('Failed to update status. Please try again.');
        setSaving(false);
      });
  };

  const f = (key: keyof FullItem): string => {
    if (!selectedItem) return '';
    const val = selectedItem[key];
    return val ? stripHtmlIfPresent(String(val)) : '';
  };

  // ── Detail load error ────────────────────────────────────────────────────────
  // Takes precedence over the list view below, which would otherwise match
  // the same (!selectedItem && !loadingDetail) state after a failed fetch —
  // silently showing what looks like an empty report list instead of
  // explaining that access was denied (or the report doesn't exist).

  if (detailError) {
    return (
      <div style={{ padding: '20px' }}>
        <a href="#" className="govuk-back-link" onClick={function(e) { e.preventDefault(); setDetailError(''); setSelectedItem(null); }}>
          Back to reports
        </a>
        <div className="govuk-error-summary" style={{ marginTop: '20px' }}>
          <div role="alert">
            <h2 className="govuk-error-summary__title">Unable to open this report</h2>
            <div className="govuk-error-summary__body"><p className="govuk-body">{detailError}</p></div>
          </div>
        </div>
      </div>
    );
  }

  // ── List view ──────────────────────────────────────────────────────────────

  if (!selectedItem && !loadingDetail) {
    return <ReportList items={items} loading={loading} isAdmin={isAdmin} controls={listControls} onOpen={openItem} />;
  }

  // ── Loading detail ─────────────────────────────────────────────────────────

  if (loadingDetail) {
    return <div style={{ padding: '20px' }}><p className="govuk-body">Loading report…</p></div>;
  }

  // ── Detail / edit view ─────────────────────────────────────────────────────

  // Most recent "closed" stamp from the comm log, for display next to the
  // status tag at the top. commLog is appended to in order, so the last
  // matching entry is the most recent one — correct as long as the report
  // is currently Closed (if it's been reopened since, this is gated out by
  // only rendering when Status is Closed).
  const closedStamp = commLog
    .filter(function(e) { return e.direction === 'StatusChange' && e.subject === 'Closed'; })
    .slice(-1)[0] || null;

  // Shared between the admin comm-log panel and the recipient comm-log panel
  // so a user who is both admin and the recipient (submitter/responsible
  // person) for this specific report gets the reply flow embedded in the
  // admin panel, instead of it disappearing because isAdmin hid the other one.
  const userReplyBanner = userReplySent && (
    <div className="govuk-notification-banner govuk-notification-banner--success no-print" role="alert" style={{ marginBottom: '16px' }}>
      <div className="govuk-notification-banner__header">
        <h2 className="govuk-notification-banner__title">Information submitted</h2>
      </div>
      <div className="govuk-notification-banner__content">
        <p className="govuk-body">Your information has been sent and logged against this report.</p>
      </div>
    </div>
  );

  const userReplyForm = showUserReply && (
    <div className="govuk-inset-text no-print" style={{ marginBottom: '16px' }}>
      <h2 className="govuk-heading-s">Submit more information</h2>
      <p className="govuk-body-s" style={{ color: '#505a5f' }}>
        Your message will be sent to the data breach notification team and logged against this report.
      </p>
      {userReplyError && (
        <p className="govuk-error-message" style={{ marginBottom: '12px' }}>{userReplyError}</p>
      )}
      <div className="govuk-form-group" style={{ marginBottom: '16px' }}>
        <label className="govuk-label" htmlFor="user-reply-text">Your message</label>
        <textarea
          id="user-reply-text"
          className="govuk-textarea"
          rows={4}
          value={userReplyText}
          onChange={function(e) { setUserReplyText(e.target.value); setUserReplyError(''); }}
          placeholder="Type your reply or additional information here…"
          disabled={userReplySending}
        />
      </div>
      <div style={{ display: 'flex', gap: '10px' }}>
        <button
          type="button"
          className={`govuk-button${(!userReplyText.trim() || userReplySending) ? ' govuk-button--disabled' : ''}`}
          disabled={!userReplyText.trim() || userReplySending}
          onClick={handleUserReply}
        >
          {userReplySending ? 'Sending…' : 'Submit'}
        </button>
        <button
          type="button"
          className="govuk-button govuk-button--secondary"
          disabled={userReplySending}
          onClick={function() { setShowUserReply(false); setUserReplyText(''); setUserReplyError(''); }}
        >
          Cancel
        </button>
      </div>
    </div>
  );

  return (
    <div id="breach-report-detail" style={{ padding: '20px' }}>
      <a className="govuk-back-link no-print" href="#" onClick={function(e) { e.preventDefault(); setSelectedItem(null); setEditingSection(null); setShowRequestInfo(false); setRequestInfoText(''); }}>
        Back to reports
      </a>

      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginTop: '20px' }}>
        <div>
          <span className="govuk-caption-l">{selectedItem ? getRef(selectedItem.Id) : ''}</span>
          <h1 className="govuk-heading-l" style={{ marginBottom: '5px' }}>Data Breach Report</h1>
        </div>
        <div style={{ textAlign: 'right' }}>
          <strong className={`govuk-tag${f('Status') === 'Closed' ? ' govuk-tag--green' : f('Status') === 'Pending' ? ' govuk-tag--yellow' : ''}`}>
            {f('Status') || 'Open'}
          </strong>
          <p className="govuk-body-s" style={{ marginTop: '5px', color: '#505a5f' }}>
            Submitted by {selectedItem && selectedItem.Author ? selectedItem.Author.Title : ''} on {formatDate(f('Created'))}
          </p>
          {f('Status') === 'Closed' && closedStamp && (
            <p className="govuk-body-s" style={{ marginTop: '5px', color: '#505a5f' }}>
              Closed by {closedStamp.sentBy || 'unknown'} on {formatDate(closedStamp.timestamp)}
            </p>
          )}
        </div>
      </div>

      <div className="no-print" style={{ display: 'flex', gap: '10px', marginTop: '16px', marginBottom: '24px', flexWrap: 'wrap' }}>
        <button type="button" className="govuk-button govuk-button--secondary" style={{ marginBottom: 0 }} onClick={printReport}>
          Print / Save as PDF
        </button>
        <button type="button" className="govuk-button govuk-button--secondary" style={{ marginBottom: 0 }} onClick={function() { if (selectedItem) { downloadSingleCsv(selectedItem); } }}>
          Download CSV
        </button>
      </div>

      {saveError && (
        <div className="govuk-error-summary">
          <div role="alert">
            <h2 className="govuk-error-summary__title">There is a problem</h2>
            <div className="govuk-error-summary__body"><p className="govuk-body">{saveError}</p></div>
          </div>
        </div>
      )}

      {commLogWarning && (
        <div className="govuk-warning-text no-print" style={{ marginBottom: '20px' }}>
          <span className="govuk-warning-text__icon" aria-hidden="true">!</span>
          <strong className="govuk-warning-text__text">
            <span className="govuk-visually-hidden">Warning</span>
            {commLogWarning}
          </strong>
        </div>
      )}

      {/* Review status banner — admins and reviewers */}
      {(isAdmin || isReviewer) && (
        <div className="no-print" style={{ background: '#f3f2f1', borderRadius: '4px', padding: '16px 20px', marginBottom: '24px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
          <div style={{ display: 'flex', gap: '24px', alignItems: 'center', flexWrap: 'wrap' }}>
            <div>
              <p className="govuk-body-s" style={{ color: '#505a5f', marginBottom: '4px' }}>Report ID</p>
              <strong className="govuk-body">{selectedItem ? getRef(selectedItem.Id) : ''}</strong>
            </div>
            <div>
              <p className="govuk-body-s" style={{ color: '#505a5f', marginBottom: '4px' }}>Review status</p>
              <strong className={`govuk-tag${f('Status') === 'Closed' ? ' govuk-tag--green' : f('Status') === 'Pending' ? ' govuk-tag--yellow' : ''}`}>
                {f('Status') || 'Open'}
              </strong>
            </div>
          </div>
          <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
            {f('Status') !== 'Pending' && f('Status') !== 'Closed' && (
              <button type="button" className="govuk-button govuk-button--secondary" style={{ marginBottom: 0 }} disabled={saving} onClick={function() { setQuickStatus('Pending'); }}>Mark as Pending</button>
            )}
            {f('Status') !== 'Closed' && (
              <button type="button" className="govuk-button" style={{ marginBottom: 0 }} disabled={saving} onClick={function() { setShowCloseConfirm(true); }}>Mark as Closed</button>
            )}
            {f('Status') === 'Closed' && (
              <button type="button" className="govuk-button govuk-button--secondary" style={{ marginBottom: 0 }} disabled={saving} onClick={function() { setReopenMessage(''); setShowReopenConfirm(true); }}>Reopen</button>
            )}
          </div>
        </div>
      )}

      {showCloseConfirm && selectedItem && (
        <CloseCaseDialog
          item={selectedItem}
          f={f}
          saving={saving}
          onConfirm={function() { setShowCloseConfirm(false); setQuickStatus('Closed'); }}
          onCancel={function() { setShowCloseConfirm(false); }}
        />
      )}

      {showReopenConfirm && selectedItem && (
        <ReopenCaseDialog
          message={reopenMessage}
          onMessageChange={setReopenMessage}
          saving={saving}
          onConfirm={function() { setShowReopenConfirm(false); setQuickStatus('Open', reopenMessage); }}
          onCancel={function() { setShowReopenConfirm(false); }}
        />
      )}

      {/* Panel recommendations — admins and reviewers. Status is deliberately
          not editable here; it's set exclusively via the quick-action
          buttons above to keep there being a single place that changes it. */}
      {(isAdmin || isReviewer) && (
        <div
          ref={panelRecommendationsRef}
          style={highlightPanelRecommendations ? {
            border: '3px solid #d4351c',
            borderRadius: '6px',
            padding: '10px',
            marginBottom: '20px',
            transition: 'border-color 0.2s ease-in-out',
          } : undefined}
        >
          {highlightPanelRecommendations && (
            <p className="govuk-body-s no-print" style={{ color: '#d4351c', fontWeight: 'bold', marginBottom: '10px' }}>
              Complete the fields below before this case can be marked as Closed.
            </p>
          )}
          <Section title="Panel recommendations" sectionKey="review" editingSection={editingSection} saving={saving} canEdit={isAdmin || isReviewer} onEdit={startEdit} onSave={saveSection} onCancel={cancelEdit} adminOnly restrictedLabel="Reviewer">
            <Field label="Reported to ICO" value={f('Reported_x0020_to_x0020_ICO')} fieldKey="Reported_x0020_to_x0020_ICO" editing={editingSection === 'review'} options={YES_NO} editingState={editingState} onChange={onChange} />
            <Field label="Staff e-Learning training" value={f('Staff_x0020_e_x002d_Learning_x00')} fieldKey="Staff_x0020_e_x002d_Learning_x00" editing={editingSection === 'review'} options={YES_NO} editingState={editingState} onChange={onChange} />
            <Field label="Procedures review required" value={f('Review_x0020_of_x0020_Procedures')} fieldKey="Review_x0020_of_x0020_Procedures" editing={editingSection === 'review'} options={YES_NO} editingState={editingState} onChange={onChange} />
            <Field label="Reason for decision" value={f('Reason_x0020_for_x0020_Decision')} fieldKey="Reason_x0020_for_x0020_Decision" editing={editingSection === 'review'} multiline editingState={editingState} onChange={onChange} />
          </Section>
        </div>
      )}

      {/* Communication Log — admins and reviewers; reviewers need this to
          request more information as part of the review process. */}
      {(isAdmin || isReviewer) && (
        <div style={{ marginBottom: '30px', background: '#f3f2f1', padding: '16px', borderRadius: '4px' }}>
          <div className="no-print" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '2px solid #1d70b8', paddingBottom: '8px', marginBottom: '16px' }}>
            <h2 className="govuk-heading-s" style={{ margin: 0 }}>
              Communication Log
              <strong className="govuk-tag" style={{ marginLeft: '10px', verticalAlign: 'middle' }}>Reviewer</strong>
            </h2>
            <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
              <button
                type="button"
                className="govuk-button govuk-button--secondary"
                style={{ marginBottom: 0 }}
                disabled={refreshingLog}
                onClick={refreshCommLog}
                title="Refresh communication log"
              >
                {refreshingLog ? '…' : '↻'}
              </button>
              {/* Hidden for the person responsible themselves — most often an
                  admin/reviewer who is also named as responsible on their
                  own report. Requesting more information sends it to
                  yourself, which never makes sense — see isCurrentUserRecipient. */}
              {!isCurrentUserRecipient() && (
                <button
                  type="button"
                  className="govuk-button"
                  style={{ marginBottom: 0 }}
                  onClick={function() {
                    // These two panels are mutually exclusive — opening one
                    // must close the other, otherwise both toggle buttons can
                    // read "Cancel" at once (only their own panel's flag is
                    // checked for the label), which is confusing to look at.
                    if (!showRequestInfo) { setShowUserReply(false); setUserReplyText(''); setUserReplyError(''); }
                    setShowRequestInfo(function(s) { return !s; });
                    setRequestInfoText(''); setSendError(''); setSendSuccess(false);
                  }}
                >
                  {showRequestInfo ? 'Cancel' : 'Request more information'}
                </button>
              )}
              {isCurrentUserRecipient() && (
                <button
                  type="button"
                  className="govuk-button govuk-button--secondary"
                  style={{ marginBottom: 0 }}
                  onClick={function() {
                    if (!showUserReply) { setShowRequestInfo(false); setRequestInfoText(''); setSendError(''); setSendSuccess(false); }
                    setShowUserReply(function(s) { return !s; });
                    setUserReplyError(''); setUserReplySent(false);
                  }}
                >
                  {showUserReply ? 'Cancel' : 'Submit more information'}
                </button>
              )}
            </div>
          </div>

          {sendSuccess && (
            <div className="govuk-notification-banner govuk-notification-banner--success no-print" role="alert" style={{ marginBottom: '16px' }}>
              <div className="govuk-notification-banner__header">
                <h2 className="govuk-notification-banner__title">Request logged</h2>
              </div>
              <div className="govuk-notification-banner__content">
                <p className="govuk-body">
                  Your request for additional information has been logged
                  {sendEmailFromMailbox && sharedMailbox ? <> and emailed to <strong>{getRequestInfoRecipient().name}</strong></> : ''}.
                </p>
              </div>
            </div>
          )}

          {isCurrentUserRecipient() && userReplyBanner}
          {isCurrentUserRecipient() && userReplyForm}

          {showRequestInfo && (
            <div className="govuk-inset-text no-print" style={{ marginBottom: '16px' }}>
              <h2 className="govuk-heading-s">Request additional information</h2>
              <p className="govuk-body-s" style={{ color: '#505a5f' }}>
                To: <strong>{getRequestInfoRecipient().name}</strong>
                {getRequestInfoRecipient().email ? ` (${getRequestInfoRecipient().email})` : ''}
                {sendEmailFromMailbox && sharedMailbox ? <> &mdash; will be sent from <strong>{sharedMailbox}</strong></> : ''}
              </p>
              {sendError && (
                <p className="govuk-error-message" style={{ marginBottom: '12px' }}>{sendError}</p>
              )}
              <div className="govuk-form-group" style={{ marginBottom: '16px' }}>
                <label className="govuk-label" htmlFor="request-info-text">What information do you need?</label>
                <textarea
                  id="request-info-text"
                  className="govuk-textarea"
                  rows={4}
                  value={requestInfoText}
                  onChange={function(e) { setRequestInfoText(e.target.value); setSendError(''); }}
                  placeholder="Describe the additional information or clarification required…"
                  disabled={sending}
                />
              </div>
              <details className="govuk-details" style={{ marginBottom: '16px' }}>
                <summary className="govuk-details__summary">
                  <span className="govuk-details__summary-text">Preview full email</span>
                </summary>
                <div className="govuk-details__text" style={{ background: '#ffffff', border: '1px solid #b1b4b6', padding: '12px 16px', borderRadius: '2px', whiteSpace: 'pre-wrap', fontFamily: 'monospace', fontSize: '13px', color: '#0b0c0c' }}>
                  {buildEmailPreview()}
                </div>
              </details>
              <div style={{ display: 'flex', gap: '10px' }}>
                <button
                  type="button"
                  className={`govuk-button${(!requestInfoText.trim() || sending) ? ' govuk-button--disabled' : ''}`}
                  disabled={!requestInfoText.trim() || sending}
                  onClick={handleSendEmail}
                >
                  {sending ? 'Sending…' : 'Send'}
                </button>
                <button
                  type="button"
                  className="govuk-button govuk-button--secondary"
                  disabled={sending}
                  onClick={function() { setShowRequestInfo(false); setRequestInfoText(''); setSendError(''); }}
                >
                  Cancel
                </button>
              </div>
            </div>
          )}

          <CommLogEntries entries={commLog} audience="staff" />
        </div>
      )}

      {/* Communication Log — visible to anyone named in the report (submitter,
          person responsible, or line manager) when they are not an admin or
          reviewer. Only the submitter/person responsible get the reply
          action — the line manager can view the history but isn't asked to
          provide further breach details themselves. Admins/reviewers who are
          also a recipient (e.g. a reviewer who is also the person
          responsible on this report) get the reply flow embedded in the
          admin/reviewer panel above instead of duplicated here — excluding
          only isAdmin here left reviewers falling through to both sections
          at once, rendering two "Communication Log" panels (each with its
          own Request/Submit more information toggle, both wired to the same
          showRequestInfo/showUserReply state) stacked on the page. */}
      {!isAdmin && !isReviewer && isCurrentUserInvolved() && (
        <div style={{ marginBottom: '30px' }}>
          <div className="no-print" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '2px solid #1d70b8', paddingBottom: '8px', marginBottom: '16px' }}>
            <h2 className="govuk-heading-s" style={{ margin: 0 }}>Communication Log</h2>
            <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
              <button
                type="button"
                className="govuk-button govuk-button--secondary"
                style={{ marginBottom: 0 }}
                disabled={refreshingLog}
                onClick={refreshCommLog}
                title="Refresh communication log"
              >
                {refreshingLog ? '…' : '↻'}
              </button>
              {isCurrentUserRecipient() && !showUserReply && (
                <button
                  type="button"
                  className="govuk-button"
                  style={{ marginBottom: 0 }}
                  onClick={function() { setShowUserReply(true); setUserReplyError(''); setUserReplySent(false); }}
                >
                  Submit more information
                </button>
              )}
            </div>
          </div>

          {isCurrentUserRecipient() && userReplyBanner}
          {isCurrentUserRecipient() && userReplyForm}

          <CommLogEntries entries={commLog} audience="involved" />
        </div>
      )}

      <SubmittedReport f={f} attachments={attachments} editingSection={editingSection} editingState={editingState} saving={saving} onChange={onChange} onEdit={startEdit} onSave={saveSection} onCancel={cancelEdit} />

    </div>
  );
};

export default ReportViewer;
