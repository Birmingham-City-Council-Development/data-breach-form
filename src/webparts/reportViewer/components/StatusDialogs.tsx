import * as React from 'react';
import { ADMIN_NOTIFICATION_EMAIL } from '../../dataBreachForm/config';
import { FullItem, FieldReader } from './reportViewerTypes';
import { getRef } from './reportViewerUtils';

const overlayStyle: React.CSSProperties = { position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000, padding: '20px' };
const panelStyle: React.CSSProperties = { background: '#fff', borderRadius: '4px', padding: '24px', maxWidth: '600px', width: '100%', maxHeight: '90vh', overflowY: 'auto', boxShadow: '0 2px 12px rgba(0,0,0,0.3)' };

// Lists who will actually receive the case-closed email — the same recipient
// set (and dedupe) sendClosedNotification uses.
function closeRecipients(item: FullItem): Array<{ label: string; name: string; email: string }> {
  const candidates: Array<{ label: string; name: string; email: string }> = [
    { label: 'Submitter', name: item.Author ? item.Author.Title : '', email: item.Author ? item.Author.EMail : '' },
    { label: 'Person responsible', name: item.Name_x0020_of_x0020_person_x0020 || '', email: item.Email_x0020_address_x0020_of_x00 || '' },
    { label: 'Line manager', name: item.Line_x0020_manager_x0020_of_x002 || '', email: item.Email_x0020_address_x0020_of_x000 || '' },
    { label: 'Shared inbox', name: '', email: ADMIN_NOTIFICATION_EMAIL },
  ];
  const seen: string[] = [];
  return candidates.filter(function(c) {
    const key = c.email.toLowerCase().trim();
    if (!key || seen.indexOf(key) !== -1) return false;
    seen.push(key);
    return true;
  });
}

export const CloseCaseDialog: React.FC<{
  item: FullItem;
  f: FieldReader;
  saving: boolean;
  onConfirm: () => void;
  onCancel: () => void;
}> = ({ item, f, saving, onConfirm, onCancel }) => {
  const rows = closeRecipients(item);
  return (
    <div className="no-print" role="dialog" aria-modal="true" aria-labelledby="close-confirm-title" style={overlayStyle}>
      <div style={panelStyle}>
        <h2 id="close-confirm-title" className="govuk-heading-m" style={{ marginTop: 0 }}>Close this case?</h2>
        <div className="govuk-warning-text" style={{ marginBottom: '16px' }}>
          <span className="govuk-warning-text__icon" aria-hidden="true">!</span>
          <strong className="govuk-warning-text__text">
            <span className="govuk-visually-hidden">Warning</span>
            This is final. Make sure Panel recommendations are correct before continuing.
          </strong>
        </div>
        <p className="govuk-body">Closing this case will:</p>
        <ul className="govuk-list govuk-list--bullet">
          <li>set the Status to <strong>Closed</strong></li>
          <li>
            automatically email everyone involved on this report — the submitter, the person responsible, the line manager, and the shared data breach inbox (whoever of those has an email on file; duplicates are only sent once if the same address applies to more than one of these):
            <ul className="govuk-list govuk-list--bullet" style={{ marginTop: '6px' }}>
              {rows.length === 0
                ? <li>no email addresses found on this report — no notification will be sent</li>
                : rows.map(function(r) {
                  return <li key={r.email}>{r.label}{r.name ? ` — ${r.name}` : ''} ({r.email})</li>;
                })}
            </ul>
          </li>
        </ul>
        <p className="govuk-body">That email will include:</p>
        <ul className="govuk-list govuk-list--bullet">
          <li>the report reference ({getRef(item.Id)})</li>
          <li>the Panel recommendations: Reported to ICO ({f('Reported_x0020_to_x0020_ICO') || 'not set'}), Staff e-Learning training ({f('Staff_x0020_e_x002d_Learning_x00') || 'not set'}), Procedures review required ({f('Review_x0020_of_x0020_Procedures') || 'not set'})</li>
          <li>the Reason for decision text{f('Reason_x0020_for_x0020_Decision') ? '' : ' (currently blank)'}</li>
        </ul>
        <div style={{ display: 'flex', gap: '10px', marginTop: '20px' }}>
          <button type="button" className="govuk-button govuk-button--warning" style={{ marginBottom: 0 }} disabled={saving} onClick={onConfirm}>
            Yes, close this case
          </button>
          <button type="button" className="govuk-button govuk-button--secondary" style={{ marginBottom: 0 }} onClick={onCancel}>
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
};

export const ReopenCaseDialog: React.FC<{
  message: string;
  onMessageChange: (value: string) => void;
  saving: boolean;
  onConfirm: () => void;
  onCancel: () => void;
}> = ({ message, onMessageChange, saving, onConfirm, onCancel }) => (
  <div className="no-print" role="dialog" aria-modal="true" aria-labelledby="reopen-confirm-title" style={overlayStyle}>
    <div style={panelStyle}>
      <h2 id="reopen-confirm-title" className="govuk-heading-m" style={{ marginTop: 0 }}>Reopen this case?</h2>
      <p className="govuk-body">Reopening this case will:</p>
      <ul className="govuk-list govuk-list--bullet">
        <li>set the Status back to <strong>Open</strong> and clear the previous Reason for decision, since it no longer applies</li>
        <li>restore access to the report for the submitter, person responsible, and line manager</li>
        <li>
          automatically email the submitter, person responsible, and line manager (plus the data breach notification team) — same recipients as a case-closed notice
        </li>
      </ul>
      <div className="govuk-form-group" style={{ marginBottom: '8px' }}>
        <label className="govuk-label" htmlFor="reopen-message">
          What additional information is needed? (optional)
        </label>
        <div className="govuk-hint">
          If the case is being reopened because something specific is required, describe it here — it will be
          included in the email so everyone notified knows what to provide. Leave this blank to send a
          general reopened notice with no specific request.
        </div>
        <textarea
          id="reopen-message"
          className="govuk-textarea"
          rows={4}
          value={message}
          onChange={function(e) { onMessageChange(e.target.value); }}
          placeholder="Describe the additional information or clarification required…"
          disabled={saving}
        />
      </div>
      <div style={{ display: 'flex', gap: '10px', marginTop: '20px' }}>
        <button type="button" className="govuk-button" style={{ marginBottom: 0 }} disabled={saving} onClick={onConfirm}>
          Yes, reopen this case
        </button>
        <button type="button" className="govuk-button govuk-button--secondary" style={{ marginBottom: 0 }} onClick={onCancel}>
          Cancel
        </button>
      </div>
    </div>
  </div>
);
