import * as React from 'react';
import { ICommLogEntry } from '../../dataBreachForm/services/SharePointService';
import { staffAttribution } from '../services/ReportPdf';
import { formatDateTime, getStatusChangeStyle } from './reportViewerUtils';

// 'staff' (admins/reviewers) sees who actioned each entry and who it went to.
// 'involved' (submitter/person responsible/line manager) sees just the
// messages — not which team member sent them or the full recipient list.
type Audience = 'staff' | 'involved';

const CASE_EVENT_STYLE: { [direction: string]: { label: string; tagClass: string; borderColor: string } } = {
  CaseClosed: { label: 'Case Closed', tagClass: 'govuk-tag--red', borderColor: '#d4351c' },
  CaseReopened: { label: 'Case Reopened', tagClass: 'govuk-tag--blue', borderColor: '#1d70b8' },
};

const metaStyle: React.CSSProperties = { color: '#505a5f', marginBottom: '4px' };

const EntryHeader: React.FC<{ tagClass: string; label: string; timestamp: string }> = ({ tagClass, label, timestamp }) => (
  <div style={{ display: 'flex', gap: '8px', alignItems: 'center', marginBottom: '6px', flexWrap: 'wrap' }}>
    <strong className={`govuk-tag${tagClass ? ' ' + tagClass : ''}`}>{label}</strong>
    <span className="govuk-body-s" style={{ color: '#505a5f' }}>{formatDateTime(timestamp)}</span>
  </div>
);

const MessageBox: React.FC<{ message: string }> = ({ message }) => (
  <div style={{ background: '#ffffff', border: '1px solid #b1b4b6', padding: '10px 14px', borderRadius: '2px' }}>
    <p className="govuk-body-s" style={{ whiteSpace: 'pre-wrap', marginBottom: 0 }}>{message}</p>
  </div>
);

const CommLogEntries: React.FC<{ entries: ICommLogEntry[]; audience: Audience }> = ({ entries, audience }) => {
  const isStaff = audience === 'staff';

  if (entries.length === 0) {
    return (
      <p className="govuk-body-s" style={{ color: '#505a5f', marginBottom: 0 }}>
        {isStaff ? 'No communications logged for this report.' : 'No communications have been logged for this report yet.'}
      </p>
    );
  }

  return (
    <div>
      {entries.map(function(entry, i) {
        const key = `${entry.direction}-${entry.timestamp}-${i}`;

        if (entry.direction === 'StatusChange') {
          const style = getStatusChangeStyle(entry.subject);
          return (
            <div key={key} style={{ borderLeft: `4px solid ${style.borderColor}`, paddingLeft: '16px', marginBottom: '20px' }}>
              <EntryHeader tagClass={style.tagClass} label={style.label} timestamp={entry.timestamp} />
              <p className="govuk-body-s" style={{ color: '#505a5f', marginBottom: 0 }}>
                {entry.message}{isStaff && <> <strong>By:</strong> {staffAttribution(entry)}</>}
              </p>
            </div>
          );
        }

        const caseEvent = CASE_EVENT_STYLE[entry.direction];
        if (caseEvent) {
          return (
            <div key={key} style={{ borderLeft: `4px solid ${caseEvent.borderColor}`, paddingLeft: '16px', marginBottom: '20px' }}>
              <EntryHeader tagClass={caseEvent.tagClass} label={caseEvent.label} timestamp={entry.timestamp} />
              {isStaff && (
                <p className="govuk-body-s" style={metaStyle}>
                  <strong>From:</strong> {staffAttribution(entry)} &rarr; <strong>To:</strong> {entry.sentTo}
                </p>
              )}
              <MessageBox message={entry.message} />
            </div>
          );
        }

        const isOutbound = entry.direction === 'Outbound';
        const replyLabel = isStaff ? 'Reply received' : 'Information submitted';
        return (
          <div key={key} style={{ borderLeft: `4px solid ${isOutbound ? '#1d70b8' : '#00703c'}`, paddingLeft: '16px', marginBottom: '20px' }}>
            <EntryHeader tagClass={isOutbound ? '' : 'govuk-tag--green'} label={isOutbound ? 'Info requested' : replyLabel} timestamp={entry.timestamp} />
            {isStaff && (
              <p className="govuk-body-s" style={metaStyle}>
                {isOutbound
                  ? <><strong>From:</strong> {staffAttribution(entry)} &rarr; <strong>To:</strong> {entry.sentTo}</>
                  : <><strong>From:</strong> {entry.sentBy}</>
                }
              </p>
            )}
            {isStaff && entry.subject && (
              <p className="govuk-body-s" style={{ color: '#505a5f', fontStyle: 'italic', marginBottom: '8px' }}>{entry.subject}</p>
            )}
            <MessageBox message={entry.message} />
          </div>
        );
      })}
    </div>
  );
};

export default CommLogEntries;
