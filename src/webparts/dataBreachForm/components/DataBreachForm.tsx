import * as React from 'react';
import styles from './DataBreachForm.module.scss';
import type { IDataBreachFormProps } from './IDataBreachFormProps';
import type { IFormState } from './IFormState';
import { createBreachRecord, uploadAttachmentElevated, searchUsers, appendCommLogEntryElevated, updatePanelRecommendationsElevated, cleanCurrentUserName, ICommLogEntry, IPanelRecommendationsFields } from '../services/SharePointService';
import {
  IBreachEmailData,
  buildSubmitterEmail,
  buildResponsibleEmail,
  buildLineManagerEmail,
  buildAdminEmail,
} from '../services/EmailTemplates';
import ReportViewer from '../../reportViewer/components/ReportViewer';
import type { IEmailAttachment } from '../../reportViewer/components/IReportViewerProps';
import { ADMIN_NOTIFICATION_EMAIL, LIVE_SERVICE_URL, MAX_ATTACHMENT_MB, ALLOWED_ATTACHMENT_EXTENSIONS } from '../config';
import {
  DIRECTORATES,
  getDivisions,
  APPROX_INDIVIDUALS,
  NATURE_OF_BREACH,
  CATEGORIES_PERSONAL_DATA,
  CATEGORIES_INDIVIDUALS,
  IMPACT_TO_SUBJECTS,
  LIKELIHOOD_OF_CONSEQUENCES,
  BREACH_CATEGORIES,
  INDIVIDUALS_INFORMED_OPTIONS,
} from './formOptions';


const STEPS = [
  'Reporting Officer',
  'Person Responsible',
  'Breach Details',
  'Third Party',
  'Consequences & Actions',
  'Supporting Information & Attachments',
  'Review',
];

const REVIEW_STEP = STEPS.length - 1;

type FormErrors = Partial<Record<keyof IFormState, string>>;
type AppMode = 'home' | 'submit' | 'view';

const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const PHONE_REGEX = /^[\d\s+\-().]{7,20}$/;

// ── Reusable components ──────────────────────────────────────────────────────

const TextInput: React.FC<{
  id: string; label: string; value: string;
  onChange: (v: string) => void; hint?: string; type?: string; width?: string; error?: string; readOnly?: boolean;
}> = ({ id, label, value, onChange, hint, type = 'text', width = 'govuk-input--width-20', error, readOnly }) => {
  const describedBy = [hint ? `${id}-hint` : '', error ? `${id}-error` : ''].filter(Boolean).join(' ') || undefined;
  return (
    <div className={`govuk-form-group${error ? ' govuk-form-group--error' : ''}`}>
      <label className="govuk-label" htmlFor={id}>{label}</label>
      {hint && <div id={`${id}-hint`} className="govuk-hint">{hint}</div>}
      {readOnly && <div className="govuk-hint">This is pre-filled from your account and cannot be changed.</div>}
      {error && (
        <p id={`${id}-error`} className="govuk-error-message">
          <span className="govuk-visually-hidden">Error:</span> {error}
        </p>
      )}
      <input
        className={`govuk-input ${width}${error ? ' govuk-input--error' : ''}`}
        id={id} name={id} type={type} value={value}
        aria-describedby={describedBy}
        readOnly={readOnly}
        style={readOnly ? { backgroundColor: '#f3f2f1', cursor: 'not-allowed' } : undefined}
        onChange={readOnly ? undefined : e => onChange(e.target.value)}
      />
    </div>
  );
};

const TextArea: React.FC<{
  id: string; label: string; value: string;
  onChange: (v: string) => void; rows?: number; error?: string;
}> = ({ id, label, value, onChange, rows = 5, error }) => (
  <div className={`govuk-form-group${error ? ' govuk-form-group--error' : ''}`}>
    <label className="govuk-label" htmlFor={id}>{label}</label>
    {error && <p className="govuk-error-message"><span className="govuk-visually-hidden">Error:</span> {error}</p>}
    <textarea
      className={`govuk-textarea${error ? ' govuk-textarea--error' : ''}`}
      id={id} name={id} rows={rows} value={value}
      onChange={e => onChange(e.target.value)}
    />
  </div>
);

const SelectInput: React.FC<{
  id: string; label: string; value: string;
  onChange: (v: string) => void; options: string[];
  disabled?: boolean; disabledHint?: string; error?: string;
}> = ({ id, label, value, onChange, options, disabled, disabledHint, error }) => (
  <div className={`govuk-form-group${error ? ' govuk-form-group--error' : ''}`}>
    <label className="govuk-label" htmlFor={id}>{label}</label>
    {disabled && disabledHint && <div className="govuk-hint">{disabledHint}</div>}
    {error && <p id={`${id}-error`} className="govuk-error-message"><span className="govuk-visually-hidden">Error:</span> {error}</p>}
    <select
      className={`govuk-select${error ? ' govuk-select--error' : ''}`}
      id={id} name={id} value={value}
      disabled={disabled}
      aria-describedby={error ? `${id}-error` : undefined}
      onChange={e => onChange(e.target.value)}
    >
      <option value="">— Select —</option>
      {options.map(o => <option key={o} value={o}>{o}</option>)}
    </select>
  </div>
);

const RadioGroup: React.FC<{
  name: string; label: string; value: string;
  onChange: (v: string) => void; options: string[]; error?: string;
}> = ({ name, label, value, onChange, options, error }) => (
  <div id={name} className={`govuk-form-group${error ? ' govuk-form-group--error' : ''}`}>
    <fieldset className="govuk-fieldset">
      <legend className="govuk-fieldset__legend govuk-fieldset__legend--s">{label}</legend>
      {error && <p id={`${name}-error`} className="govuk-error-message"><span className="govuk-visually-hidden">Error:</span> {error}</p>}
      <div className="govuk-radios govuk-radios--inline">
        {options.map(o => (
          <div key={o} className="govuk-radios__item">
            <input
              className="govuk-radios__input"
              id={`${name}-${o}`} name={name} type="radio" value={o}
              checked={value === o} onChange={() => onChange(o)}
            />
            <label className="govuk-label govuk-radios__label" htmlFor={`${name}-${o}`}>{o}</label>
          </div>
        ))}
      </div>
    </fieldset>
  </div>
);

const CheckboxGroup: React.FC<{
  id: string;
  label: string;
  hint?: string;
  values: string[];
  options: string[];
  onToggle: (value: string) => void;
  error?: string;
  otherValue?: string;
  onOtherChange?: (v: string) => void;
  otherError?: string;
}> = ({ id, label, hint, values, options, onToggle, error, otherValue, onOtherChange, otherError }) => (
  <div id={id} className={`govuk-form-group${error ? ' govuk-form-group--error' : ''}`}>
    <fieldset className="govuk-fieldset">
      <legend className="govuk-fieldset__legend govuk-fieldset__legend--s">{label}</legend>
      {hint && <div className="govuk-hint">{hint}</div>}
      {error && <p className="govuk-error-message"><span className="govuk-visually-hidden">Error:</span> {error}</p>}
      <div className="govuk-checkboxes govuk-checkboxes--small">
        {options.map((option, i) => (
          <div key={option} className="govuk-checkboxes__item">
            <input
              className="govuk-checkboxes__input"
              id={`${id}-${i}`}
              name={`${id}[]`}
              type="checkbox"
              value={option}
              checked={values.indexOf(option) !== -1}
              onChange={() => onToggle(option)}
            />
            <label className="govuk-label govuk-checkboxes__label" htmlFor={`${id}-${i}`}>
              {option}
            </label>
          </div>
        ))}
      </div>
      {values.indexOf('Other') !== -1 && onOtherChange !== undefined && (
        <div className={`govuk-form-group${otherError ? ' govuk-form-group--error' : ''}`} style={{ marginTop: '10px', marginBottom: 0 }}>
          {otherError && <p className="govuk-error-message"><span className="govuk-visually-hidden">Error:</span> {otherError}</p>}
          <label className="govuk-label" htmlFor={`${id}Other`}>Please specify</label>
          <input
            className={`govuk-input${otherError ? ' govuk-input--error' : ''}`}
            id={`${id}Other`}
            type="text"
            value={otherValue || ''}
            onChange={e => onOtherChange(e.target.value)}
          />
        </div>
      )}
    </fieldset>
  </div>
);

const DateInput: React.FC<{
  id: string; label: string; hint?: string; error?: string;
  dayValue: string; monthValue: string; yearValue: string;
  onDayChange: (v: string) => void;
  onMonthChange: (v: string) => void;
  onYearChange: (v: string) => void;
}> = ({ id, label, hint = 'For example, 27 3 2025', error, dayValue, monthValue, yearValue, onDayChange, onMonthChange, onYearChange }) => {

  const digitsOnly = (raw: string, maxLen: number): string =>
    raw.replace(/\D/g, '').slice(0, maxLen);

  const clamp = (val: string, min: number, max: number, onChange: (v: string) => void): void => {
    const n = parseInt(val, 10);
    if (!val || isNaN(n)) return;
    if (n < min) onChange(min < 10 ? `0${min}` : `${min}`);
    else if (n > max) onChange(max < 10 ? `0${max}` : `${max}`);
  };

  return (
    <div className={`govuk-form-group${error ? ' govuk-form-group--error' : ''}`}>
      <fieldset className="govuk-fieldset" role="group" aria-describedby={`${id}-hint`}>
        <legend className="govuk-fieldset__legend govuk-fieldset__legend--s">{label}</legend>
        <div id={`${id}-hint`} className="govuk-hint">{hint}</div>
        {error && <p className="govuk-error-message"><span className="govuk-visually-hidden">Error:</span> {error}</p>}
        <div className="govuk-date-input" id={id}>
          <div className="govuk-date-input__item">
            <div className="govuk-form-group">
              <label className="govuk-label govuk-date-input__label" htmlFor={`${id}-day`}>Day</label>
              <input
                className="govuk-input govuk-date-input__input govuk-input--width-2"
                id={`${id}-day`} name={`${id}-day`} type="text" inputMode="numeric"
                value={dayValue} maxLength={2}
                onChange={e => onDayChange(digitsOnly(e.target.value, 2))}
                onBlur={() => clamp(dayValue, 1, 31, onDayChange)}
              />
            </div>
          </div>
          <div className="govuk-date-input__item">
            <div className="govuk-form-group">
              <label className="govuk-label govuk-date-input__label" htmlFor={`${id}-month`}>Month</label>
              <input
                className="govuk-input govuk-date-input__input govuk-input--width-2"
                id={`${id}-month`} name={`${id}-month`} type="text" inputMode="numeric"
                value={monthValue} maxLength={2}
                onChange={e => onMonthChange(digitsOnly(e.target.value, 2))}
                onBlur={() => clamp(monthValue, 1, 12, onMonthChange)}
              />
            </div>
          </div>
          <div className="govuk-date-input__item">
            <div className="govuk-form-group">
              <label className="govuk-label govuk-date-input__label" htmlFor={`${id}-year`}>Year</label>
              <input
                className="govuk-input govuk-date-input__input govuk-input--width-4"
                id={`${id}-year`} name={`${id}-year`} type="text" inputMode="numeric"
                value={yearValue} maxLength={4}
                onChange={e => onYearChange(digitsOnly(e.target.value, 4))}
                onBlur={() => clamp(yearValue, 1900, 2100, onYearChange)}
              />
            </div>
          </div>
        </div>
      </fieldset>
    </div>
  );
};

function formatDate(day: string, month: string, year: string): string {
  if (!day && !month && !year) return '';
  return [day, month, year].filter(Boolean).join('/');
}

function validateDate(day: string, month: string, year: string): string | undefined {
  const d = parseInt(day, 10);
  const m = parseInt(month, 10);
  const y = parseInt(year, 10);
  if (!day || !month || !year) return 'Enter a complete date';
  if (isNaN(d) || d < 1 || d > 31) return 'Enter a valid day between 1 and 31';
  if (isNaN(m) || m < 1 || m > 12) return 'Enter a valid month between 1 and 12';
  if (isNaN(y) || year.length < 4) return 'Year must include 4 numbers';
  const date = new Date(y, m - 1, d);
  if (date.getFullYear() !== y || date.getMonth() !== m - 1 || date.getDate() !== d) {
    return `${day}/${month}/${year} is not a real date`;
  }
  return undefined;
}

const TIME_REGEX = /^([01]\d|2[0-3]):[0-5]\d$/;

const TimeTextInput: React.FC<{
  id: string; label: string; value: string;
  onChange: (v: string) => void; error?: string;
}> = ({ id, label, value, onChange, error }) => {
  const handleChange = (raw: string): void => {
    const digits = raw.replace(/\D/g, '').slice(0, 4);
    if (digits.length <= 2) {
      onChange(digits);
    } else {
      onChange(`${digits.slice(0, 2)}:${digits.slice(2)}`);
    }
  };
  const handleBlur = (): void => {
    if (!value) return;
    const digits = value.replace(/\D/g, '');
    if (digits.length < 3) { onChange(''); return; }
    let h = parseInt(digits.slice(0, 2), 10);
    let m = parseInt(digits.slice(2, 4), 10);
    if (isNaN(h) || isNaN(m)) { onChange(''); return; }
    if (h > 23) h = 23;
    if (m > 59) m = 59;
    onChange(`${h < 10 ? '0' : ''}${h}:${m < 10 ? '0' : ''}${m}`);
  };
  const describedBy = [`${id}-hint`, error ? `${id}-error` : ''].filter(Boolean).join(' ');
  return (
    <div className={`govuk-form-group${error ? ' govuk-form-group--error' : ''}`}>
      <label className="govuk-label" htmlFor={id}>{label}</label>
      <div id={`${id}-hint`} className="govuk-hint">Use 24-hour clock, for example 09:00 or 14:30</div>
      {error && <p id={`${id}-error`} className="govuk-error-message"><span className="govuk-visually-hidden">Error:</span> {error}</p>}
      <input
        className={`govuk-input govuk-input--width-5${error ? ' govuk-input--error' : ''}`}
        id={id} name={id} type="text" inputMode="numeric"
        value={value} maxLength={5}
        aria-describedby={describedBy}
        onChange={e => handleChange(e.target.value)}
        onBlur={handleBlur}
        placeholder="HH:MM"
      />
    </div>
  );
};

// ── People picker ────────────────────────────────────────────────────────────

const PeoplePickerInput: React.FC<{
  nameId: string; emailId: string;
  nameLabel: string; emailLabel: string;
  nameValue: string; emailValue: string;
  onNameChange: (v: string) => void; onEmailChange: (v: string) => void;
  nameError?: string; emailError?: string;
}> = ({ nameId, emailId, nameLabel, emailLabel, nameValue, emailValue, onNameChange, onEmailChange, nameError, emailError }) => {
  const [suggestions, setSuggestions] = React.useState<{ displayName: string; email: string }[]>([]);
  const [isSearching, setIsSearching] = React.useState(false);
  const [open, setOpen] = React.useState(false);
  const debounceRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  const onInput = (value: string): void => {
    onNameChange(value);
    onEmailChange('');
    if (debounceRef.current) clearTimeout(debounceRef.current);
    if (value.trim().length < 2) { setSuggestions([]); setOpen(false); return; }
    debounceRef.current = setTimeout(async () => {
      setIsSearching(true);
      const found = await searchUsers(value);
      setSuggestions(found);
      setOpen(found.length > 0);
      setIsSearching(false);
    }, 350);
  };

  const onPick = (person: { displayName: string; email: string }): void => {
    onNameChange(person.displayName);
    onEmailChange(person.email);
    setSuggestions([]);
    setOpen(false);
  };

  return (
    <>
      <div style={{ position: 'relative' }}>
        <div className={`govuk-form-group${nameError ? ' govuk-form-group--error' : ''}`}>
          <label className="govuk-label" htmlFor={nameId}>{nameLabel}</label>
          <div className="govuk-hint">Start typing to search the organisation address book</div>
          {nameError && <p className="govuk-error-message"><span className="govuk-visually-hidden">Error:</span> {nameError}</p>}
          <input
            className={`govuk-input govuk-input--width-20${nameError ? ' govuk-input--error' : ''}`}
            id={nameId} name={nameId} type="text" autoComplete="off"
            value={nameValue}
            aria-autocomplete="list" aria-expanded={open} aria-haspopup="listbox"
            onChange={e => onInput(e.target.value)}
            onBlur={() => setTimeout(() => setOpen(false), 200)}
          />
          {isSearching && <span className="govuk-hint" style={{ marginTop: '4px', display: 'block' }}>Searching…</span>}
        </div>
        {open && (
          <ul role="listbox" style={{
            position: 'absolute', top: '100%', left: 0, zIndex: 1000,
            background: '#ffffff', border: '2px solid #0b0c0c',
            margin: 0, padding: 0, listStyle: 'none',
            minWidth: '320px', maxHeight: '260px', overflowY: 'auto',
            boxShadow: '0 4px 8px rgba(0,0,0,0.15)',
          }}>
            {suggestions.map(s => (
              <li key={s.email} role="option" aria-selected={false}
                style={{ padding: '10px 15px', cursor: 'pointer', borderBottom: '1px solid #f3f2f1' }}
                onMouseDown={e => { e.preventDefault(); onPick(s); }}
              >
                <span style={{ display: 'block', fontWeight: 700, fontSize: '16px', color: '#0b0c0c' }}>{s.displayName}</span>
                <span style={{ display: 'block', fontSize: '14px', color: '#505a5f' }}>{s.email}</span>
              </li>
            ))}
          </ul>
        )}
      </div>
      <TextInput
        id={emailId} label={emailLabel} value={emailValue} onChange={onEmailChange}
        type="email" hint="Auto-filled when a name is selected above, or enter manually"
        error={emailError}
      />
    </>
  );
};

// ── Step sidebar ─────────────────────────────────────────────────────────────

const StepNav: React.FC<{
  current: number;
  maxReached: number;
  skippedSteps: number[];
  onNavigate: (s: number) => void;
}> = ({ current, maxReached, skippedSteps, onNavigate }) => (
  <nav aria-label="Form progress" style={{ background: '#f3f2f1', padding: '20px', marginBottom: '20px' }}>
    <p className="govuk-body-s" style={{ fontWeight: 'bold', marginBottom: '15px', color: '#0b0c0c' }}>
      {current < REVIEW_STEP ? `Step ${current + 1} of ${REVIEW_STEP}` : 'Review your answers'}
    </p>
    <ol style={{ listStyle: 'none', margin: 0, padding: 0 }}>
      {STEPS.map((stepName, i) => {
        const isSkipped = skippedSteps.indexOf(i) !== -1;
        const isCompleted = i < current && !isSkipped;
        const isCurrent = i === current;
        const isAccessible = i <= maxReached && !isCurrent && !isSkipped;

        return (
          <li
            key={i}
            style={{
              padding: '8px 0',
              borderBottom: i < STEPS.length - 1 ? '1px solid #b1b4b6' : 'none',
              display: 'flex',
              alignItems: 'center',
              gap: '10px',
            }}
          >
            <span
              aria-hidden="true"
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                justifyContent: 'center',
                minWidth: '28px',
                height: '28px',
                borderRadius: '50%',
                background: isSkipped ? '#b1b4b6' : isCurrent ? '#1d70b8' : isCompleted ? '#00703c' : '#b1b4b6',
                color: '#ffffff',
                fontSize: '13px',
                fontWeight: 'bold',
              }}
            >
              {isSkipped ? '–' : isCompleted ? '✓' : i + 1}
            </span>
            {isAccessible ? (
              <a
                href="#"
                className="govuk-link"
                style={{ fontSize: '15px' }}
                onClick={e => { e.preventDefault(); onNavigate(i); }}
                aria-label={`Go to step ${i + 1}: ${stepName}`}
              >
                {stepName}
              </a>
            ) : (
              <span style={{
                fontSize: '15px',
                fontWeight: isCurrent ? 'bold' : 'normal',
                color: isSkipped ? '#b1b4b6' : isCurrent ? '#0b0c0c' : '#505a5f',
              }}>
                {stepName}{isSkipped ? ' — not applicable' : ''}
              </span>
            )}
          </li>
        );
      })}
    </ol>
  </nav>
);

// ── Review helpers ───────────────────────────────────────────────────────────

const ReviewSection: React.FC<{
  title: string;
  stepIndex: number;
  rows: Array<{ label: string; value: string }>;
  onEdit: (s: number) => void;
}> = ({ title, stepIndex, rows, onEdit }) => (
  <div style={{ marginBottom: '30px' }}>
    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', borderBottom: '2px solid #1d70b8', paddingBottom: '8px', marginBottom: '0' }}>
      <h2 className="govuk-heading-s" style={{ margin: 0 }}>{title}</h2>
      <a className="govuk-link govuk-body-s" href="#" onClick={e => { e.preventDefault(); onEdit(stepIndex); }}>Change</a>
    </div>
    <dl className="govuk-summary-list" style={{ marginBottom: 0 }}>
      {rows.map(r => (
        <div key={r.label} className="govuk-summary-list__row">
          <dt className="govuk-summary-list__key" style={{ width: '50%' }}>{r.label}</dt>
          <dd className="govuk-summary-list__value" style={{ color: r.value ? '#0b0c0c' : '#505a5f' }}>
            {r.value || 'Not provided'}
          </dd>
        </div>
      ))}
    </dl>
  </div>
);

// ── Test mode ────────────────────────────────────────────────────────────────
// Injects a red banner into the top of an email HTML body (before the yellow
// accent bar) so test emails are visually distinct. Works because every email
// produced by EmailTemplates.ts includes the <!-- Yellow accent bar --> comment.
function addTestBanner(html: string): string {
  const banner = [
    '        <!-- Test mode banner -->',
    '        <tr>',
    '          <td bgcolor="#d4351c" style="background-color:#d4351c;padding:12px 28px;text-align:center;">',
    '            <p style="margin:0;font-family:Arial,sans-serif;font-size:13px;font-weight:bold;',
    '                      color:#ffffff;letter-spacing:0.5px;">',
    '              &#9888; TEST EMAIL &#8212; DO NOT ACT ON THIS MESSAGE',
    '            </p>',
    '          </td>',
    '        </tr>',
  ].join('\n');
  return html.replace('        <!-- Yellow accent bar -->', banner + '\n        <!-- Yellow accent bar -->');
}

// Shown on-screen whenever TEST_MODE is on, so it disappears automatically
// once the flag is flipped to false ahead of a production deployment.
const TestModeWarning: React.FC = () => (
  <div className="govuk-warning-text" style={{ marginTop: 0, marginBottom: '20px', borderColor: '#d4351c' }}>
    <span className="govuk-warning-text__icon" aria-hidden="true" style={{ borderColor: '#d4351c', color: '#d4351c' }}>!</span>
    <strong className="govuk-warning-text__text" style={{ color: '#d4351c' }}>
      <span className="govuk-visually-hidden">Warning</span>
      This is a test version of this service and is not monitored. Do not use it to report a real data breach — go to the{' '}
      <a className="govuk-link" href={LIVE_SERVICE_URL} target="_blank" rel="noopener noreferrer">live Report a Data Protection Breach service</a> instead.
    </strong>
  </div>
);

// ── Main component ───────────────────────────────────────────────────────────

const DataBreachForm: React.FC<IDataBreachFormProps> = ({ spContext, senderEmail, logicAppUrl, sendMailApiKey, appendCommLogUrl, appendCommLogApiKey, updatePanelRecommendationsUrl, updatePanelRecommendationsApiKey, uploadAttachmentUrl, uploadAttachmentApiKey }) => {

  // ↓ Flip TEST_MODE to false before deploying to production
  const TEST_MODE = false;
  const _logicAppUrl = logicAppUrl || '';
  // APIM subscription key for the sendmail operation — matches the
  // DataBreachSendMailKey Named Value configured in APIM's check-header policy.
  const _sendMailApiKey = sendMailApiKey || '';
  // APIM-fronted Logic App that appends a communication log entry on behalf
  // of a caller who only holds Read at the item level (submitter/person
  // responsible) — see appendCommLogEntryElevated in SharePointService.ts.
  const _appendCommLogUrl = appendCommLogUrl || '';
  const _appendCommLogApiKey = appendCommLogApiKey || '';
  // APIM-fronted Logic App that updates the Panel recommendations section on
  // behalf of a caller who only holds Read at the item level (reviewers) —
  // see updatePanelRecommendationsElevated in SharePointService.ts.
  const _updatePanelRecommendationsUrl = updatePanelRecommendationsUrl || '';
  const _updatePanelRecommendationsApiKey = updatePanelRecommendationsApiKey || '';
  // APIM-fronted Logic App that attaches a file to a report on behalf of the
  // submitter, who only holds Add (not Edit) on the list — see
  // uploadAttachmentElevated in SharePointService.ts.
  const _uploadAttachmentUrl = uploadAttachmentUrl || '';
  const _uploadAttachmentApiKey = uploadAttachmentApiKey || '';
  const _senderEmail = senderEmail || '';
  // Distribution mailbox for new-report notifications — separate from
  // _senderEmail, which is the shared "from"/reply-routing address used
  // elsewhere (e.g. sharedMailbox on ReportViewer), not a notification target.
  const _adminNotificationEmail = ADMIN_NOTIFICATION_EMAIL;
  // NOTE: attachment support here requires a matching change to the sendmail
  // Logic App workflow (external, not tracked in this repo): its "Send an
  // email (V2)" action must map these fields into its Attachments array.
  // Until that's done, the Logic App will just ignore the extra
  // attachmentFilename/attachmentContentType/attachmentContentBytes fields
  // below and send the email without the attachment — it won't error, it'll
  // silently drop it.
  const sendEmailFromMailbox = (to: string, subject: string, body: string, attachment?: IEmailAttachment): Promise<void> => {
    const finalSubject = TEST_MODE ? `[TEST] ${subject}` : subject;
    const finalBody    = TEST_MODE ? addTestBanner(body) : body;
    return fetch(_logicAppUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Ocp-Apim-Subscription-Key': _sendMailApiKey },
      body: JSON.stringify({
        to, subject: finalSubject, body: finalBody,
        ...(attachment ? {
          attachmentFilename: attachment.filename,
          attachmentContentType: attachment.contentType,
          attachmentContentBytes: attachment.base64,
        } : {}),
      }),
    }).then(function(res) {
      if (!res.ok) throw new Error(`Send failed: ${res.status}`);
    });
  };
  const appendCommLogViaFlow = (itemId: number, callerEmail: string, entry: ICommLogEntry): Promise<ICommLogEntry[]> =>
    appendCommLogEntryElevated(_appendCommLogUrl, _appendCommLogApiKey, itemId, callerEmail, entry);
  const updatePanelRecommendationsViaFlow = (itemId: number, callerEmail: string, fields: IPanelRecommendationsFields): Promise<void> =>
    updatePanelRecommendationsElevated(_updatePanelRecommendationsUrl, _updatePanelRecommendationsApiKey, itemId, callerEmail, fields);
  const [mode, setMode] = React.useState<AppMode>(() => {
    const m = window.location.search.match(/[?&]reportId=(\d+)/);
    return m ? 'view' : 'home';
  });
  const [targetReportId, setTargetReportId] = React.useState<number | null>(() => {
    const m = window.location.search.match(/[?&]reportId=(\d+)/);
    return m ? parseInt(m[1], 10) : null;
  });
  const [step, setStep] = React.useState(0);
  const [maxReached, setMaxReached] = React.useState(0);
  const [submitted, setSubmitted] = React.useState(false);
  const [submittedId, setSubmittedId] = React.useState<number | null>(null);
  const [isSubmitting, setIsSubmitting] = React.useState(false);
  const [submitError, setSubmitError] = React.useState('');
  const [errors, setErrors] = React.useState<FormErrors>({});
  const [selfResponsible, setSelfResponsible] = React.useState<string>('');
  const [attachedFiles, setAttachedFiles] = React.useState<File[]>([]);
  const [attachmentErrors, setAttachmentErrors] = React.useState<string[]>([]);
  const errorSummaryRef = React.useRef<HTMLDivElement>(null);

  const [form, setForm] = React.useState<IFormState>({
    officerName: cleanCurrentUserName(spContext.pageContext.user.displayName, spContext.pageContext.user.email || ''),
    officerEmail: spContext.pageContext.user.email || '',
    officerDirectorate: '',
    officerDivision: '',
    responsibleName: '',
    responsibleEmail: '',
    responsibleTraining: 'No',
    responsibleDirectorate: '',
    responsibleDivision: '',
    lineManager: '',
    lineManagerEmail: '',
    personalDataAtRisk: '',
    dataBreachedOutline: '',
    breachDetails: '',
    approxIndividuals: '',
    categoriseNature: [], categoriseNatureOther: '',
    categoriesPersonalData: [], categoriesPersonalDataOther: '',
    categoriesIndividuals: [], categoriesIndividualsOther: '',
    impactToSubjects: '',
    dateDiscoveredDay: '',
    dateDiscoveredMonth: '',
    dateDiscoveredYear: '',
    timeDiscovered: '',
    dateHappenedDay: '',
    dateHappenedMonth: '',
    dateHappenedYear: '',
    timeHappened: '',
    thirdPartyInvolved: 'No',
    thirdPartyName: '',
    thirdPartyAddress: '',
    thirdPartyContact: '',
    thirdPartyEmail: '',
    thirdPartyTelephone: '',
    likelyConsequences: '',
    consequencesExplanation: '',
    stepsTaken: '',
    individualsInformed: '',
    individualsInformedDetails: '',
    evidencePrevent: '',
    thirdPartyAgreement: 'No',
    infoShared: '',
    otherOrgsInformed: 'No',
    otherOrgsDetails: '',
    mediaCovered: 'No',
    mediaDetails: '',
    anyOtherInfo: '',
    anyOtherInfoYesNo: 'No',
    breachCategories: [],
    panelRecommendations: '',
    notificationToIco: 'No',
    reasonForDecision: '',
  });

  const toggleArray = (field: 'categoriseNature' | 'categoriesPersonalData' | 'categoriesIndividuals' | 'breachCategories') => (value: string): void => {
    setForm(prev => {
      const current = prev[field];
      const next = current.indexOf(value) !== -1 ? current.filter(v => v !== value) : current.concat(value);
      return { ...prev, [field]: next };
    });
    setErrors(prev => { const n = { ...prev }; delete n[field]; return n; });
  };

  const set = (field: keyof IFormState) => (value: string): void => {
    setForm(prev => ({ ...prev, [field]: value }));
    setErrors(prev => {
      const next = { ...prev };
      delete next[field];
      return next;
    });
  };

  const validateStep = (currentStep: number): boolean => {
    const newErrors: FormErrors = {};
    const requiredMsg = 'This field is required';
    const emailMsg = 'Enter an email address in the correct format, like name@example.com';
    const phoneMsg = 'Enter a telephone number in the correct format, like 01234 567890';

    if (currentStep === 0) {
      if (!form.officerName.trim()) newErrors.officerName = requiredMsg;
      if (!form.officerDirectorate) newErrors.officerDirectorate = requiredMsg;
      if (form.officerDirectorate && !form.officerDivision) newErrors.officerDivision = requiredMsg;
      if (!form.officerEmail.trim()) newErrors.officerEmail = requiredMsg;
      else if (!EMAIL_REGEX.test(form.officerEmail)) newErrors.officerEmail = emailMsg;
    }
    if (currentStep === 1) {
      if (!form.responsibleName.trim()) newErrors.responsibleName = requiredMsg;
      if (!form.responsibleEmail.trim()) newErrors.responsibleEmail = requiredMsg;
      else if (!EMAIL_REGEX.test(form.responsibleEmail)) newErrors.responsibleEmail = emailMsg;
      if (!form.responsibleDirectorate) newErrors.responsibleDirectorate = requiredMsg;
      if (form.responsibleDirectorate && !form.responsibleDivision) newErrors.responsibleDivision = requiredMsg;
      if (!form.lineManager.trim()) newErrors.lineManager = requiredMsg;
      if (!form.lineManagerEmail.trim()) newErrors.lineManagerEmail = requiredMsg;
      else if (!EMAIL_REGEX.test(form.lineManagerEmail)) newErrors.lineManagerEmail = emailMsg;
    }
    if (currentStep === 2) {
      const timeMsg = 'Enter a valid time in 24-hour format, like 09:00 or 14:30';
      if (!form.personalDataAtRisk) newErrors.personalDataAtRisk = requiredMsg;
      if (!form.breachDetails.trim()) newErrors.breachDetails = requiredMsg;
      if (!form.approxIndividuals) newErrors.approxIndividuals = requiredMsg;
      if (form.categoriseNature.length === 0) newErrors.categoriseNature = 'Select at least one option';
      else if (form.categoriseNature.indexOf('Other') !== -1 && !form.categoriseNatureOther.trim()) newErrors.categoriseNatureOther = requiredMsg;
      if (form.personalDataAtRisk === 'Yes') {
        if (!form.dataBreachedOutline.trim()) newErrors.dataBreachedOutline = requiredMsg;
        if (form.categoriesPersonalData.length === 0) newErrors.categoriesPersonalData = 'Select at least one option';
        else if (form.categoriesPersonalData.indexOf('Other') !== -1 && !form.categoriesPersonalDataOther.trim()) newErrors.categoriesPersonalDataOther = requiredMsg;
        if (form.categoriesIndividuals.length === 0) newErrors.categoriesIndividuals = 'Select at least one option';
        else if (form.categoriesIndividuals.indexOf('Other') !== -1 && !form.categoriesIndividualsOther.trim()) newErrors.categoriesIndividualsOther = requiredMsg;
        if (!form.impactToSubjects) newErrors.impactToSubjects = requiredMsg;
      }
      if (form.breachCategories.length === 0) newErrors.breachCategories = 'Select at least one option';
      const discoveredErr = validateDate(form.dateDiscoveredDay, form.dateDiscoveredMonth, form.dateDiscoveredYear);
      if (discoveredErr) newErrors.dateDiscoveredDay = `Date breach was discovered: ${discoveredErr.toLowerCase()}`;
      const happenedErr = validateDate(form.dateHappenedDay, form.dateHappenedMonth, form.dateHappenedYear);
      if (happenedErr) newErrors.dateHappenedDay = `Date breach happened: ${happenedErr.toLowerCase()}`;
      if (!form.timeDiscovered.trim()) newErrors.timeDiscovered = requiredMsg;
      else if (!TIME_REGEX.test(form.timeDiscovered)) newErrors.timeDiscovered = timeMsg;
      if (!form.timeHappened.trim()) newErrors.timeHappened = requiredMsg;
      else if (!TIME_REGEX.test(form.timeHappened)) newErrors.timeHappened = timeMsg;
    }
    if (currentStep === 3) {
      if (!form.thirdPartyName.trim()) newErrors.thirdPartyName = requiredMsg;
      if (!form.thirdPartyAddress.trim()) newErrors.thirdPartyAddress = requiredMsg;
      if (!form.thirdPartyContact.trim()) newErrors.thirdPartyContact = requiredMsg;
      if (!form.thirdPartyEmail.trim()) newErrors.thirdPartyEmail = requiredMsg;
      else if (!EMAIL_REGEX.test(form.thirdPartyEmail)) newErrors.thirdPartyEmail = emailMsg;
      if (!form.thirdPartyTelephone.trim()) newErrors.thirdPartyTelephone = requiredMsg;
      else if (!PHONE_REGEX.test(form.thirdPartyTelephone)) newErrors.thirdPartyTelephone = phoneMsg;
    }
    if (currentStep === 4) {
      if (!form.likelyConsequences) newErrors.likelyConsequences = requiredMsg;
      if (!form.consequencesExplanation.trim()) newErrors.consequencesExplanation = requiredMsg;
      if (!form.stepsTaken.trim()) newErrors.stepsTaken = requiredMsg;
      if (!form.evidencePrevent.trim()) newErrors.evidencePrevent = requiredMsg;
      if (!form.individualsInformed) newErrors.individualsInformed = requiredMsg;
      if (form.individualsInformed && !form.individualsInformedDetails.trim()) newErrors.individualsInformedDetails = requiredMsg;
      if (!form.mediaCovered) newErrors.mediaCovered = requiredMsg;
      if (form.mediaCovered === 'Yes' && !form.mediaDetails.trim()) newErrors.mediaDetails = requiredMsg;
      if (!form.otherOrgsInformed) newErrors.otherOrgsInformed = requiredMsg;
      if (form.otherOrgsInformed === 'Yes' && !form.otherOrgsDetails.trim()) newErrors.otherOrgsDetails = requiredMsg;
    }
    if (currentStep === 5) {
      if (form.anyOtherInfoYesNo === 'Yes' && !form.anyOtherInfo.trim()) newErrors.anyOtherInfo = requiredMsg;
    }

    setErrors(newErrors);
    if (Object.keys(newErrors).length > 0) {
      setTimeout(() => { if (errorSummaryRef.current) errorSummaryRef.current.focus(); }, 50);
      return false;
    }
    return true;
  };

  const focusField = (key: string): void => {
    const container = document.getElementById(key);
    if (!container) return;
    const target = container.matches('input, select, textarea')
      ? container
      : container.querySelector<HTMLElement>('input, select, textarea') || container;
    target.scrollIntoView({ behavior: 'smooth', block: 'center' });
    target.focus();
  };

  const navigateTo = (target: number): void => {
    setErrors({});
    setStep(target);
    setMaxReached(prev => Math.max(prev, target));
    window.scrollTo(0, 0);
  };

  const next = (): void => {
    if (!validateStep(step)) return;
    const target = step === 2 && form.thirdPartyInvolved === 'No' ? 4 : step + 1;
    navigateTo(target);
  };
  const back = (): void => {
    const target = step === 4 && form.thirdPartyInvolved === 'No' ? 2 : step - 1;
    navigateTo(target);
  };

  const sendSubmissionEmails = (id: number): void => {
    try {
      const ref = `DBN-${id}`;
      const isSelfResponsible = !!(form.responsibleEmail &&
        form.officerEmail && form.officerEmail.toLowerCase() === form.responsibleEmail.toLowerCase());

      const data: IBreachEmailData = {
        ref,
        itemId:            id,
        siteUrl:           spContext.pageContext.web.absoluteUrl,
        officerName:       form.officerName || '',
        officerEmail:      form.officerEmail || '',
        responsibleName:   form.responsibleName || '',
        responsibleEmail:  form.responsibleEmail || '',
        lineManager:       form.lineManager || '',
        lineManagerEmail:  form.lineManagerEmail || '',
        officerDirectorate: form.officerDirectorate || '',
        officerDivision:   form.officerDivision || '',
        dateOccurred:      formatDate(form.dateHappenedDay, form.dateHappenedMonth, form.dateHappenedYear),
        timeOccurred:      form.timeHappened ? ` at ${form.timeHappened}` : '',
        dateDiscovered:    formatDate(form.dateDiscoveredDay, form.dateDiscoveredMonth, form.dateDiscoveredYear),
        timeDiscovered:    form.timeDiscovered ? ` at ${form.timeDiscovered}` : '',
        breachDetails:     form.breachDetails || '',
        approxIndividuals: form.approxIndividuals || '',
        impactToSubjects:  form.impactToSubjects || '',
        natureOfBreach:    (form.categoriseNature || []).join(', '),
        thirdPartyInvolved: form.thirdPartyInvolved || '',
        isSelfResponsible,
      };

      if (form.officerEmail) {
        sendEmailFromMailbox(form.officerEmail, `Data Breach Report Submitted — ${ref}`, buildSubmitterEmail(data))
          .catch(function(err: Error) { console.error('[Email] Submitter email failed:', err); });
      }
      if (form.responsibleEmail) {
        sendEmailFromMailbox(form.responsibleEmail, `Data Breach — You Have Been Identified as the Person Responsible (${ref})`, buildResponsibleEmail(data))
          .catch(function(err: Error) { console.error('[Email] Responsible person email failed:', err); });
      }
      if (form.lineManagerEmail) {
        sendEmailFromMailbox(form.lineManagerEmail, `Data Breach Notification — ${form.responsibleName} (${ref})`, buildLineManagerEmail(data))
          .catch(function(err: Error) { console.error('[Email] Line manager email failed:', err); });
      }
      sendEmailFromMailbox(_adminNotificationEmail, `New Data Breach Report Received — ${ref}`, buildAdminEmail(data))
        .catch(function(err: Error) { console.error('[Email] Admin email failed:', err); });
    } catch (err) {
      console.error('[Email] sendSubmissionEmails threw synchronously:', err);
    }
  };

  // Reads a File as raw base64 (no "data:...;base64," prefix) for posting to
  // the upload-attachment Logic App — same convention as ReportPdf.ts's
  // buildReportPdfBase64.
  const fileToBase64 = (file: File): Promise<string> => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result as string;
      resolve(result.substring(result.indexOf('base64,') + 'base64,'.length));
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });

  const uploadSubmissionAttachments = (itemId: number): void => {
    if (!attachedFiles.length) return;
    const callerEmail = spContext.pageContext.user.email || '';
    attachedFiles.forEach(function(file) {
      fileToBase64(file)
        .then(function(contentBytes) {
          return uploadAttachmentElevated(_uploadAttachmentUrl, _uploadAttachmentApiKey, itemId, callerEmail, {
            filename: file.name,
            contentType: file.type || 'application/octet-stream',
            contentBytes,
          });
        })
        .catch(function(err: Error) { console.error(`[Attachments] Upload failed for "${file.name}":`, err); });
    });
  };

  const handleSubmit = (): void => {
    setSubmitError('');
    setIsSubmitting(true);
    createBreachRecord(form)
      .then(id => {
        setSubmittedId(id);
        setSubmitted(true);
        setIsSubmitting(false);
        window.scrollTo(0, 0);
        sendSubmissionEmails(id);
        uploadSubmissionAttachments(id);
      })
      .catch(() => {
        setSubmitError('There was a problem submitting the report. Please try again.');
        setIsSubmitting(false);
      });
  };

  const resetForm = (): void => {
    setMode('home');
    setStep(0);
    setMaxReached(0);
    setSubmitted(false);
    setSubmittedId(null);
    setIsSubmitting(false);
    setSubmitError('');
    setErrors({});
    setSelfResponsible('');
    setAttachedFiles([]);
    setAttachmentErrors([]);
    setForm({
      officerName: cleanCurrentUserName(spContext.pageContext.user.displayName, spContext.pageContext.user.email || ''),
      officerEmail: spContext.pageContext.user.email || '',
      officerDirectorate: '', officerDivision: '', responsibleName: '',
      responsibleEmail: '', responsibleTraining: 'No', responsibleDirectorate: '',
      responsibleDivision: '', lineManager: '', lineManagerEmail: '',
      personalDataAtRisk: '', dataBreachedOutline: '', breachDetails: '', approxIndividuals: '',
      categoriseNature: [], categoriseNatureOther: '',
      categoriesPersonalData: [], categoriesPersonalDataOther: '',
      categoriesIndividuals: [], categoriesIndividualsOther: '', impactToSubjects: '',
      dateDiscoveredDay: '', dateDiscoveredMonth: '', dateDiscoveredYear: '',
      timeDiscovered: '', dateHappenedDay: '', dateHappenedMonth: '',
      dateHappenedYear: '', timeHappened: '', thirdPartyInvolved: 'No', thirdPartyName: '',
      thirdPartyAddress: '', thirdPartyContact: '', thirdPartyEmail: '',
      thirdPartyTelephone: '', likelyConsequences: '', consequencesExplanation: '',
      stepsTaken: '', individualsInformed: '', individualsInformedDetails: '',
      evidencePrevent: '', thirdPartyAgreement: 'No', infoShared: '',
      otherOrgsInformed: 'No', otherOrgsDetails: '', mediaCovered: 'No',
      mediaDetails: '', anyOtherInfo: '', anyOtherInfoYesNo: 'No', breachCategories: [],
      panelRecommendations: '', notificationToIco: 'No', reasonForDecision: '',
    });
  };

  if (mode === 'home') {
    return (
      <div className={styles.dataBreachForm}>
        <div className="govuk-width-container">
          <div className="govuk-phase-banner">
            <p className="govuk-phase-banner__content">
              <strong className="govuk-tag govuk-phase-banner__content__tag">Beta</strong>
              <span className="govuk-phase-banner__text">Data Breach Reporting</span>
            </p>
          </div>
          <main className="govuk-main-wrapper" id="main-content">
            <div className="govuk-grid-row">
              <div className="govuk-grid-column-two-thirds">
                {TEST_MODE && <TestModeWarning />}
                <h1 className="govuk-heading-xl">Data Breach Reporting</h1>
                <p className="govuk-body-l">Use this service to report a data breach or view your existing reports.</p>
                <hr className="govuk-section-break govuk-section-break--m govuk-section-break--visible" />
                <h2 className="govuk-heading-m">Submit a new report</h2>
                <p className="govuk-body">Report a new data breach incident. You will be asked to provide details about the breach, the individuals affected, and the steps taken.</p>
                <button type="button" className="govuk-button" onClick={() => setMode('submit')}>
                  Report a data breach
                </button>
                <hr className="govuk-section-break govuk-section-break--m govuk-section-break--visible" />
                <h2 className="govuk-heading-m">View existing reports</h2>
                <p className="govuk-body">View and edit breach reports you have previously submitted.</p>
                <button type="button" className="govuk-button govuk-button--secondary" onClick={() => setMode('view')}>
                  View reports
                </button>
              </div>
            </div>
          </main>
        </div>
      </div>
    );
  }

  if (mode === 'view') {
    return (
      <div className={styles.dataBreachForm}>
        <div className="govuk-width-container">
          <div className="govuk-phase-banner">
            <p className="govuk-phase-banner__content">
              <strong className="govuk-tag govuk-phase-banner__content__tag">Beta</strong>
              <span className="govuk-phase-banner__text">Data Breach Reporting</span>
            </p>
          </div>
          <main className="govuk-main-wrapper" id="main-content">
            <a href="#" className="govuk-back-link" onClick={e => { e.preventDefault(); setMode('home'); }}>Back to home</a>
            {TEST_MODE && <TestModeWarning />}
            <ReportViewer spContext={spContext} sharedMailbox={_senderEmail} sendEmailFromMailbox={sendEmailFromMailbox} targetItemId={targetReportId} onTargetItemOpened={() => setTargetReportId(null)} testMode={TEST_MODE} appendCommLogEntryElevated={appendCommLogViaFlow} updatePanelRecommendationsElevated={updatePanelRecommendationsViaFlow} />
          </main>
        </div>
      </div>
    );
  }

  if (submitted) {
    return (
      <div className={styles.dataBreachForm}>
        <div className="govuk-width-container">
          <main className="govuk-main-wrapper">
            {TEST_MODE && <TestModeWarning />}
            <div className="govuk-panel govuk-panel--confirmation">
              <h1 className="govuk-panel__title">Report submitted</h1>
              {submittedId !== null && (
                <div className="govuk-panel__body">Reference: <strong>DBN-{submittedId}</strong></div>
              )}
            </div>
            <p className="govuk-body" style={{ marginTop: '30px' }}>Your data breach report has been submitted successfully.</p>
            <button type="button" className="govuk-button" onClick={resetForm}>
              Submit another report
            </button>
            <button type="button" className="govuk-button govuk-button--secondary" style={{ marginLeft: '15px' }} onClick={() => { resetForm(); setMode('home'); }}>
              Back to home
            </button>
          </main>
        </div>
      </div>
    );
  }

  const isReview = step === REVIEW_STEP;

  return (
    <div className={styles.dataBreachForm}>
      <div className="govuk-width-container">

        {/* Phase banner */}
        <div className="govuk-phase-banner">
          <p className="govuk-phase-banner__content">
            <strong className="govuk-tag govuk-phase-banner__content__tag">Beta</strong>
            <span className="govuk-phase-banner__text">Data Breach Reporting — DBN-7</span>
          </p>
        </div>

        <main className="govuk-main-wrapper" id="main-content">

          {TEST_MODE && <TestModeWarning />}

          {/* Back link */}
          {step === 0 ? (
            <a href="#" className="govuk-back-link" onClick={e => { e.preventDefault(); setMode('home'); }}>Back to home</a>
          ) : (
            <a href="#" className="govuk-back-link" onClick={e => { e.preventDefault(); back(); }}>Back</a>
          )}

          {/* Error summary */}
          {Object.keys(errors).some(k => Boolean(errors[k as keyof IFormState])) && (
            <div
              className="govuk-error-summary"
              data-module="govuk-error-summary"
              tabIndex={-1}
              ref={errorSummaryRef}
            >
              <div role="alert">
                <h2 className="govuk-error-summary__title">There is a problem</h2>
                <div className="govuk-error-summary__body">
                  <ul className="govuk-list govuk-error-summary__list">
                    {Object.keys(errors)
                      .filter(k => Boolean(errors[k as keyof IFormState]))
                      .map(k => (
                        <li key={k}>
                          <a href={`#${k}`} onClick={e => { e.preventDefault(); focusField(k); }}>{errors[k as keyof IFormState]}</a>
                        </li>
                      ))}
                  </ul>
                </div>
              </div>
            </div>
          )}

          <div className="govuk-grid-row">

            {/* Step navigation sidebar */}
            <div className="govuk-grid-column-one-third">
              <StepNav current={step} maxReached={maxReached} skippedSteps={form.thirdPartyInvolved === 'No' ? [3] : []} onNavigate={navigateTo} />
            </div>

            {/* Form content */}
            <div className="govuk-grid-column-two-thirds">

              {/* Step 0: Reporting Officer */}
              {step === 0 && (
                <div>
                  <span className="govuk-caption-l">Step 1 of {REVIEW_STEP}</span>
                  <h1 className="govuk-heading-l">Reporting Officer</h1>
                  <TextInput id="officerName" label="Name of the officer reporting the breach" value={form.officerName} onChange={set('officerName')} error={errors.officerName} readOnly />
                  <TextInput id="officerEmail" label="Email of the reporting officer" value={form.officerEmail} onChange={set('officerEmail')} type="email" error={errors.officerEmail} readOnly />
                  <SelectInput id="officerDirectorate" label="Directorate of reporting officer" value={form.officerDirectorate} onChange={v => { set('officerDirectorate')(v); set('officerDivision')(''); }} options={DIRECTORATES} error={errors.officerDirectorate} />
                  <SelectInput
                    id="officerDivision"
                    label="Division of reporting officer"
                    value={form.officerDivision}
                    onChange={set('officerDivision')}
                    options={getDivisions(form.officerDirectorate)}
                    disabled={!form.officerDirectorate}
                    disabledHint="Select a directorate first"
                  />
                </div>
              )}

              {/* Step 1: Person Responsible */}
              {step === 1 && (
                <div>
                  <span className="govuk-caption-l">Step 2 of {REVIEW_STEP}</span>
                  <h1 className="govuk-heading-l">Person Responsible For Data Breach</h1>
                  <RadioGroup
                    name="selfResponsible"
                    label="Are you the person responsible for this data breach?"
                    value={selfResponsible}
                    options={['Yes', 'No']}
                    onChange={v => {
                      setSelfResponsible(v);
                      if (v === 'Yes') {
                        setForm(prev => ({
                          ...prev,
                          responsibleName: form.officerName,
                          responsibleEmail: form.officerEmail,
                          responsibleDirectorate: form.officerDirectorate,
                          responsibleDivision: form.officerDivision,
                        }));
                        setErrors(prev => { const n = { ...prev }; delete n.responsibleName; delete n.responsibleEmail; return n; });
                      } else {
                        setForm(prev => ({ ...prev, responsibleName: '', responsibleEmail: '', responsibleDirectorate: '', responsibleDivision: '' }));
                      }
                    }}
                  />
                  {selfResponsible === 'No' && (
                    <>
                      <PeoplePickerInput
                        nameId="responsibleName" emailId="responsibleEmail"
                        nameLabel="Name of person responsible"
                        emailLabel="Email address of person responsible"
                        nameValue={form.responsibleName} emailValue={form.responsibleEmail}
                        onNameChange={set('responsibleName')} onEmailChange={set('responsibleEmail')}
                        nameError={errors.responsibleName} emailError={errors.responsibleEmail}
                      />
                      <SelectInput id="responsibleDirectorate" label="Directorate of person responsible" value={form.responsibleDirectorate} onChange={v => { set('responsibleDirectorate')(v); set('responsibleDivision')(''); }} options={DIRECTORATES} error={errors.responsibleDirectorate} />
                      <SelectInput
                        id="responsibleDivision"
                        label="Division of person responsible"
                        value={form.responsibleDivision}
                        onChange={set('responsibleDivision')}
                        options={getDivisions(form.responsibleDirectorate)}
                        disabled={!form.responsibleDirectorate}
                        disabledHint="Select a directorate first"
                      />
                    </>
                  )}
                  {selfResponsible !== '' && (
                    <PeoplePickerInput
                      nameId="lineManager" emailId="lineManagerEmail"
                      nameLabel={selfResponsible === 'Yes' ? 'Your line manager' : 'Line manager of person responsible'}
                      emailLabel={selfResponsible === 'Yes' ? 'Your line manager\'s email address' : 'Email address of line manager'}
                      nameValue={form.lineManager} emailValue={form.lineManagerEmail}
                      onNameChange={set('lineManager')} onEmailChange={set('lineManagerEmail')}
                      nameError={errors.lineManager} emailError={errors.lineManagerEmail}
                    />
                  )}
                </div>
              )}

              {/* Step 2: Breach Details */}
              {step === 2 && (
                <div>
                  <span className="govuk-caption-l">Step 3 of {REVIEW_STEP}</span>
                  <h1 className="govuk-heading-l">Breach Details</h1>
                  <TextArea id="breachDetails" label="Details of the breach" value={form.breachDetails} onChange={set('breachDetails')} error={errors.breachDetails} />
                  <RadioGroup
                    name="personalDataAtRisk"
                    label="Has any personal data been placed at risk?"
                    value={form.personalDataAtRisk}
                    options={['Yes', 'No']}
                    error={errors.personalDataAtRisk}
                    onChange={v => {
                      setForm(prev => ({
                        ...prev,
                        personalDataAtRisk: v,
                        ...(v === 'No' ? { dataBreachedOutline: '', categoriesPersonalData: [], categoriesPersonalDataOther: '', categoriesIndividuals: [], categoriesIndividualsOther: '', impactToSubjects: '' } : {}),
                      }));
                      setErrors(prev => { const n = { ...prev }; delete n.personalDataAtRisk; if (v === 'No') { delete n.dataBreachedOutline; delete n.categoriesPersonalData; delete n.categoriesIndividuals; delete n.impactToSubjects; } return n; });
                    }}
                  />
                  {form.personalDataAtRisk === 'No' && (
                    <div className="govuk-inset-text">
                      If no personal data has been placed at risk, this incident may not be classified as a data breach and the risk to individuals may be low. You can still continue to report the incident.
                    </div>
                  )}
                  {form.personalDataAtRisk === 'Yes' && (
                    <div style={{ borderLeft: '4px solid #b1b4b6', paddingLeft: '15px', marginBottom: '20px' }}>
                      <TextArea id="dataBreachedOutline" label="Outline of what the data breached consists of" value={form.dataBreachedOutline} onChange={set('dataBreachedOutline')} error={errors.dataBreachedOutline} />
                      <CheckboxGroup
                        id="categoriesPersonalData" label="Categories of personal data in breach (select all that apply)"
                        values={form.categoriesPersonalData} options={CATEGORIES_PERSONAL_DATA} error={errors.categoriesPersonalData}
                        onToggle={v => setForm(prev => { const cur = prev.categoriesPersonalData; const next = cur.indexOf(v) !== -1 ? cur.filter(x => x !== v) : cur.concat(v); return { ...prev, categoriesPersonalData: next, ...(v === 'Other' && cur.indexOf(v) !== -1 ? { categoriesPersonalDataOther: '' } : {}) }; })}
                        otherValue={form.categoriesPersonalDataOther} onOtherChange={set('categoriesPersonalDataOther')} otherError={errors.categoriesPersonalDataOther}
                      />
                      <CheckboxGroup
                        id="categoriesIndividuals" label="Categories of individuals in breach (select all that apply)"
                        values={form.categoriesIndividuals} options={CATEGORIES_INDIVIDUALS} error={errors.categoriesIndividuals}
                        onToggle={v => setForm(prev => { const cur = prev.categoriesIndividuals; const next = cur.indexOf(v) !== -1 ? cur.filter(x => x !== v) : cur.concat(v); return { ...prev, categoriesIndividuals: next, ...(v === 'Other' && cur.indexOf(v) !== -1 ? { categoriesIndividualsOther: '' } : {}) }; })}
                        otherValue={form.categoriesIndividualsOther} onOtherChange={set('categoriesIndividualsOther')} otherError={errors.categoriesIndividualsOther}
                      />
                      <SelectInput id="impactToSubjects" label="Impact to Data Subjects" value={form.impactToSubjects} onChange={set('impactToSubjects')} options={IMPACT_TO_SUBJECTS} error={errors.impactToSubjects} />
                    </div>
                  )}
                  <SelectInput id="approxIndividuals" label="Approx individuals affected" value={form.approxIndividuals} onChange={set('approxIndividuals')} options={APPROX_INDIVIDUALS} error={errors.approxIndividuals} />
                  <CheckboxGroup
                    id="categoriseNature" label="Categorise nature of breach (select all that apply)"
                    values={form.categoriseNature} options={NATURE_OF_BREACH} error={errors.categoriseNature}
                    onToggle={v => setForm(prev => { const cur = prev.categoriseNature; const next = cur.indexOf(v) !== -1 ? cur.filter(x => x !== v) : cur.concat(v); return { ...prev, categoriseNature: next, ...(v === 'Other' && cur.indexOf(v) !== -1 ? { categoriseNatureOther: '' } : {}) }; })}
                    otherValue={form.categoriseNatureOther} onOtherChange={set('categoriseNatureOther')} otherError={errors.categoriseNatureOther}
                  />
                  <CheckboxGroup id="breachCategories" label="How did the breach occur? (select all that apply)" values={form.breachCategories} options={BREACH_CATEGORIES} onToggle={toggleArray('breachCategories')} error={errors.breachCategories} />
                  <DateInput
                    id="dateDiscovered" label="Date breach discovered"
                    dayValue={form.dateDiscoveredDay} monthValue={form.dateDiscoveredMonth} yearValue={form.dateDiscoveredYear}
                    onDayChange={set('dateDiscoveredDay')} onMonthChange={set('dateDiscoveredMonth')} onYearChange={set('dateDiscoveredYear')}
                    error={errors.dateDiscoveredDay}
                  />
                  <TimeTextInput id="timeDiscovered" label="Time breach discovered" value={form.timeDiscovered} onChange={set('timeDiscovered')} error={errors.timeDiscovered} />
                  <DateInput
                    id="dateHappened" label="Date breach happened"
                    dayValue={form.dateHappenedDay} monthValue={form.dateHappenedMonth} yearValue={form.dateHappenedYear}
                    onDayChange={set('dateHappenedDay')} onMonthChange={set('dateHappenedMonth')} onYearChange={set('dateHappenedYear')}
                    error={errors.dateHappenedDay}
                  />
                  <TimeTextInput id="timeHappened" label="Time breach happened" value={form.timeHappened} onChange={set('timeHappened')} error={errors.timeHappened} />
                  <RadioGroup
                    name="thirdPartyInvolved"
                    label="Did the breach involve a third party?"
                    value={form.thirdPartyInvolved}
                    onChange={v => {
                      setForm(prev => ({
                        ...prev,
                        thirdPartyInvolved: v,
                        ...(v === 'No' ? {
                          thirdPartyAgreement: 'No',
                          thirdPartyName: '', thirdPartyAddress: '', thirdPartyContact: '',
                          thirdPartyEmail: '', thirdPartyTelephone: '',
                        } : {}),
                      }));
                      setErrors(prev => { const n = { ...prev }; delete n.thirdPartyInvolved; return n; });
                    }}
                    options={['Yes', 'No']}
                  />
                  {form.thirdPartyInvolved === 'Yes' && (
                    <RadioGroup
                      name="thirdPartyAgreement"
                      label="Does the breach relate to an agreement with a third party?"
                      value={form.thirdPartyAgreement}
                      onChange={set('thirdPartyAgreement')}
                      options={['Yes', 'No']}
                    />
                  )}
                </div>
              )}

              {/* Step 3: Third Party */}
              {step === 3 && (
                <div>
                  <span className="govuk-caption-l">Step 4 of {REVIEW_STEP}</span>
                  <h1 className="govuk-heading-l">Third Party Details</h1>
                  <TextInput id="thirdPartyName" label="Name of third party" value={form.thirdPartyName} onChange={set('thirdPartyName')} error={errors.thirdPartyName} />
                  <TextArea id="thirdPartyAddress" label="Address of third party" value={form.thirdPartyAddress} onChange={set('thirdPartyAddress')} rows={3} error={errors.thirdPartyAddress} />
                  <TextInput id="thirdPartyContact" label="Contact for third party" value={form.thirdPartyContact} onChange={set('thirdPartyContact')} error={errors.thirdPartyContact} />
                  <TextInput id="thirdPartyEmail" label="Email for third party" value={form.thirdPartyEmail} onChange={set('thirdPartyEmail')} type="email" error={errors.thirdPartyEmail} />
                  <TextInput id="thirdPartyTelephone" label="Telephone for third party" value={form.thirdPartyTelephone} onChange={set('thirdPartyTelephone')} type="tel" error={errors.thirdPartyTelephone} />
                </div>
              )}

              {/* Step 4: Consequences & Actions */}
              {step === 4 && (
                <div>
                  <span className="govuk-caption-l">Step 5 of {REVIEW_STEP}</span>
                  <h1 className="govuk-heading-l">Consequences &amp; Actions</h1>
                  <RadioGroup name="responsibleTraining" label="Has the person responsible completed mandatory data protection training?" value={form.responsibleTraining} onChange={set('responsibleTraining')} options={['Yes', 'No']} />
                  <SelectInput id="likelyConsequences" label="What is the likelihood that the individuals affected will experience significant consequences because of the data breach?" value={form.likelyConsequences} onChange={set('likelyConsequences')} options={LIKELIHOOD_OF_CONSEQUENCES} error={errors.likelyConsequences} />
                  <TextArea id="consequencesExplanation" label="If not yet known — please provide details" value={form.consequencesExplanation} onChange={set('consequencesExplanation')} rows={3} error={errors.consequencesExplanation} />
                  <TextArea id="stepsTaken" label="Steps taken for resolution" value={form.stepsTaken} onChange={set('stepsTaken')} rows={4} error={errors.stepsTaken} />
                  <TextArea id="evidencePrevent" label="Please evidence any action you have taken to prevent similar incidents in the future" value={form.evidencePrevent} onChange={set('evidencePrevent')} rows={3} error={errors.evidencePrevent} />
                  <SelectInput
                    id="individualsInformed"
                    label="Have you informed the affected individuals that this incident has occurred?"
                    value={form.individualsInformed}
                    options={INDIVIDUALS_INFORMED_OPTIONS}
                    error={errors.individualsInformed}
                    onChange={v => {
                      setForm(prev => ({ ...prev, individualsInformed: v }));
                      setErrors(prev => { const n = { ...prev }; delete n.individualsInformed; return n; });
                    }}
                  />
                  {form.individualsInformed && (
                    <TextArea
                      id="individualsInformedDetails"
                      label={
                        form.individualsInformed === 'Yes' ? "Give details of how and when" :
                          form.individualsInformed === 'No' ? "Give details of why you haven't" :
                            form.individualsInformed === "We're about to" ? "Please give details of how and when you plan to inform them" :
                              "Please give details of why this is not applicable"
                      }
                      value={form.individualsInformedDetails}
                      onChange={set('individualsInformedDetails')}
                      rows={3}
                      error={errors.individualsInformedDetails}
                    />
                  )}
                  <RadioGroup
                    name="mediaCovered"
                    label="Has the DB caused any media attention?"
                    value={form.mediaCovered}
                    options={['Yes', 'No']}
                    error={errors.mediaCovered}
                    onChange={v => {
                      setForm(prev => ({ ...prev, mediaCovered: v, ...(v === 'No' ? { mediaDetails: '' } : {}) }));
                      setErrors(prev => { const n = { ...prev }; delete n.mediaCovered; if (v === 'No') delete n.mediaDetails; return n; });
                    }}
                  />
                  {form.mediaCovered === 'Yes' && (
                    <TextArea id="mediaDetails" label="Please provide details" value={form.mediaDetails} onChange={set('mediaDetails')} rows={3} error={errors.mediaDetails} />
                  )}
                  <RadioGroup
                    name="otherOrgsInformed"
                    label="Have you told, or are you planning to tell any other organisations about the breach? e.g. the police, other regulators or supervisory authorities."
                    value={form.otherOrgsInformed}
                    options={['Yes', 'No']}
                    error={errors.otherOrgsInformed}
                    onChange={v => {
                      setForm(prev => ({ ...prev, otherOrgsInformed: v, ...(v === 'No' ? { otherOrgsDetails: '' } : {}) }));
                      setErrors(prev => { const n = { ...prev }; delete n.otherOrgsInformed; if (v === 'No') delete n.otherOrgsDetails; return n; });
                    }}
                  />
                  {form.otherOrgsInformed === 'Yes' && (
                    <TextArea id="otherOrgsDetails" label="Please provide details" value={form.otherOrgsDetails} onChange={set('otherOrgsDetails')} rows={3} error={errors.otherOrgsDetails} />
                  )}
                </div>
              )}

              {/* Step 5: Supporting Information & Attachments */}
              {step === 5 && (
                <div>
                  <span className="govuk-caption-l">Step 6 of {REVIEW_STEP}</span>
                  <h1 className="govuk-heading-l">Supporting Information &amp; Attachments</h1>
                  <RadioGroup name="anyOtherInfoYesNo" label="Is there any other information you wish to add?" value={form.anyOtherInfoYesNo} onChange={v => {
                    setForm(prev => ({ ...prev, anyOtherInfoYesNo: v, ...(v === 'No' ? { anyOtherInfo: '' } : {}) }));
                    setErrors(prev => { const n = { ...prev }; delete n.anyOtherInfoYesNo; if (v === 'No') delete n.anyOtherInfo; return n; });
                  }} options={['Yes', 'No']} />
                  {form.anyOtherInfoYesNo === 'Yes' && (
                    <TextArea id="anyOtherInfo" label="Please provide any other relevant information" value={form.anyOtherInfo} onChange={set('anyOtherInfo')} rows={4} error={errors.anyOtherInfo} />
                  )}
                  <div className={`govuk-form-group${attachmentErrors.length ? ' govuk-form-group--error' : ''}`} style={{ marginTop: '20px' }}>
                    <label className="govuk-label" htmlFor="attachments">Upload supporting files (optional)</label>
                    <div className="govuk-hint" id="attachments-hint">
                      You can upload multiple files. Each file must be smaller than {MAX_ATTACHMENT_MB}MB and be one of: {ALLOWED_ATTACHMENT_EXTENSIONS.join(', ')}.
                    </div>
                    {attachmentErrors.map((msg, i) => (
                      <p key={i} className="govuk-error-message">
                        <span className="govuk-visually-hidden">Error:</span> {msg}
                      </p>
                    ))}
                    <input
                      className={`govuk-file-upload${attachmentErrors.length ? ' govuk-file-upload--error' : ''}`}
                      id="attachments"
                      name="attachments"
                      type="file"
                      multiple
                      accept={ALLOWED_ATTACHMENT_EXTENSIONS.map(ext => '.' + ext).join(',')}
                      aria-describedby="attachments-hint"
                      onChange={e => {
                        // accept= only filters the picker dialog (users can
                        // switch it to "All files"), so check every file here too.
                        // The upload-attachment Logic App is the real enforcement.
                        const fileList = e.target.files;
                        const picked: File[] = [];
                        const rejected: string[] = [];
                        if (fileList) {
                          for (let i = 0; i < fileList.length; i++) {
                            const file = fileList[i];
                            const dot = file.name.lastIndexOf('.');
                            const ext = dot !== -1 ? file.name.slice(dot + 1).toLowerCase() : '';
                            if (ALLOWED_ATTACHMENT_EXTENSIONS.indexOf(ext) === -1) {
                              rejected.push(`${file.name} must be one of: ${ALLOWED_ATTACHMENT_EXTENSIONS.join(', ')}`);
                            } else if (file.size > MAX_ATTACHMENT_MB * 1024 * 1024) {
                              rejected.push(`${file.name} must be smaller than ${MAX_ATTACHMENT_MB}MB`);
                            } else if (file.size === 0) {
                              rejected.push(`${file.name} is empty`);
                            } else {
                              picked.push(file);
                            }
                          }
                        }
                        setAttachmentErrors(rejected);
                        setAttachedFiles(prev => [...prev, ...picked]);
                        e.target.value = '';
                      }}
                    />
                    {attachedFiles.length > 0 && (
                      <ul className="govuk-list" style={{ marginTop: '10px' }}>
                        {attachedFiles.map((file, i) => (
                          <li key={`${file.name}-${i}`}>
                            {file.name}{' '}
                            <a
                              className="govuk-link"
                              href="#"
                              onClick={e => {
                                e.preventDefault();
                                setAttachedFiles(prev => prev.filter((_, idx) => idx !== i));
                              }}
                            >
                              Remove
                            </a>
                          </li>
                        ))}
                      </ul>
                    )}
                  </div>
                </div>
              )}

              {/* Review */}
              {isReview && (
                <div>
                  <h1 className="govuk-heading-l">Check your answers</h1>
                  <p className="govuk-body">Review all the information below before submitting your report.</p>

                  <ReviewSection title="Reporting Officer" stepIndex={0} onEdit={navigateTo} rows={[
                    { label: 'Officer name', value: form.officerName },
                    { label: 'Officer email', value: form.officerEmail },
                    { label: 'Directorate', value: form.officerDirectorate },
                    { label: 'Division', value: form.officerDivision },
                  ]} />
                  <ReviewSection title="Person Responsible" stepIndex={1} onEdit={navigateTo} rows={[
                    { label: 'Name', value: form.responsibleName },
                    { label: 'Email', value: form.responsibleEmail },
                    { label: 'Directorate', value: form.responsibleDirectorate },
                    { label: 'Division', value: form.responsibleDivision },
                    { label: 'Line manager', value: form.lineManager },
                    { label: 'Line manager email', value: form.lineManagerEmail },
                  ]} />
                  <ReviewSection title="Breach Details" stepIndex={2} onEdit={navigateTo} rows={[
                    { label: 'Details of the breach', value: form.breachDetails },
                    { label: 'Personal data placed at risk', value: form.personalDataAtRisk },
                    ...(form.personalDataAtRisk === 'Yes' ? [{ label: 'Outline of what the data breached consists of', value: form.dataBreachedOutline }] : []),
                    { label: 'Approx individuals affected', value: form.approxIndividuals },
                    { label: 'Nature of breach', value: form.categoriseNature.join(', ') },
                    ...(form.categoriseNature.indexOf('Other') !== -1 && form.categoriseNatureOther ? [{ label: 'Nature of breach — other', value: form.categoriseNatureOther }] : []),
                    ...(form.personalDataAtRisk === 'Yes' ? [
                      { label: 'Categories of personal data', value: form.categoriesPersonalData.join(', ') },
                      ...(form.categoriesPersonalData.indexOf('Other') !== -1 && form.categoriesPersonalDataOther ? [{ label: 'Personal data categories — other', value: form.categoriesPersonalDataOther }] : []),
                      { label: 'Categories of individuals', value: form.categoriesIndividuals.join(', ') },
                      ...(form.categoriesIndividuals.indexOf('Other') !== -1 && form.categoriesIndividualsOther ? [{ label: 'Individuals categories — other', value: form.categoriesIndividualsOther }] : []),
                      { label: 'Impact to data subjects', value: form.impactToSubjects },
                    ] : []),
                    { label: 'How did the breach occur', value: form.breachCategories.join(', ') },
                    { label: 'Date discovered', value: formatDate(form.dateDiscoveredDay, form.dateDiscoveredMonth, form.dateDiscoveredYear) },
                    { label: 'Time discovered', value: form.timeDiscovered },
                    { label: 'Date happened', value: formatDate(form.dateHappenedDay, form.dateHappenedMonth, form.dateHappenedYear) },
                    { label: 'Time happened', value: form.timeHappened },
                    { label: 'Third party involved', value: form.thirdPartyInvolved },
                    ...(form.thirdPartyInvolved === 'Yes' ? [{ label: 'Breach relates to third party agreement', value: form.thirdPartyAgreement }] : []),
                  ]} />
                  {form.thirdPartyInvolved === 'Yes' && (
                    <ReviewSection title="Third Party" stepIndex={3} onEdit={navigateTo} rows={[
                      { label: 'Name', value: form.thirdPartyName },
                      { label: 'Address', value: form.thirdPartyAddress },
                      { label: 'Contact', value: form.thirdPartyContact },
                      { label: 'Email', value: form.thirdPartyEmail },
                      { label: 'Telephone', value: form.thirdPartyTelephone },
                    ]} />
                  )}
                  <ReviewSection title="Consequences & Actions" stepIndex={4} onEdit={navigateTo} rows={[
                    { label: 'Mandatory training completed', value: form.responsibleTraining },
                    { label: 'Likelihood of significant consequences', value: form.likelyConsequences },
                    { label: 'If not yet known — details', value: form.consequencesExplanation },
                    { label: 'Steps taken', value: form.stepsTaken },
                    { label: 'Action taken to prevent similar incidents', value: form.evidencePrevent },
                    { label: 'Affected individuals informed of incident', value: form.individualsInformed },
                    ...(form.individualsInformedDetails ? [{ label: 'Details', value: form.individualsInformedDetails }] : []),
                    { label: 'Has the DB caused any media attention?', value: form.mediaCovered },
                    ...(form.mediaCovered === 'Yes' ? [{ label: 'Media attention details', value: form.mediaDetails }] : []),
                    { label: 'Other organisations informed about the breach?', value: form.otherOrgsInformed },
                    ...(form.otherOrgsInformed === 'Yes' ? [{ label: 'Other organisations details', value: form.otherOrgsDetails }] : []),
                  ]} />
                  <ReviewSection title="Supporting Information & Attachments" stepIndex={5} onEdit={navigateTo} rows={[
                    { label: 'Any other information', value: form.anyOtherInfoYesNo },
                    ...(form.anyOtherInfoYesNo === 'Yes' ? [{ label: 'Other information details', value: form.anyOtherInfo }] : []),
                    { label: 'Attached files', value: attachedFiles.length ? attachedFiles.map(f => f.name).join(', ') : 'None' },
                  ]} />

                  <div className="govuk-warning-text" style={{ marginTop: '20px' }}>
                    <span className="govuk-warning-text__icon" aria-hidden="true">!</span>
                    <strong className="govuk-warning-text__text">
                      <span className="govuk-visually-hidden">Warning</span>
                      By submitting this report you confirm the information is accurate to the best of your knowledge.
                    </strong>
                  </div>
                </div>
              )}

              {/* Navigation buttons */}
              <div style={{ marginTop: '30px', paddingTop: '20px', borderTop: '1px solid #b1b4b6' }}>
                {isReview ? (
                  <div>
                    <div style={{ backgroundColor: '#fff3cd', border: '3px solid #ffb700', borderRadius: '4px', padding: '15px 20px', marginBottom: '20px' }}>
                      <p className="govuk-body govuk-!-font-weight-bold" style={{ marginBottom: '4px' }}>Email notifications will be sent to:</p>
                      <ul className="govuk-list govuk-list--bullet" style={{ marginBottom: 0 }}>
                        {(() => {
                          const recipients = [form.responsibleEmail, form.lineManagerEmail];
                          if (form.officerEmail && form.officerEmail !== form.responsibleEmail) recipients.push(form.officerEmail);
                          return recipients.filter(Boolean).map(email => <li key={email}>{email}</li>);
                        })()}
                      </ul>
                    </div>
                    {submitError && (
                      <p className="govuk-error-message" style={{ marginBottom: '15px' }}>
                        <span className="govuk-visually-hidden">Error:</span> {submitError}
                      </p>
                    )}
                    <button
                      type="button"
                      className={`govuk-button${isSubmitting ? ' govuk-button--disabled' : ''}`}
                      disabled={isSubmitting}
                      aria-disabled={isSubmitting}
                      onClick={handleSubmit}
                    >
                      {isSubmitting ? 'Submitting…' : 'Accept and submit'}
                    </button>
                  </div>
                ) : (
                  <button type="button" className="govuk-button" onClick={next}>
                    {step === REVIEW_STEP - 1 ? 'Check your answers' : 'Continue'}
                  </button>
                )}
              </div>

            </div>
          </div>
        </main>
      </div>
    </div>
  );
};

export default DataBreachForm;
