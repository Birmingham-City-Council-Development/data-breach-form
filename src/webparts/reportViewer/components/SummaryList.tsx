import * as React from 'react';
import { EditingState } from './reportViewerTypes';

export const Field: React.FC<{
  label: string;
  value: string;
  fieldKey: keyof EditingState;
  editing: boolean;
  multiline?: boolean;
  options?: string[];
  onChangeCallback?: (value: string) => void;
  editingState: EditingState;
  onChange: (key: keyof EditingState, value: string) => void;
  // Shown in place of the default "Not provided" when the field is blank —
  // for fields that are only ever blank because a branch of the form didn't
  // apply to this report (e.g. "Has personal data been placed at risk?" was
  // No, so nothing downstream of it was collected), "N/A" reads correctly;
  // "Not provided" reads correctly for fields that are always relevant but
  // were genuinely left empty.
  emptyText?: string;
}> = ({ label, value, fieldKey, editing, multiline, options, onChangeCallback, editingState, onChange, emptyText }) => {
  const displayValue = editing && editingState[fieldKey] !== undefined
    ? (editingState[fieldKey] as string)
    : value;

  const handleChange = (val: string): void => {
    onChange(fieldKey, val);
    if (onChangeCallback) onChangeCallback(val);
  };

  return (
    <div className="govuk-summary-list__row">
      <dt className="govuk-summary-list__key" style={{ width: '35%' }}>{label}</dt>
      <dd className="govuk-summary-list__value">
        {editing ? (
          options ? (
            <select
              className="govuk-select"
              value={displayValue || ''}
              onChange={e => handleChange(e.target.value)}
            >
              <option value="">— Select —</option>
              {options.map(o => <option key={o} value={o}>{o}</option>)}
            </select>
          ) : multiline ? (
            <textarea
              className="govuk-textarea"
              rows={3}
              value={displayValue || ''}
              style={{ marginBottom: 0 }}
              onChange={e => handleChange(e.target.value)}
            />
          ) : (
            <input
              className="govuk-input"
              type="text"
              value={displayValue || ''}
              onChange={e => handleChange(e.target.value)}
            />
          )
        ) : (
          <span style={{ color: displayValue ? '#0b0c0c' : '#505a5f' }}>
            {displayValue || emptyText || 'Not provided'}
          </span>
        )}
      </dd>
    </div>
  );
};

// Note: fields become editable whenever editingSection matches sectionKey,
// independent of canEdit — canEdit only hides the manual "Edit" link. Any
// programmatic startEdit(...) call must respect the section's role rules itself.
export const Section: React.FC<{
  title: string;
  children: React.ReactNode;
  sectionKey: string;
  editingSection: string | null;
  saving: boolean;
  canEdit: boolean;
  onEdit: (key: string) => void;
  onSave: () => void;
  onCancel: () => void;
  adminOnly?: boolean;
  restrictedLabel?: string;
}> = ({ title, children, sectionKey, editingSection, saving, canEdit, onEdit, onSave, onCancel, adminOnly, restrictedLabel }) => {
  const isEditing = editingSection === sectionKey;
  return (
    <div style={{ marginBottom: '30px', ...(adminOnly ? { background: '#f3f2f1', padding: '16px', borderRadius: '4px' } : {}) }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', borderBottom: '2px solid #1d70b8', paddingBottom: '8px' }}>
        <h2 className="govuk-heading-s" style={{ margin: 0 }}>
          {title}
          {adminOnly && <strong className="govuk-tag" style={{ marginLeft: '10px', verticalAlign: 'middle' }}>{restrictedLabel || 'Admin only'}</strong>}
        </h2>
        {canEdit && !isEditing && !editingSection && (
          <a className="govuk-link govuk-body-s no-print" href="#" onClick={e => { e.preventDefault(); onEdit(sectionKey); }}>Edit</a>
        )}
        {isEditing && (
          <div className="no-print" style={{ display: 'flex', gap: '10px' }}>
            <button
              type="button"
              className={`govuk-button govuk-button--small${saving ? ' govuk-button--disabled' : ''}`}
              disabled={saving}
              onClick={onSave}
            >
              {saving ? 'Saving…' : 'Save'}
            </button>
            <button type="button" className="govuk-button govuk-button--secondary govuk-button--small" onClick={onCancel}>
              Cancel
            </button>
          </div>
        )}
      </div>
      <dl className="govuk-summary-list" style={{ marginBottom: 0 }}>{children}</dl>
    </div>
  );
};
