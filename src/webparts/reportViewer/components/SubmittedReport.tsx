import * as React from 'react';
import { IAttachmentInfo } from '../../dataBreachForm/services/SharePointService';
import { INDIVIDUALS_INFORMED_OPTIONS, DIRECTORATES, getDivisions } from '../../dataBreachForm/components/formOptions';
import { EditingState, FieldReader } from './reportViewerTypes';
import { formatDate, YES_NO } from './reportViewerUtils';
import { Field, Section } from './SummaryList';

// The report as submitted by the reporting officer.
const SubmittedReport: React.FC<{
  f: FieldReader;
  attachments: IAttachmentInfo[];
  editingSection: string | null;
  editingState: EditingState;
  saving: boolean;
  onChange: (key: keyof EditingState, value: string) => void;
  onEdit: (key: string) => void;
  onSave: () => void;
  onCancel: () => void;
}> = ({ f, attachments, editingSection, editingState, saving, onChange, onEdit, onSave, onCancel }) => {
  const sectionProps = { editingSection, saving, onEdit, onSave, onCancel };
  const fieldProps = { editingState, onChange };
  return (
    <>
      <div style={{ borderTop: '3px solid #b1b4b6', paddingTop: '20px', marginBottom: '24px' }}>
        <p className="govuk-body-s" style={{ color: '#505a5f', marginBottom: 0 }}>
          <strong>Submitted information</strong> — provided by the reporting officer
        </p>
      </div>

      {/* Editing disabled for report integrity — the submitted report must
          not be alterable by anyone, admins included. canEdit is hardcoded
          false rather than removing the edit machinery, so this can be
          reversed by flipping the flag if that policy ever changes. */}
      <Section title="Reporting Officer" sectionKey="officer" canEdit={false} {...sectionProps}>
        <Field label="Name" value={f('Details_x0020_of_x0020_the_x0020')} fieldKey="Details_x0020_of_x0020_the_x0020" editing={editingSection === 'officer'} {...fieldProps} />
        <Field label="Email" value={f('Email_x0020_of_x0020_the_x0020_r')} fieldKey="Email_x0020_of_x0020_the_x0020_r" editing={editingSection === 'officer'} {...fieldProps} />
        <Field label="Directorate" value={f('Directorate_x0020_of_x0020_repor')} fieldKey="Directorate_x0020_of_x0020_repor" editing={editingSection === 'officer'} options={DIRECTORATES} onChangeCallback={() => onChange('Division_x0020_of_x0020_reportin', '')} {...fieldProps} />
        <Field label="Division" value={f('Division_x0020_of_x0020_reportin')} fieldKey="Division_x0020_of_x0020_reportin" editing={editingSection === 'officer'} options={getDivisions((editingState['Directorate_x0020_of_x0020_repor'] as string) || f('Directorate_x0020_of_x0020_repor'))} {...fieldProps} />
      </Section>

      <Section title="Person Responsible" sectionKey="responsible" canEdit={false} {...sectionProps}>
        <Field label="Name" value={f('Name_x0020_of_x0020_person_x0020')} fieldKey="Name_x0020_of_x0020_person_x0020" editing={editingSection === 'responsible'} {...fieldProps} />
        <Field label="Email" value={f('Email_x0020_address_x0020_of_x00')} fieldKey="Email_x0020_address_x0020_of_x00" editing={editingSection === 'responsible'} {...fieldProps} />
        <Field label="Mandatory training completed" value={f('Has_x0020_person_x0020_responsib')} fieldKey="Has_x0020_person_x0020_responsib" editing={editingSection === 'responsible'} options={YES_NO} {...fieldProps} />
        <Field label="Directorate" value={f('Directorate_x0020_of_x0020_perso')} fieldKey="Directorate_x0020_of_x0020_perso" editing={editingSection === 'responsible'} options={DIRECTORATES} onChangeCallback={() => onChange('Division_x0020_of_x0020_person_x', '')} {...fieldProps} />
        <Field label="Division" value={f('Division_x0020_of_x0020_person_x')} fieldKey="Division_x0020_of_x0020_person_x" editing={editingSection === 'responsible'} options={getDivisions((editingState['Directorate_x0020_of_x0020_perso'] as string) || f('Directorate_x0020_of_x0020_perso'))} {...fieldProps} />
        <Field label="Line manager" value={f('Line_x0020_manager_x0020_of_x002')} fieldKey="Line_x0020_manager_x0020_of_x002" editing={editingSection === 'responsible'} {...fieldProps} />
        <Field label="Line manager email" value={f('Email_x0020_address_x0020_of_x000')} fieldKey="Email_x0020_address_x0020_of_x000" editing={editingSection === 'responsible'} {...fieldProps} />
      </Section>

      <Section title="Breach Details" sectionKey="breach" canEdit={false} {...sectionProps}>
        <Field label="Details of the breach" value={f('Details_x0020_of_x0020_the_x00200')} fieldKey="Details_x0020_of_x0020_the_x00200" editing={editingSection === 'breach'} multiline {...fieldProps} />
        {/* "Has any personal data been placed at risk?" itself isn't a stored
            SharePoint field (never was mapped at submission — see
            mapFormToListItem) — inferred here instead of shown as a gap,
            since dataBreachedOutline is enforced as mandatory on the form
            whenever that answer is Yes, making its presence a reliable proxy.
            Grouped with the fields below since they're all gated behind that
            same Yes/No branch on the submission form and are otherwise
            confusingly blank with no explanation when the answer was No. */}
        <div className="govuk-summary-list__row">
          <dt className="govuk-summary-list__key" style={{ width: '35%' }}>Has any personal data been placed at risk?</dt>
          <dd className="govuk-summary-list__value">
            <span style={{ color: '#0b0c0c' }}>{f('Outline_x0020_of_x0020_data_x002') ? 'Yes' : 'No'}</span>
          </dd>
        </div>
        <Field label="Outline of data breached" value={f('Outline_x0020_of_x0020_data_x002')} fieldKey="Outline_x0020_of_x0020_data_x002" editing={editingSection === 'breach'} multiline {...fieldProps} emptyText="N/A" />
        <Field label="Categories of personal data" value={f('Categories_x0020_of_x0020_person')} fieldKey="Categories_x0020_of_x0020_person" editing={editingSection === 'breach'} {...fieldProps} emptyText="N/A" />
        <Field label="Personal data categories — other" value={f('Personal_x0020_data_x0020_other')} fieldKey="Personal_x0020_data_x0020_other" editing={editingSection === 'breach'} {...fieldProps} emptyText="N/A" />
        <Field label="Categories of individuals" value={f('Categories_x0020_of_x0020_indivi')} fieldKey="Categories_x0020_of_x0020_indivi" editing={editingSection === 'breach'} {...fieldProps} emptyText="N/A" />
        <Field label="Individuals categories — other" value={f('Individuals_x0020_other')} fieldKey="Individuals_x0020_other" editing={editingSection === 'breach'} {...fieldProps} emptyText="N/A" />
        <Field label="Impact to data subjects" value={f('Impact_x0020_to_x0020_Data_x0020')} fieldKey="Impact_x0020_to_x0020_Data_x0020" editing={editingSection === 'breach'} {...fieldProps} emptyText="N/A" />
        <Field label="Approx individuals affected" value={f('Approx_x0020_individuals_x0020_a')} fieldKey="Approx_x0020_individuals_x0020_a" editing={editingSection === 'breach'} {...fieldProps} />
        <Field label="Nature of breach" value={f('Categorise_x0020_nature_x0020_of')} fieldKey="Categorise_x0020_nature_x0020_of" editing={editingSection === 'breach'} {...fieldProps} />
        <Field label="Nature of breach — other" value={f('Categorise_x0020_nature_x0020_ot')} fieldKey="Categorise_x0020_nature_x0020_ot" editing={editingSection === 'breach'} {...fieldProps} />
        <Field label="How did the breach occur" value={f('Breach_x0020_categories')} fieldKey="Breach_x0020_categories" editing={editingSection === 'breach'} {...fieldProps} />
        <Field label="Date discovered" value={formatDate(f('Date_x0020_breach_x0020_discover'))} fieldKey="Date_x0020_breach_x0020_discover" editing={editingSection === 'breach'} {...fieldProps} />
        <Field label="Time discovered" value={f('Time_x0020_breach_x0020_discover')} fieldKey="Time_x0020_breach_x0020_discover" editing={editingSection === 'breach'} {...fieldProps} />
        <Field label="Date happened" value={formatDate(f('Date_x0020_breach_x0020_happened'))} fieldKey="Date_x0020_breach_x0020_happened" editing={editingSection === 'breach'} {...fieldProps} />
        <Field label="Time happened" value={f('Time_x0020_breach_x0020_happened')} fieldKey="Time_x0020_breach_x0020_happened" editing={editingSection === 'breach'} {...fieldProps} />
      </Section>

      <Section title="Third Party" sectionKey="thirdparty" canEdit={false} {...sectionProps}>
        <Field label="Name" value={f('Name_x0020_of_x0020_third_x0020_')} fieldKey="Name_x0020_of_x0020_third_x0020_" editing={editingSection === 'thirdparty'} {...fieldProps} />
        <Field label="Address" value={f('Address_x0020_of_x0020_third_x00')} fieldKey="Address_x0020_of_x0020_third_x00" editing={editingSection === 'thirdparty'} multiline {...fieldProps} />
        <Field label="Contact" value={f('Contact_x0020_for_x0020_third_x0')} fieldKey="Contact_x0020_for_x0020_third_x0" editing={editingSection === 'thirdparty'} {...fieldProps} />
        <Field label="Email" value={f('Email_x0020_for_x0020_third_x002')} fieldKey="Email_x0020_for_x0020_third_x002" editing={editingSection === 'thirdparty'} {...fieldProps} />
        <Field label="Telephone" value={f('Telephone_x0020_for_x0020_third_')} fieldKey="Telephone_x0020_for_x0020_third_" editing={editingSection === 'thirdparty'} {...fieldProps} />
      </Section>

      <Section title="Consequences & Actions" sectionKey="consequences" canEdit={false} {...sectionProps}>
        <Field label="Likely consequences" value={f('Likely_x0020_consequences_x0020_')} fieldKey="Likely_x0020_consequences_x0020_" editing={editingSection === 'consequences'} multiline {...fieldProps} />
        <Field label="Consequences explanation" value={f('Consequences_x0020_not_x0020_yet')} fieldKey="Consequences_x0020_not_x0020_yet" editing={editingSection === 'consequences'} multiline {...fieldProps} />
        <Field label="Steps taken" value={f('Steps_x0020_taken_x0020_for_x002')} fieldKey="Steps_x0020_taken_x0020_for_x002" editing={editingSection === 'consequences'} multiline {...fieldProps} />
        <Field label="Evidence to prevent recurrence" value={f('Evidence_x0020_to_x0020_prevent_')} fieldKey="Evidence_x0020_to_x0020_prevent_" editing={editingSection === 'consequences'} multiline {...fieldProps} />
      </Section>

      <Section title="Notifications" sectionKey="notifications" canEdit={false} {...sectionProps}>
        <Field label="Individuals informed" value={f('Have_x0020_affected_x0020_indivi')} fieldKey="Have_x0020_affected_x0020_indivi" editing={editingSection === 'notifications'} options={INDIVIDUALS_INFORMED_OPTIONS} {...fieldProps} />
        <Field label="Details of informing individuals" value={f('Details_x0020_of_x0020_informing')} fieldKey="Details_x0020_of_x0020_informing" editing={editingSection === 'notifications'} multiline {...fieldProps} />
        <Field label="Third party agreement" value={f('Does_x0020_breach_x0020_relate_x')} fieldKey="Does_x0020_breach_x0020_relate_x" editing={editingSection === 'notifications'} options={YES_NO} {...fieldProps} />
        <Field label="Information shared" value={f('What_x0020_information_x0020_sha')} fieldKey="What_x0020_information_x0020_sha" editing={editingSection === 'notifications'} multiline {...fieldProps} />
        <Field label="Other organisations informed" value={f('Have_x0020_other_x0020_organisat')} fieldKey="Have_x0020_other_x0020_organisat" editing={editingSection === 'notifications'} options={YES_NO} {...fieldProps} />
        <Field label="Details of other orgs" value={f('Details_x0020_of_x0020_informing0')} fieldKey="Details_x0020_of_x0020_informing0" editing={editingSection === 'notifications'} multiline {...fieldProps} />
        <Field label="Media covered" value={f('Has_x0020_media_x0020_covered_x0')} fieldKey="Has_x0020_media_x0020_covered_x0" editing={editingSection === 'notifications'} options={YES_NO} {...fieldProps} />
        <Field label="Media details" value={f('Media_x0020_coverage_x0020_of_x0')} fieldKey="Media_x0020_coverage_x0020_of_x0" editing={editingSection === 'notifications'} multiline {...fieldProps} />
        <Field label="Any other info" value={f('Any_x0020_other_x0020_info')} fieldKey="Any_x0020_other_x0020_info" editing={editingSection === 'notifications'} multiline {...fieldProps} />
      </Section>

      <Section title="Attachments" sectionKey="attachments" canEdit={false} {...sectionProps}>
        <div className="govuk-summary-list__row">
          <dt className="govuk-summary-list__key" style={{ width: '35%' }}>Files</dt>
          <dd className="govuk-summary-list__value">
            {attachments.length === 0 ? (
              <span style={{ color: '#505a5f' }}>None</span>
            ) : (
              <ul className="govuk-list" style={{ marginBottom: 0 }}>
                {attachments.map(a => (
                  <li key={a.FileName}>
                    <a className="govuk-link" href={a.ServerRelativeUrl} target="_blank" rel="noopener noreferrer">{a.FileName}</a>
                  </li>
                ))}
              </ul>
            )}
          </dd>
        </div>
      </Section>
    </>
  );
};

export default SubmittedReport;
