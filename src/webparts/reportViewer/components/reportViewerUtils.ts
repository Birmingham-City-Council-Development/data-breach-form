import { IReportSection } from '../services/ReportPdf';
import { REVIEW_TEAM_NAME } from '../../dataBreachForm/config';
import { FullItem } from './reportViewerTypes';

export function formatDate(iso: string): string {
  if (!iso) return '';
  return new Date(iso).toLocaleDateString('en-GB');
}

export function formatDateTime(iso: string): string {
  if (!iso) return '';
  return new Date(iso).toLocaleString('en-GB', {
    day: '2-digit', month: '2-digit', year: 'numeric',
    hour: '2-digit', minute: '2-digit',
  });
}

// The Power Automate flow that used to generate Data_x0020_Breach_x0020_ID
// has been switched off — it's no longer written for new reports, and old
// values in that column shouldn't be trusted going forward either. The
// reference is now derived purely from the list's own auto-increment Id,
// matching the format DataBreachForm.tsx already uses for the post-submit
// confirmation (`DBN-${id}`), so this doesn't need any SharePoint flow to
// keep working.
export function getRef(id: number): string {
  return `DBN-${id}`;
}

export const YES_NO = ['Yes', 'No'];

export const STATUS_ORDER: { [key: string]: number } = { 'Open': 0, 'Pending': 1, 'Closed': 2 };

export const LIST_PAGE_SIZE = 25;

// Labelling/colour for StatusChange comm-log entries, keyed by the resulting
// status (carried in entry.subject) — so closing, reopening, and pending
// each read distinctly instead of all sharing one generic "Status change"
// tag that only differs in the message text underneath.
const STATUS_CHANGE_STYLE: { [status: string]: { label: string; tagClass: string; borderColor: string } } = {
  Closed: { label: 'Case Closed', tagClass: 'govuk-tag--red', borderColor: '#d4351c' },
  Open: { label: 'Case Reopened', tagClass: 'govuk-tag--blue', borderColor: '#1d70b8' },
  Pending: { label: 'Marked Pending', tagClass: 'govuk-tag--yellow', borderColor: '#ffb700' },
};

export function getStatusChangeStyle(subject: string): { label: string; tagClass: string; borderColor: string } {
  return STATUS_CHANGE_STYLE[subject] || { label: 'Status change', tagClass: 'govuk-tag--grey', borderColor: '#505a5f' };
}

// Pre-filled into "Reason for decision" when an admin starts editing Panel
// recommendations and the field is still blank — this is the outcome the
// data breach notification team gives most of the time, so it saves retyping it every
// time. Still fully editable before saving; only ever seeds an empty field,
// never overwrites an existing decision.
export const DEFAULT_REASON_FOR_DECISION = [
  `${REVIEW_TEAM_NAME} considered the circumstances of the breach and the outcome is that the risk to the individuals and the level of severity would be minimal and therefore does not meet the criteria for notification to the ICO.`,
  'The incident was found to be a breach with no material effect based on the severity of the potential or actual impact on the individuals. It is very unlikely that there would be any risk of harm or detriment to the data subject(s) that would lead to physical, material or non-material damage for the individuals whose data has been breached.',
  'As the breach was unlikely to result in a high risk to the rights and freedoms of individuals, the UK GDPR would not require you to inform those concerned.',
].join('\n\n');

// Shared by the CSV export and the PDF attached to case-closed emails, so
// both "full report dump" outputs stay in sync with a single field list.
export function getReportFieldRows(item: FullItem): Array<[string, string]> {
  return [
    ['Reference',                        getRef(item.Id)],
    ['Status',                           item.Status || 'Open'],
    ['Submitted by',                     item.Author ? item.Author.Title : ''],
    ['Date submitted',                   formatDate(item.Created)],
    ['Reported to ICO',                  item.Reported_x0020_to_x0020_ICO || ''],
    ['Reporting officer name',           item.Details_x0020_of_x0020_the_x0020 || ''],
    ['Reporting officer email',          item.Email_x0020_of_x0020_the_x0020_r || ''],
    ['Reporting officer directorate',    item.Directorate_x0020_of_x0020_repor || ''],
    ['Reporting officer division',       item.Division_x0020_of_x0020_reportin || ''],
    ['Person responsible name',          item.Name_x0020_of_x0020_person_x0020 || ''],
    ['Person responsible email',         item.Email_x0020_address_x0020_of_x00 || ''],
    ['Mandatory training completed',     item.Has_x0020_person_x0020_responsib || ''],
    ['Person responsible directorate',   item.Directorate_x0020_of_x0020_perso || ''],
    ['Person responsible division',      item.Division_x0020_of_x0020_person_x || ''],
    ['Line manager',                     item.Line_x0020_manager_x0020_of_x002 || ''],
    ['Line manager email',               item.Email_x0020_address_x0020_of_x000 || ''],
    ['Details of breach',                item.Details_x0020_of_x0020_the_x00200 || ''],
    ['Outline of data breached',         item.Outline_x0020_of_x0020_data_x002 || ''],
    ['Approx individuals affected',      item.Approx_x0020_individuals_x0020_a || ''],
    ['Nature of breach',                 item.Categorise_x0020_nature_x0020_of || ''],
    ['Nature of breach — other',         item.Categorise_x0020_nature_x0020_ot || ''],
    ['Categories of personal data',      item.Categories_x0020_of_x0020_person || ''],
    ['Personal data categories — other', item.Personal_x0020_data_x0020_other || ''],
    ['Categories of individuals',        item.Categories_x0020_of_x0020_indivi || ''],
    ['Individuals categories — other',   item.Individuals_x0020_other || ''],
    ['Impact to data subjects',          item.Impact_x0020_to_x0020_Data_x0020 || ''],
    ['How breach occurred',              item.Breach_x0020_categories || ''],
    ['Date discovered',                  item.Date_x0020_breach_x0020_discover ? formatDate(item.Date_x0020_breach_x0020_discover) : ''],
    ['Time discovered',                  item.Time_x0020_breach_x0020_discover || ''],
    ['Date happened',                    item.Date_x0020_breach_x0020_happened ? formatDate(item.Date_x0020_breach_x0020_happened) : ''],
    ['Time happened',                    item.Time_x0020_breach_x0020_happened || ''],
    ['Third party name',                 item.Name_x0020_of_x0020_third_x0020_ || ''],
    ['Third party address',              item.Address_x0020_of_x0020_third_x00 || ''],
    ['Third party contact',              item.Contact_x0020_for_x0020_third_x0 || ''],
    ['Third party email',                item.Email_x0020_for_x0020_third_x002 || ''],
    ['Third party telephone',            item.Telephone_x0020_for_x0020_third_ || ''],
    ['Likely consequences',              item.Likely_x0020_consequences_x0020_ || ''],
    ['Consequences explanation',         item.Consequences_x0020_not_x0020_yet || ''],
    ['Steps taken',                      item.Steps_x0020_taken_x0020_for_x002 || ''],
    ['Evidence to prevent recurrence',   item.Evidence_x0020_to_x0020_prevent_ || ''],
    ['Individuals informed',             item.Have_x0020_affected_x0020_indivi || ''],
    ['Details of informing individuals', item.Details_x0020_of_x0020_informing || ''],
    ['Third party agreement',            item.Does_x0020_breach_x0020_relate_x || ''],
    ['Information shared',               item.What_x0020_information_x0020_sha || ''],
    ['Other organisations informed',     item.Have_x0020_other_x0020_organisat || ''],
    ['Details of other organisations',   item.Details_x0020_of_x0020_informing0 || ''],
    ['Media covered',                    item.Has_x0020_media_x0020_covered_x0 || ''],
    ['Media details',                    item.Media_x0020_coverage_x0020_of_x0 || ''],
    ['Any other info',                   item.Any_x0020_other_x0020_info || ''],
    ['Panel recommendations',            item.Panel_x0020_Recommendations || ''],
    ['Staff e-Learning training',        item.Staff_x0020_e_x002d_Learning_x00 || ''],
    ['Procedures review required',       item.Review_x0020_of_x0020_Procedures || ''],
    ['Reason for decision',              item.Reason_x0020_for_x0020_Decision || ''],
  ];
}

// For the case-closed PDF only (buildReportPdfBase64) — grouped to mirror
// exactly what the <Section> blocks in SubmittedReport.tsx and the Panel
// recommendations section show on screen, field-for-field, so the PDF reads
// like the report page rather than a flat dump. Deliberately separate from
// getReportFieldRows (used by the CSV export): that flat list includes a
// couple of fields — notably Panel_x0020_Recommendations — that were part of
// the original CSV export but were never surfaced in the on-screen Panel
// recommendations Section (which shows Reported to ICO / e-Learning /
// Procedures / Reason for decision instead), so reusing it here would show a
// field on the PDF that nobody sees when actually viewing the report.
export function getReportFieldSections(item: FullItem): IReportSection[] {
  return [
    { title: 'Reporting Officer', rows: [
      ['Name', item.Details_x0020_of_x0020_the_x0020 || ''],
      ['Email', item.Email_x0020_of_x0020_the_x0020_r || ''],
      ['Directorate', item.Directorate_x0020_of_x0020_repor || ''],
      ['Division', item.Division_x0020_of_x0020_reportin || ''],
    ] },
    { title: 'Person Responsible', rows: [
      ['Name', item.Name_x0020_of_x0020_person_x0020 || ''],
      ['Email', item.Email_x0020_address_x0020_of_x00 || ''],
      ['Mandatory training completed', item.Has_x0020_person_x0020_responsib || ''],
      ['Directorate', item.Directorate_x0020_of_x0020_perso || ''],
      ['Division', item.Division_x0020_of_x0020_person_x || ''],
      ['Line manager', item.Line_x0020_manager_x0020_of_x002 || ''],
      ['Line manager email', item.Email_x0020_address_x0020_of_x000 || ''],
    ] },
    { title: 'Breach Details', rows: [
      ['Details of the breach', item.Details_x0020_of_x0020_the_x00200 || ''],
      ['Outline of data breached', item.Outline_x0020_of_x0020_data_x002 || ''],
      ['Approx individuals affected', item.Approx_x0020_individuals_x0020_a || ''],
      ['Nature of breach', item.Categorise_x0020_nature_x0020_of || ''],
      ['Nature of breach — other', item.Categorise_x0020_nature_x0020_ot || ''],
      ['Categories of personal data', item.Categories_x0020_of_x0020_person || ''],
      ['Personal data categories — other', item.Personal_x0020_data_x0020_other || ''],
      ['Categories of individuals', item.Categories_x0020_of_x0020_indivi || ''],
      ['Individuals categories — other', item.Individuals_x0020_other || ''],
      ['Impact to data subjects', item.Impact_x0020_to_x0020_Data_x0020 || ''],
      ['How did the breach occur', item.Breach_x0020_categories || ''],
      ['Date discovered', item.Date_x0020_breach_x0020_discover ? formatDate(item.Date_x0020_breach_x0020_discover) : ''],
      ['Time discovered', item.Time_x0020_breach_x0020_discover || ''],
      ['Date happened', item.Date_x0020_breach_x0020_happened ? formatDate(item.Date_x0020_breach_x0020_happened) : ''],
      ['Time happened', item.Time_x0020_breach_x0020_happened || ''],
    ] },
    { title: 'Third Party', rows: [
      ['Name', item.Name_x0020_of_x0020_third_x0020_ || ''],
      ['Address', item.Address_x0020_of_x0020_third_x00 || ''],
      ['Contact', item.Contact_x0020_for_x0020_third_x0 || ''],
      ['Email', item.Email_x0020_for_x0020_third_x002 || ''],
      ['Telephone', item.Telephone_x0020_for_x0020_third_ || ''],
    ] },
    { title: 'Consequences & Actions', rows: [
      ['Likely consequences', item.Likely_x0020_consequences_x0020_ || ''],
      ['Consequences explanation', item.Consequences_x0020_not_x0020_yet || ''],
      ['Steps taken', item.Steps_x0020_taken_x0020_for_x002 || ''],
      ['Evidence to prevent recurrence', item.Evidence_x0020_to_x0020_prevent_ || ''],
    ] },
    { title: 'Notifications', rows: [
      ['Individuals informed', item.Have_x0020_affected_x0020_indivi || ''],
      ['Details of informing individuals', item.Details_x0020_of_x0020_informing || ''],
      ['Third party agreement', item.Does_x0020_breach_x0020_relate_x || ''],
      ['Information shared', item.What_x0020_information_x0020_sha || ''],
      ['Other organisations informed', item.Have_x0020_other_x0020_organisat || ''],
      ['Details of other orgs', item.Details_x0020_of_x0020_informing0 || ''],
      ['Media covered', item.Has_x0020_media_x0020_covered_x0 || ''],
      ['Media details', item.Media_x0020_coverage_x0020_of_x0 || ''],
      ['Any other info', item.Any_x0020_other_x0020_info || ''],
    ] },
    { title: 'Panel Recommendations', rows: [
      ['Reported to ICO', item.Reported_x0020_to_x0020_ICO || ''],
      ['Staff e-Learning training', item.Staff_x0020_e_x002d_Learning_x00 || ''],
      ['Procedures review required', item.Review_x0020_of_x0020_Procedures || ''],
      ['Reason for decision', item.Reason_x0020_for_x0020_Decision || ''],
    ] },
  ];
}
