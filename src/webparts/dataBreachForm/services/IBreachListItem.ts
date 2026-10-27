// SharePoint column internal names taken directly from the list schema export.
// Keys on this interface ARE the column internal names used by PnPjs.

export interface IBreachListItem {
  // ── Overview ────────────────────────────────────────────────────────────────
  Title: string;
  Reported_x0020_to_x0020_ICO: string;

  // ── Reporting officer ────────────────────────────────────────────────────────
  Details_x0020_of_x0020_the_x0020: string;   // "Details of the officer reporting the breach" (officer name)
  Email_x0020_of_x0020_the_x0020_r: string;   // "Email of the reporting officer"
  Directorate_x0020_of_x0020_repor: string;   // "Directorate of reporting officer"
  Division_x0020_of_x0020_reportin: string;   // "Division of reporting officer"

  // ── Person responsible ───────────────────────────────────────────────────────
  Name_x0020_of_x0020_person_x0020: string;   // "Name of person responsible"
  Email_x0020_address_x0020_of_x00: string;   // "Email address of person responsible"
  Has_x0020_person_x0020_responsib: string;   // "Has person responsible completed mandatory training"
  Directorate_x0020_of_x0020_perso: string;   // "Directorate of person responsible"
  Division_x0020_of_x0020_person_x: string;   // "Division of person responsible"
  Line_x0020_manager_x0020_of_x002: string;   // "Line manager of person responsible"
  Email_x0020_address_x0020_of_x000: string;  // "Email address of line manager" (note trailing 0)

  // ── Breach details ───────────────────────────────────────────────────────────
  Details_x0020_of_x0020_the_x00200: string;  // "Details of the breach" (Note; trailing 0 distinguishes from officer field)
  Outline_x0020_of_x0020_data_x002: string;   // "Outline of data breached"
  Approx_x0020_individuals_x0020_a: string;   // "Approx individuals affected"
  Categorise_x0020_nature_x0020_of: string;   // "Categorise nature of breach"
  Categorise_x0020_nature_x0020_ot: string;   // "Categorise nature other" — new column
  Categories_x0020_of_x0020_person: string;   // "Categories of personal data in breach"
  Personal_x0020_data_x0020_other:  string;   // "Personal data other" — new column
  Categories_x0020_of_x0020_indivi: string;   // "Categories of individuals in breach"
  Individuals_x0020_other:          string;   // "Individuals other" — new column
  Impact_x0020_to_x0020_Data_x0020: string;   // "Impact to Data Subjects"
  Breach_x0020_categories: string;            // "Breach categories" — semicolon-separated multi-select
  Date_x0020_breach_x0020_discover: string | null;   // "Date breach discovered" (ISO date)
  Time_x0020_breach_x0020_discover: string;          // "Time breach discovered"
  Date_x0020_breach_x0020_happened: string | null;   // "Date breach happened" (ISO date)
  Time_x0020_breach_x0020_happened: string;          // "Time breach happened"

  // ── Third party ──────────────────────────────────────────────────────────────
  Name_x0020_of_x0020_third_x0020_: string;   // "Name of third party"
  Address_x0020_of_x0020_third_x00: string;   // "Address of third party"
  Contact_x0020_for_x0020_third_x0: string;   // "Contact for third party"
  Email_x0020_for_x0020_third_x002: string;   // "Email for third party"
  Telephone_x0020_for_x0020_third_: string;   // "Telephone for third party"

  // ── Consequences & actions ───────────────────────────────────────────────────
  Likely_x0020_consequences_x0020_: string;   // "Likely consequences of breach"
  Consequences_x0020_not_x0020_yet: string;   // "Consequences not yet known explanation"
  Steps_x0020_taken_x0020_for_x002: string;   // "Steps taken for resolution"
  Evidence_x0020_to_x0020_prevent_: string;   // "Evidence to prevent future incidents"

  // ── Notifications ────────────────────────────────────────────────────────────
  Have_x0020_affected_x0020_indivi: string;   // "Have affected individuals been informed"
  Details_x0020_of_x0020_informing: string;   // "Details of informing affected individuals"
  Does_x0020_breach_x0020_relate_x: string;   // "Does breach relate to agreement with third party"
  What_x0020_information_x0020_sha: string;   // "What information shared about breach"
  Have_x0020_other_x0020_organisat: string;   // "Have other organisations been informed"
  Details_x0020_of_x0020_informing0: string;  // "Details of informing other organisations" (trailing 0)
  Has_x0020_media_x0020_covered_x0: string;   // "Has media covered incident"
  Media_x0020_coverage_x0020_of_x0: string;   // "Media coverage of incident details"
  Any_x0020_other_x0020_info: string;         // "Any other info"

  // ── Panel & decision ─────────────────────────────────────────────────────────
  Panel_x0020_Recommendations: string;        // "Panel Recommendations"
  Notification_x0020_to_x0020_ICO_: string;   // "Notification to ICO (exists)"
  Staff_x0020_e_x002d_Learning_x00: string;   // "Staff e-Learning training required"
  Review_x0020_of_x0020_Procedures: string;   // "Review of Procedures/Processes required"
  Reason_x0020_for_x0020_Decision: string;    // "Reason for Decision"

  // ── Communication log ────────────────────────────────────────────────────────
  CommunicationLog?: string | null;            // "CommunicationLog" — JSON array of ICommLogEntry, append-only thread
}
