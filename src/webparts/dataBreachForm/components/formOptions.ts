// ─────────────────────────────────────────────────────────────────────────────
// Central options file — edit this to update any dropdown in the form
// ─────────────────────────────────────────────────────────────────────────────

export const DIRECTORATES_WITH_DIVISIONS: Record<string, string[]> = {
  'Adult Social Care & Health': [
    'Public Health',
    'Adult Social Care',
    'Commissioning',
    'Early Intervention & Prevention & Libraries',
  ],
  'Children & Families': [
    'Schools & Employability',
    'Early years SEND & Inclusion',
    'Commissioning & Performance',
    'Birmingham Childrens Trust',
    'Thriving Children & Families',
  ],
  'City Housing': [
    'Housing Solutions & Support, Inc. Homelessness',
    'Asset Management',
    'Housing Management',
    'Regulation & Strategy, inc. Affordable Housing',
    'Druids Heath & Ladywood',
    'Regeneration Delivery',
  ],
  'City Operations': [
    'Emergency Planning & Operational Control Centre',
    'Regulation & Enforcement',
    'Waste & Street Scene',
    'Highways & Infrastructure',
    'Neighbourhoods',
    'Community Safety & Cohesion',
  ],
  'Finance S151': [
    'Finance functions',
    'Oracle SRO',
    'Procurement',
    'Companies & Commercial',
    'Revenues, Benefits & Rents',
    'Audit, Risk & Assurance',
    'IT & Digital',
  ],
  'HR, OD and Transformation': [
    'Payroll',
    'OD, Culture & Organisation',
    'Transformation',
    'HR Operations & Policy',
    'Equal Pay SRO',
    'Employee & Industrial Relations',
    'Contact Centre / Shared Services',
    'Business Support / Digital Mail',
  ],
  'Place, Prosperity & Sustainability': [
    'Economic Development',
    'Corporate Landlord',
    'Planning',
    'Regeneration',
    'Strategy & Delivery',
    'Asset Sales',
    'Major Projects & CA',
    'Link',
    'Employment & Skills',
  ],
  'Strategy, Insight & Governance': [
    'Deputy CE',
    'Intervention',
    'Response',
    'Comms',
    'CPMO',
    'Day to Day Political Interface',
    'Strategic Partnerships',
    'Policy, Performance & Insights',
    'Law & Governance / The Monitoring Officer',
  ],
};

export const DIRECTORATES: string[] = Object.keys(DIRECTORATES_WITH_DIVISIONS);

export function getDivisions(directorate: string): string[] {
  return DIRECTORATES_WITH_DIVISIONS[directorate] || [];
}

// ── Breach details ────────────────────────────────────────────────────────────

export const APPROX_INDIVIDUALS: string[] = [
  '1 to 9',
  '10 to 49',
  '50 to 99',
  '100 to 499',
  '500 to 999',
  '1,000 to 4,999',
  '5,000 to 9,999',
  '10,000 or more',
];

export const NATURE_OF_BREACH: string[] = [
  'High volume of personal data',
  'Sensitive personal data',
  'Data made public',
  'Financial data',
  'Other',
];

export const CATEGORIES_PERSONAL_DATA: string[] = [
  'Political opinions',
  'Health data',
  'Financial data',
  'Criminal records',
  'Racial or ethnic origin',
  'Religious beliefs',
  'Trade union membership',
  'Biometric data',
  'Genetic data',
  'Sex life or sexual orientation',
  'Other',
];

export const CATEGORIES_INDIVIDUALS: string[] = [
  'Patient',
  'Employee',
  'Member of public',
  'Child',
  'Vulnerable adult',
  'Other',
];

export const IMPACT_TO_SUBJECTS: string[] = [
  '1 - Low',
  '2 - Low-Medium',
  '3 - Medium',
  '4 - High',
  '5 - Critical',
];

export const LIKELIHOOD_OF_CONSEQUENCES: string[] = [
  '1 - Very unlikely',
  '2 - Unlikely',
  '3 - Possible',
  '4 - Likely',
  '5 - Very likely',
];

// ── Breach mechanism ──────────────────────────────────────────────────────────

export const BREACH_CATEGORIES: string[] = [
  'Email Disclosure – Personal data sent to the wrong recipient via email.',
  'System Misconfiguration – Incorrect system settings exposing data (e.g., shared drives).',
  'Cybersecurity Incident – Hacking, phishing, malware, or ransomware.',
  'Lost/Stolen Device – Laptops, tablets, or mobile phones containing sensitive data.',
  'Lost Memory Stick – USB or external storage device lost or stolen.',
  'Lost Documents – Paper records misplaced or lost.',
  'Stolen from Vehicle – Documents or devices stolen from a council or personal vehicle.',
  'Incorrect Disposal – Sensitive data disposed of improperly (e.g., not shredded).',
  'Misaddressed Mail – Letters or packages sent to the wrong recipient.',
  'Contact Centre Error – Personal data disclosed during a call or interaction.',
  'Verbal Disclosure – Sensitive information shared inappropriately in conversation.',
  'Unauthorised Disclosure – Staff accessing or sharing data without proper authority.',
  'Access Control Failure – Inadequate restrictions allowing inappropriate access.',
  'Third-Party Breach – Breach involving contractors, suppliers, or external partners.',
  'Data Shared Without Agreement – Data shared with external parties without a valid agreement or lawful basis.',
];

// ── Notifications ─────────────────────────────────────────────────────────────

export const INDIVIDUALS_INFORMED_OPTIONS: string[] = [
  'Yes',
  'No',
  "We're about to",
  'Not applicable',
];
