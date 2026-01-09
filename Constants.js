/** @file Constants.gs - Configuration constants, headers, enums, and global settings. */

// ---------- Version Information
const ATS_VERSION = '1.0.0';
const ATS_TIER = 'Tier 1';

// ---------- Timezone Configuration
// Note: Must match timeZone in appsscript.json
const TZ = 'America/New_York';

// ---------- Sheet Names
const SHEET_REQUISITIONS = 'Requisitions';
const SHEET_ALL          = 'Candidate Database';
const SHEET_ACTIVE       = 'Active Candidates';
const SHEET_SYS_LOG      = 'SYS_LOGS';
const SHEET_SETTINGS     = 'Settings';
const SHEET_DASHBOARD    = 'Dashboard';
const SHEET_DASH_DATA    = 'Dashboard_Data';

// ---------- Dynamic Header Configuration
const MAX_HEADER_SEARCH_ROWS = 20; // How many rows to scan to find the header

// ---------- Logging
const LOG_PREFIX = '[ATS]';

// ---------- Cache TTLs (seconds)
const CACHE_TTL_SHORT    = 15;   // Recent edit guard window
const CACHE_TTL_MUTATION = 120;  // Header map, etc.

// ---------- Lock Timeouts (milliseconds)
const LOCK_TIMEOUT_MS = 5000;      // Standard operations (5 seconds)
const LOCK_TIMEOUT_LONG_MS = 8000; // Long-running operations (8 seconds)

// ---------- Debounce Delays (milliseconds)
const DEBOUNCE_RECONCILE_MS = 5000;     // Reconciliation delay (5 seconds)
const DEBOUNCE_LINK_HYGIENE_MS = 3000;  // Link hygiene delay (3 seconds)

// ---------- Import/Resume Limits
const IMPORT_MAX_ROWS = 1000;           // Maximum rows per import
const IMPORT_MAX_FIELD_LENGTH = 5000;   // Maximum characters per field
const MAX_FILE_SIZE_MB = 10;            // Maximum file size for resume scanning
const MAX_FILES_PER_RUN = 50;           // Maximum files to process per execution
const MAX_TEXT_LENGTH = 100000;         // Maximum characters to extract from file

// ---------- Validation Limits
const PHONE_MIN_DIGITS = 10;  // Minimum digits in a valid phone number
const PHONE_MAX_DIGITS = 15;  // Maximum digits in a valid phone number

// ---------- Cache Keys
const CACHE_KEYS = {
  HDR_INFO: (sheetId) => `h_info:${sheetId}`,
  MUTE_ROW: (sheetName, key) => `mute:${sheetName}:${key}`,
  REC_GUARD: () => 'recursion:guard',
};

// ---------- Document Properties Keys
const PROP_QUEUE_JOBIDS    = 'ATS:queue:jobIds';
const PROP_QUEUE_SCHEDULED = 'ATS:queue:scheduled';
const PROP_FORM_ID         = 'ATS:form:id';
const PROP_JOBSEQ_PREFIX   = 'ATS:jobseq:';
const PROP_SETTINGS_HASH   = 'ATS:settings:hash';

// ---------- Job Status Sets
const OPEN_STATUSES = new Set(['Open', 'On Hold']);

// ---------- Requisitions Headers
const H_REQ = {
  JobStatus:          'Job Status',
  JobTitle:           'Job Title',
  JobID:              'Job ID',
  Priority:           'Priority',
  ReasonForHire:      'Reason For Hire?',
  WorkModel:          'Work Model',
  ReportingLocation:  'Reporting Location',
  HeadcountTarget:    'Headcount Target',
  EmploymentType:     'Employment Type',
  MinSalary:          'Minimum Salary',
  MaxSalary:          'Maximum Salary',
  HourlyMin:          'Hourly Pay Rate Minimum',
  HourlyMax:          'Hourly Pay Rate Maximum',
  JobOwner:           'Job Owner/Recruiter',
  HiringManager:      'Hiring Manager',
  Created:            'Created Date',
  Opened:             'Opened Date',
  DaysOpen:           'Days Open',
  OnHoldDate:         'On Hold Date',
  ClosedDate:         'Closed Date',
  PositionHiredDate:  'Position Hired Date',
  HiredCandidateName: 'Hired Candidate\'s Name',
};

// ---------- Candidate Database Headers
const H_ALL = {
  FullName:       'Full Name',
  Resume:         'Resume Link',
  LinkedIn:       'LinkedIn Profile',
  Phone:          'Phone Number',
  Email:          'Email Address',
  HomeAddress:    'Home Address',
  City:           'City',
  State:          'State',
  Zip:            'Zip Code',
  TargetSalary:   'Targeted\nCompensation (Salary)',
  TargetHourly:   'Targeted\nCompensation (Hourly)',
  WorkPref:       'Work Environment Preference',
  JobID:          'Job ID',
  JobStatus:      'Job Status',
  JobTitle:       'Job Title',
  Stage:          'Candidate\nWorkflow Status',
  RejectedReason: 'Rejected Reasoning',
  InterviewNotes: 'Interview Notes',
  Relocate:       'Willingness to Relocate?',
  Source:         'Candidate Source',
  Created:        'Created Date',
  Updated:        'Last Updated',
  HiredDate:      'Hired Date',
};

// ---------- Active Candidates Headers
const H_ACT = {
  FullName:       'Full Name',
  JobID:          'Job ID',
  JobStatus:      'Job Status',
  JobTitle:       'Job Title',
  Stage:          'Candidate\nWorkflow Status',
  RejectedReason: 'Rejected Reasoning',
  InterviewNotes: 'Interview Notes',
  Resume:         'Resume Link',
  LinkedIn:       'LinkedIn Profile',
  Phone:          'Phone Number',
  Email:          'Email Address',
  City:           'City',
  State:          'State',
  TargetSalary:   'Targeted\nCompensation (Salary)',
  TargetHourly:   'Targeted\nCompensation (Hourly)',
};

// ---------- Anchor Headers for Dynamic Finding
const ANCHOR_HEADER_REQ = H_REQ.JobID;
const ANCHOR_HEADER_ALL = H_ALL.FullName;
const ANCHOR_HEADER_ACT = H_ACT.FullName;

// ---------- Mirrored Fields (All ↔ Active)
const MIRRORED_FIELDS = [
  H_ALL.Stage, 
  H_ALL.RejectedReason, 
  H_ALL.InterviewNotes,
  H_ALL.Resume, 
  H_ALL.LinkedIn, 
  H_ALL.Phone, 
  H_ALL.Email,
  H_ALL.City, 
  H_ALL.State, 
  H_ALL.TargetSalary, 
  H_ALL.TargetHourly,
  H_ALL.JobTitle, 
  H_ALL.JobStatus, 
  H_ALL.FullName, 
  H_ALL.JobID,
  H_ALL.Created, 
  H_ALL.HiredDate
];