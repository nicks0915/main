// ===== CORE CONFIGURATION =====
// This file contains ONLY configuration constants and mappings
// Web app entry points, getters, and utility functions are in Code_util.gs

// ===== CONFIGURATION CONSTANTS =====

const CONFIG = {
  DASHBOARD: {
    name: 'Entertainment On Green Dashboard'
  },

  // ===== FEATURE FLAGS =====
  // Set to true to ENABLE a feature, false to DISABLE (button/checkbox remains visible but is disabled)
  FEATURE_FLAGS: {
    betaLaunchGating: false,   // true = Beta Launch Gating checkbox is enabled; false = disabled (greyed out)
    gtmLaunchGating: false,    // true = GTM Launch Gating checkbox is enabled; false = disabled (greyed out)
    assessmentsTab: false,     // true = Assessments tab button is clickable; false = disabled (greyed out)
    testingTab: false          // true = Testing Status tab button is clickable; false = disabled (greyed out)
  },

  
  JIRA: {
    baseUrl: 'https://api.atlassian.com/ex/jira/7953168e-96a0-48c2-88d1-212eabe7c2ef'
    // Credentials stored in Script Properties:
    //   JIRA_SERVICE_ACCOUNT_ID — Service Account ID provided by admin
    //   JIRA_SECRET_KEY         — Secret Key provided by admin
    //   JIRA_CLOUD_ID           — 7953168e-96a0-48c2-88d1-212eabe7c2ef
  },
  
  FILTERS: {
    defects: '35173',      // Main defects filter
    epics: '38217',        // Epics and stories filter
    assessments: '38217',  // Assessments filter (same as epics for now)
    gtm: '35173'           // GTM launch gating filter
  },
  
  SHEETS: {
    testingStatusId: '1vSVw-vWtsX9ZIPgYrUtpj2uMzA6-NUrlliZMQRYOG00',
    testingStatusGid: 1819050178,
    assessmentsExportId: '1mh-GcZfujwO5Via9dPG_iC7MzpA_VYjnpKzurtwxxck',
    assessmentsExportGid: 0
  },
  
  SLACK: {
    // ===== SLACK CHANNELS =====
    // Add more channels here as needed. Each entry needs an id and a name.
    channels: [
      { id: 'C09KTBN4AA1', name: 'TestNotification' },
      { id: 'C03UMGV7DDE', name: 'Prod_Green_Commerce' }
      // Example: { id: 'CXXXXXXXXX', name: 'My New Channel' }
    ]
    // Bot Token is stored in Script Properties as 'SLACK_BOT_TOKEN'
  }
};

const FIELD_MAPPINGS = {
  // Defects-specific fields
  severity: 'customfield_19351',              // Severity SODP (cf[19351]) — used for Bug issue type
  severityAlt: 'customfield_19632',           // Severity of Defect (cf[19632]) — used for Experience Defect issue type
  application: 'customfield_19182',           // Application (cf[19182]) — used for Bug issue type
  applicationBug: 'customfield_18937',        // Application Name (Bug) (cf[18937]) — used for Experience Defect issue type
  environmentBug: 'customfield_19453',        // Environments (migrated) (cf[19453]) — used for Bug issue type
  environmentExpDefect: 'customfield_13505',  // Environment (cf[13505]) — used for Experience Defect issue type
  defectAge: 'customfield_19352',             // Defect Age (if used)
  
  // Epics-specific fields
  health: 'customfield_19357',          // Health (cf[19357])
  statusDetails: 'customfield_19361',   // Status Details/Risk (cf[19361])
  
  // Assessments-specific fields
  costEstimate: 'customfield_17265',    // Gate 2 - Cost Estimate (cf[17265])
  
  // Common fields
  startDate: 'customfield_21750'        // Start Date (cf[21750])
};

const DIRECTOR_TEAM_MAPPING = {
  // ===== EXISTING ENTRIES (preserved) =====
  'SoE WebApp': { name: 'Natasha', displayName: 'Natasha' },
  'Unified Hardware & Manage Apps (UHMA)': { name: 'Natasha', displayName: 'Natasha' },
  'C3': { name: 'andre', displayName: 'andre' },
  'NGC': { name: 'andre', displayName: 'andre' },
  'C3X': { name: 'Neal', displayName: 'Neal' },
  'Universal MFE': { name: 'Neal', displayName: 'Neal' },
  'EOM': { name: 'nicole.tytaneck', displayName: 'nicole.tytaneck' },
  'Merlin / NC_Cloud BSS': { name: 'nicole.tytaneck', displayName: 'nicole.tytaneck' },
  'NC Vendor': { name: 'nicole.tytaneck', displayName: 'nicole.tytaneck' },
  'Offer Admin': { name: 'nicole.tytaneck', displayName: 'nicole.tytaneck' },
  'CASA': { name: 'Brianne', displayName: 'brianne' },
  'WFM': { name: 'Nelson', displayName: 'Nelson' },
  'Change Request': { name: 'Amanda Palmer', displayName: 'Amanda Palmer' },
  'Enterprise Arch': { name: 'Pavel Belyavsky', displayName: 'Pavel Belyavsky' },
  'MyTELUS': { name: 'Vatsal Bhatt', displayName: 'Vatsal' },
  'ECP': { name: 'Vatsal Bhatt', displayName: 'Vatsal' },

  // ===== NEW ENTRIES FROM DirectorMapping.csv (Column E → Column A) =====

  // Nicole Tytaneck
  'CloudBSS': { name: 'Nicole Tytaneck', displayName: 'Nicole Tytaneck' },
  'GEM': { name: 'Nicole Tytaneck', displayName: 'Nicole Tytaneck' },
  'Bundle Qualification API': { name: 'Nicole Tytaneck', displayName: 'Nicole Tytaneck' },

  // Natasha Lander
  'Service Address Selection MFE': { name: 'Natasha Lander', displayName: 'Natasha Lander' },
  'Appointment Selection MFE': { name: 'Natasha Lander', displayName: 'Natasha Lander' },
  'Cross Sell MFE': { name: 'Natasha Lander', displayName: 'Natasha Lander' },
  'Commerce Shell': { name: 'Natasha Lander', displayName: 'Natasha Lander' },
  'Commerce Session API': { name: 'Natasha Lander', displayName: 'Natasha Lander' },
  'Auth Status API': { name: 'Natasha Lander', displayName: 'Natasha Lander' },
  'Hardware Selection MFE': { name: 'Natasha Lander', displayName: 'Natasha Lander' },
  'Send Quote MFE': { name: 'Natasha Lander', displayName: 'Natasha Lander' },
  'Hardware Shell': { name: 'Natasha Lander', displayName: 'Natasha Lander' },
  'Manage Licenses MFE': { name: 'Natasha Lander', displayName: 'Natasha Lander' },
  'Hardware Return MFE': { name: 'Natasha Lander', displayName: 'Natasha Lander' },
  'Product Subscription Overview (PSO) MFE': { name: 'Natasha Lander', displayName: 'Natasha Lander' },
  'Purple Shop & Buy': { name: 'Natasha Lander', displayName: 'Natasha Lander' },

  // Neal McGann
  'Checkout Shell': { name: 'Neal McGann', displayName: 'Neal McGann' },
  'Add On Selection MFE (V2)': { name: 'Neal McGann', displayName: 'Neal McGann' },
  'Payment MFE': { name: 'Neal McGann', displayName: 'Neal McGann' },
  'Send Communication MFE': { name: 'Neal McGann', displayName: 'Neal McGann' },
  'Pro Install MFE': { name: 'Neal McGann', displayName: 'Neal McGann' },
  'Credit Assessment MFE': { name: 'Neal McGann', displayName: 'Neal McGann' },
  'Product Type Selection MFE': { name: 'Neal McGann', displayName: 'Neal McGann' },
  'Device Selection MFE': { name: 'Neal McGann', displayName: 'Neal McGann' },
  'Phone Number Selection MFE': { name: 'Neal McGann', displayName: 'Neal McGann' },
  'Subscriber Information MFE': { name: 'Neal McGann', displayName: 'Neal McGann' },
  'Sales Summary MFE': { name: 'Neal McGann', displayName: 'Neal McGann' },
  'Order Summary MFE': { name: 'Neal McGann', displayName: 'Neal McGann' },
  'Device Selection (Sim) MFE': { name: 'Neal McGann', displayName: 'Neal McGann' },
  'Coupon Code MFE': { name: 'Neal McGann', displayName: 'Neal McGann' },
  'Package Selection MFE': { name: 'Neal McGann', displayName: 'Neal McGann' },
  'Create Account MFE': { name: 'Neal McGann', displayName: 'Neal McGann' },
  'Create Customer MFE': { name: 'Neal McGann', displayName: 'Neal McGann' },
  'Order Status MFE': { name: 'Neal McGann', displayName: 'Neal McGann' },
  'Digital Signature MFE': { name: 'Neal McGann', displayName: 'Neal McGann' },
  'Universal MFE + Shell (UOST)': { name: 'Neal McGann', displayName: 'Neal McGann' },
  'UDE Checkout (purple /yellow)': { name: 'Neal McGann', displayName: 'Neal McGann' },

  // Vatsal Bhatt
  'Manage Shell': { name: 'Vatsal Bhatt', displayName: 'Vatsal Bhatt' },
  'Cancel / Renew MFE': { name: 'Vatsal Bhatt', displayName: 'Vatsal Bhatt' },
  'My TELUS Billing UI': { name: 'Vatsal Bhatt', displayName: 'Vatsal Bhatt' },
  'My TELUS Payments UI': { name: 'Vatsal Bhatt', displayName: 'Vatsal Bhatt' },
  'My TELUS Appointments UI': { name: 'Vatsal Bhatt', displayName: 'Vatsal Bhatt' },
  'My TELUS Purple Manage Apps': { name: 'Vatsal Bhatt', displayName: 'Vatsal Bhatt' },
  'My TELUS Overview UI': { name: 'Vatsal Bhatt', displayName: 'Vatsal Bhatt' },
  'overview-api': { name: 'Vatsal Bhatt', displayName: 'Vatsal Bhatt' },
  'My TELUS Navigation': { name: 'Vatsal Bhatt', displayName: 'Vatsal Bhatt' },
  'My TELUS App (Expo - Front end)': { name: 'Vatsal Bhatt', displayName: 'Vatsal Bhatt' },
  'My TELUS App (BFF)': { name: 'Vatsal Bhatt', displayName: 'Vatsal Bhatt' },
  'My TELUS Plans & Devices': { name: 'Vatsal Bhatt', displayName: 'Vatsal Bhatt' },
  'My TELUS B2B WLN': { name: 'Vatsal Bhatt', displayName: 'Vatsal Bhatt' },
  'My TELUS B2B WLS': { name: 'Vatsal Bhatt', displayName: 'Vatsal Bhatt' },
  'customer-api': { name: 'Vatsal Bhatt', displayName: 'Vatsal Bhatt' },

  // Andre Medeiros
  'Promotion Selection MFE': { name: 'Andre Medeiros', displayName: 'Andre Medeiros' },
  'Rate Plan Selection MFE': { name: 'Andre Medeiros', displayName: 'Andre Medeiros' },
  'Port-In MFE': { name: 'Andre Medeiros', displayName: 'Andre Medeiros' },
  'Add On Selection MFE': { name: 'Andre Medeiros', displayName: 'Andre Medeiros' },
  'Product Qualification': { name: 'Andre Medeiros', displayName: 'Andre Medeiros' },
  'NCCS': { name: 'Andre Medeiros', displayName: 'Andre Medeiros' },
  'C30C': { name: 'Andre Medeiros', displayName: 'Andre Medeiros' },
  'Channel Dashboard': { name: 'Andre Medeiros', displayName: 'Andre Medeiros' },

  // Nelson Gillette
  'TMF-645 (Service Qual API)': { name: 'Nelson Gillette', displayName: 'Nelson Gillette' },
  'Spatial Net': { name: 'Nelson Gillette', displayName: 'Nelson Gillette' },
  'AMS': { name: 'Nelson Gillette', displayName: 'Nelson Gillette' },
  'AES': { name: 'Nelson Gillette', displayName: 'Nelson Gillette' },
  'WFM Sales Force Scheduler': { name: 'Nelson Gillette', displayName: 'Nelson Gillette' },

  // Harbinder Mann
  'CAMP': { name: 'Harbinder Mann', displayName: 'Harbinder Mann' },
  'CMF': { name: 'Harbinder Mann', displayName: 'Harbinder Mann' },
  'OutCollect': { name: 'Harbinder Mann', displayName: 'Harbinder Mann' },
  'EPS - Enterprise Payment Systems': { name: 'Harbinder Mann', displayName: 'Harbinder Mann' },

  // Matt Powell
  'Casa Partner': { name: 'Matt Powell', displayName: 'Matt Powell' },
  'Casa TBS': { name: 'Matt Powell', displayName: 'Matt Powell' },
  'Casa CE': { name: 'Matt Powell', displayName: 'Matt Powell' },

  // Andrew Ah Yong
  'CES': { name: 'Andrew Ah Yong', displayName: 'Andrew Ah Yong' },
  'Subscription Management': { name: 'Andrew Ah Yong', displayName: 'Andrew Ah Yong' },
  'Knowbility Decommission': { name: 'Andrew Ah Yong', displayName: 'Andrew Ah Yong' },
  'Knowbility to GCP': { name: 'Andrew Ah Yong', displayName: 'Andrew Ah Yong' },
  'Bill Presentment & Analytics': { name: 'Andrew Ah Yong', displayName: 'Andrew Ah Yong' },
  'Billing API': { name: 'Andrew Ah Yong', displayName: 'Andrew Ah Yong' },
  'Enterprise Account Management APIs': { name: 'Andrew Ah Yong', displayName: 'Andrew Ah Yong' },
  'Kenan': { name: 'Andrew Ah Yong', displayName: 'Andrew Ah Yong' },

  // Jerome Berube
  'TechHub': { name: 'Jerome Berube', displayName: 'Jerome Berube' },
  'PaaS': { name: 'Jerome Berube', displayName: 'Jerome Berube' },
  'SWORM': { name: 'Jerome Berube', displayName: 'Jerome Berube' },

  // Alyssa Borges
  'Experience Design': { name: 'Alyssa Borges', displayName: 'Alyssa Borges' },

  // Alex Itelman
  'SSNS': { name: 'Alex Itelman', displayName: 'Alex Itelman' },

  // Jim Bremner
  'Authentication MFE': { name: 'Jim Bremner', displayName: 'Jim Bremner' },

  // Dinesh M
  'NBA MFE': { name: 'Dinesh M', displayName: 'Dinesh M' }
};

const API_SETTINGS = {
  maxRetries: 3,
  maxResults: 100,
  retryDelayBase: 1000,  // Base delay in milliseconds for exponential backoff
  userAgent: 'Google-Apps-Script-EOG-Dashboard/1.0'
};

const PROJECT_LINKS = {
  jiraTimeline: {
    name: 'JIRA Epics/Timelines',
    url: 'https://telus-cio.atlassian.net/jira/plans/1204/scenarios/1212/timeline?vid=3448',
    icon: 'jira'
  },
  rpp: {
    name: 'RPP',
    url: 'https://rpp.tsl.telus.com/team/projects.cfm?id=124760',
    icon: 'rpp'
  },
  googleDrive: {
    name: 'Google Project Drive',
    url: 'https://drive.google.com/drive/folders/0AIcnOLoiINmjUk9PVA',
    icon: 'drive'
  },
  figma: {
    name: 'Figma (Customer Journey)',
    url: 'https://figma.com/board/GF7sX6vweyn2LjUYyBlzNr/Optik-TV-Strategy?t=5ZjzGF02q3urkRuY-0',
    icon: 'figma'
  },
  gate0: {
    name: 'Gate 0',
    url: 'https://docs.google.com/spreadsheets/d/14_Chs4zFgVVYepayd8UyuzmNj8RysslbycMlJxGsN4I/edit?gid=1200578436#gid=1200578436',
    icon: 'gate0'
  }
};
