// ===== DEFECTS-SPECIFIC FUNCTIONS =====
// This file contains all defects-related data fetching and transformation logic
// Configuration and shared utilities are in Core.gs

/**
 * Fetch data from JIRA using the configured filter with token-based pagination.
 * Uses POST with a JSON body (same pattern as GREEN_code.gs) to ensure all
 * custom fields (customfield_18937, customfield_19632, customfield_13505, etc.)
 * are reliably returned by the JIRA API.
 */
function fetchJiraData() {
  try {
    console.log('=== STARTING JIRA DATA FETCH WITH TOKEN-BASED PAGINATION ===');
    
    // First, validate JIRA connection and configuration
    const validationResult = validateJiraConnection();
    if (!validationResult.success) {
      throw new Error(`JIRA connection validation failed: ${validationResult.error}`);
    }
    
    let allIssues = [];
    let nextPageToken = null;
    const maxResults = API_SETTINGS.maxResults;
    let total = 0;
    let pageCount = 0;

    // Base URL — POST to /search/jql (same endpoint as GREEN_code.gs)
    const searchUrl = `${CONFIG.JIRA.baseUrl}/rest/api/3/search/jql`;

    // Auth header from existing helper
    const baseOptions = getJiraFetchOptions();
    const authHeader = baseOptions.headers['Authorization'];
    
    do {
      pageCount++;
      console.log(`=== PAGINATION CALL ${pageCount} ===`);

      // Build POST body — fields as a JSON array so JIRA reliably returns all custom fields
      const requestBody = buildJiraRequestBody(maxResults, nextPageToken);
      console.log(`POST body: ${JSON.stringify(requestBody).substring(0, 200)}...`);

      const postOptions = {
        method: 'POST',
        headers: {
          'Authorization': authHeader,
          'Accept': 'application/json',
          'Content-Type': 'application/json',
          'User-Agent': 'GoogleAppsScript/1.0'
        },
        payload: JSON.stringify(requestBody),
        muteHttpExceptions: true
      };
      
      let response;
      let success = false;
      
      // Retry logic with exponential backoff
      for (let attempt = 0; attempt <= API_SETTINGS.maxRetries; attempt++) {
        try {
          response = UrlFetchApp.fetch(searchUrl, postOptions);
          const responseCode = response.getResponseCode();
          
          console.log(`Attempt ${attempt + 1}: Response code ${responseCode}`);
          
          if (responseCode === 200) {
            success = true;
            break;
          } else if (responseCode === 401) {
            throw new Error('Authentication failed. Please check your JIRA email and API token in CONFIG.JIRA.');
          } else if (responseCode === 403) {
            throw new Error('Access forbidden. Please check if your account has permission to access the specified filter.');
          } else if (responseCode === 404) {
            throw new Error(`Filter not found. Please verify that filter ID ${CONFIG.FILTERS.defects} exists and is accessible.`);
          } else if (responseCode === 410) {
            throw new Error('JIRA API endpoint has been removed. Please check the JIRA API documentation for the correct endpoint.');
          } else if (responseCode >= 500) {
            // Server error - retry
            if (attempt < API_SETTINGS.maxRetries) {
              const delay = Math.pow(2, attempt) * API_SETTINGS.retryDelayBase;
              console.log(`Server error ${responseCode}, retrying in ${delay}ms...`);
              Utilities.sleep(delay);
              continue;
            } else {
              throw new Error(`JIRA server error ${responseCode}: ${response.getContentText()}`);
            }
          } else {
            // Other error codes
            const errorDetails = response.getContentText();
            console.error(`JIRA API Error ${responseCode}:`, errorDetails);
            
            // Try to parse error details
            try {
              const errorData = JSON.parse(errorDetails);
              if (errorData.errorMessages && errorData.errorMessages.length > 0) {
                throw new Error(`JIRA API returned ${responseCode}: ${errorData.errorMessages.join(', ')}`);
              }
            } catch (parseError) {
              // If we can't parse the error, use the raw response
            }
            
            throw new Error(`JIRA API returned ${responseCode}: ${errorDetails}`);
          }
        } catch (error) {
          if (attempt === API_SETTINGS.maxRetries) {
            throw error;
          }
          console.log(`Attempt ${attempt + 1} failed: ${error.message}`);
          const delay = Math.pow(2, attempt) * API_SETTINGS.retryDelayBase;
          Utilities.sleep(delay);
        }
      }
      
      if (!success) {
        throw new Error('Failed to fetch data after maximum retries');
      }
      
      const data = JSON.parse(response.getContentText());
      
      // Log response structure for debugging
      console.log(`Response keys: ${Object.keys(data).join(', ')}`);
      console.log(`Issues returned: ${data.issues ? data.issues.length : 0}`);
      
      if (pageCount === 1) {
        total = data.total;
        console.log(`Total issues in filter: ${total}`);
        
        if (total === 0) {
          console.warn('No issues found in the specified filter');
          return [['Issue key', 'Summary', 'Custom field (Severity SODP)', 'Priority', 'Status', 'Due date', 'Custom field (Defect Age)', 'Custom field (Application )', 'Labels', 'Assignee']];
        }
      }
      
      // Add issues to our collection
      if (data.issues && data.issues.length > 0) {
        allIssues = allIssues.concat(data.issues);
        console.log(`Fetched ${allIssues.length} of ${total} issues`);
      }
      
      // Extract nextPageToken for subsequent requests
      nextPageToken = data.nextPageToken || null;
      console.log(`Next page token: ${nextPageToken ? nextPageToken : 'None (last page)'}`);
      
      // Continue if we have more pages (via token or if we haven't reached total)
    } while (nextPageToken || (allIssues.length < total && allIssues.length > 0));
    
    console.log(`=== PAGINATION COMPLETE ===`);
    console.log(`Total API calls made: ${pageCount}`);
    console.log(`Successfully fetched all ${allIssues.length} issues from JIRA`);
    console.log(`Expected: ${total}, Retrieved: ${allIssues.length}, Match: ${allIssues.length === total ? 'YES' : 'NO'}`);
    
    return transformJiraData(allIssues);
    
  } catch (error) {
    console.error('Error fetching JIRA data:', error);
    
    // Enhanced error message for user
    let userFriendlyMessage = 'Failed to fetch data from JIRA.';
    
    if (error.message.includes('Authentication failed')) {
      userFriendlyMessage += ' Please check your JIRA credentials.';
    } else if (error.message.includes('Access forbidden')) {
      userFriendlyMessage += ' Please check your JIRA permissions.';
    } else if (error.message.includes('Filter not found')) {
      userFriendlyMessage += ' Please verify the filter ID is correct.';
    } else if (error.message.includes('API endpoint has been removed')) {
      userFriendlyMessage += ' The JIRA API has been updated. Please contact your administrator.';
    } else {
      userFriendlyMessage += ` Error: ${error.message}`;
    }
    
    throw new Error(userFriendlyMessage);
  }
}

/**
 * Build the POST request body for the JIRA /search/jql endpoint.
 * Uses a JSON array for fields — same pattern as GREEN_code.gs — so JIRA
 * reliably returns all custom fields including customfield_18937, customfield_19632,
 * customfield_13505, etc.
 */
function buildJiraRequestBody(maxResults = 100, nextPageToken = null) {
  const jql = `filter=${CONFIG.FILTERS.defects}`;
  console.log(`Using filter ${CONFIG.FILTERS.defects} with JQL: ${jql}`);

  const body = {
    jql: jql,
    maxResults: maxResults,
    fields: [
      'summary', 'status', 'priority', 'assignee', 'labels',
      'created', 'duedate', 'issuetype',
      FIELD_MAPPINGS.severity,           // customfield_19351 — Severity SODP (Bug)
      FIELD_MAPPINGS.severityAlt,        // customfield_19632 — Severity of Defect (Experience Defect)
      FIELD_MAPPINGS.application,        // customfield_19182 — Application (Bug)
      FIELD_MAPPINGS.applicationBug,     // customfield_18937 — Application Name (Bug) (Experience Defect)
      FIELD_MAPPINGS.environmentBug,     // customfield_19453 — Environments (migrated) (Bug)
      FIELD_MAPPINGS.environmentExpDefect // customfield_13505 — Environment checkboxes (Experience Defect)
    ]
  };

  // Add nextPageToken for subsequent pages
  if (nextPageToken) {
    body.nextPageToken = nextPageToken;
    console.log(`Using nextPageToken for pagination: ${nextPageToken}`);
  }

  return body;
}

/**
 * Transform JIRA JSON data to CSV-like format for the dashboard
 */
function transformJiraData(issues) {
  try {
    console.log('=== STARTING TRANSFORM JIRA DATA ===');
    
    // Comprehensive input validation
    if (!issues) {
      console.warn('Issues parameter is null or undefined');
      return [['Issue key', 'Summary', 'Custom field (Severity SODP)', 'Priority', 'Status', 'Due date', 'Custom field (Defect Age)', 'Custom field (Application )', 'Labels', 'Assignee']];
    }
    
    if (!Array.isArray(issues)) {
      console.warn('Issues parameter is not an array:', typeof issues);
      return [['Issue key', 'Summary', 'Custom field (Severity SODP)', 'Priority', 'Status', 'Due date', 'Custom field (Defect Age)', 'Custom field (Application )', 'Labels', 'Assignee']];
    }
    
    if (issues.length === 0) {
      console.warn('Issues array is empty');
      return [['Issue key', 'Summary', 'Custom field (Severity SODP)', 'Priority', 'Status', 'Due date', 'Custom field (Defect Age)', 'Custom field (Application )', 'Labels', 'Assignee']];
    }
    
    console.log(`Processing ${issues.length} issues`);
    
    const csvData = [];
    
    // Add header row — no separate "Severity of Defect" column; Environment column added
    csvData.push([
        'Issue key',
        'Summary',
        'Custom field (Severity SODP)',
        'Priority',
        'Status',
        'Due date',
        'Custom field (Defect Age)',
        'Custom field (Application )',
        'Environment',
        'Labels',
        'Assignee'
    ]);

  // Field mappings from Core.gs
  // Severity:    customfield_19351 (Bug) | customfield_19632 (Experience Defect)
  // Application: customfield_19182 (Bug) | customfield_18937 (Experience Defect)
  // Environment: customfield_19453 (Bug) | customfield_13505 (Experience Defect)
  const severityField        = FIELD_MAPPINGS.severity;           // customfield_19351
  const severityAltField     = FIELD_MAPPINGS.severityAlt;        // customfield_19632
  const applicationField     = FIELD_MAPPINGS.application;        // customfield_19182
  const applicationBugField  = FIELD_MAPPINGS.applicationBug;     // customfield_18937
  const environmentBugField  = FIELD_MAPPINGS.environmentBug;     // customfield_19453
  const environmentExpField  = FIELD_MAPPINGS.environmentExpDefect; // customfield_13505
  
  if (issues.length > 0) {
    // Add comprehensive null checking before any field access
    const firstIssue = issues[0];
    if (!firstIssue || !firstIssue.fields) {
      console.warn('First issue is malformed, skipping field validation');
    } else {
      const sampleFields = firstIssue.fields;
      
      // Safely check field status for debugging with proper null checking
      try {
        if (sampleFields && typeof sampleFields === 'object' && sampleFields.hasOwnProperty(severityField)) {
          console.log(`Primary Severity field (${severityField}) found: ${JSON.stringify(sampleFields[severityField])}`);
        } else {
          console.log(`Primary Severity field (${severityField}) NOT found — will use fallback`);
        }
        
        if (sampleFields && typeof sampleFields === 'object' && sampleFields.hasOwnProperty(severityAltField)) {
          console.log(`Fallback Severity field (${severityAltField}) found: ${JSON.stringify(sampleFields[severityAltField])}`);
        } else {
          console.log(`Fallback Severity field (${severityAltField}) also not found`);
        }
        
        if (sampleFields && typeof sampleFields === 'object' && sampleFields.hasOwnProperty(applicationField)) {
          console.log(`Application field (${applicationField}) found: ${JSON.stringify(sampleFields[applicationField])}`);
        } else {
          console.log(`Warning: Application field ${applicationField} not found in response`);
        }
      } catch (fieldCheckError) {
        console.warn('Error during field validation:', fieldCheckError.message);
      }
    }
  }
  
  console.log(`Primary Severity field: ${severityField}`);
  console.log(`Fallback Severity field: ${severityAltField}`);
  console.log(`Application field: ${applicationField}`);
  
  // Counters for summary debug log
  let countPrimaryFilled = 0;
  let countAltFilled = 0;
  let countAltUsedAsFallback = 0;
  let countBothEmpty = 0;

  // Transform each issue
  issues.forEach(issue => {
    // Add defensive checks for issue structure
    if (!issue || !issue.fields) {
      console.warn('Skipping malformed issue:', JSON.stringify(issue));
      return;
    }
    
    const fields = issue.fields;
    
    // Determine routing by issue key prefix (more reliable than issuetype field)
    // SODP-xxxxx → Bug fields   |   GREEN-xxxxx → Experience Defect fields
    const issueKeyPrefix = issue.key ? issue.key.split('-')[0].toUpperCase() : '';
    const isBug = issueKeyPrefix === 'SODP';
    const isExpDefect = issueKeyPrefix === 'GREEN';
    // Keep issueTypeName for display/debug only
    const issueTypeName = (fields.issuetype && fields.issuetype.name) ? fields.issuetype.name : issueKeyPrefix;

    // ── DEBUG: Log issue type routing so we can verify field selection ────────
    console.log(`[DEBUG] ${issue.key} | keyPrefix="${issueKeyPrefix}" | issueType="${issueTypeName}" | isBug=${isBug} | isExpDefect=${isExpDefect}`);
    console.log(`[DEBUG] ${issue.key} | severity field → ${isBug ? severityField : severityAltField} | raw: ${JSON.stringify(isBug ? fields[severityField] : fields[severityAltField])}`);
    console.log(`[DEBUG] ${issue.key} | application field → ${isExpDefect ? applicationBugField : applicationField} | raw: ${JSON.stringify(isExpDefect ? fields[applicationBugField] : fields[applicationField])}`);
    console.log(`[DEBUG] ${issue.key} | environment field → ${isBug ? environmentBugField : environmentExpField} | raw: ${JSON.stringify(isBug ? fields[environmentBugField] : fields[environmentExpField])}`);

    // Calculate defect age in days with error handling
    let defectAge = 0;
    try {
      if (fields.created) {
        const createdDate = new Date(fields.created);
        const today = new Date();
        defectAge = Math.floor((today - createdDate) / (1000 * 60 * 60 * 24));
      }
    } catch (error) {
      console.warn(`Error calculating defect age for ${issue.key}:`, error);
    }
    
    // Format due date with error handling
    let dueDate = '';
    try {
      dueDate = fields.duedate ? formatJiraDate(fields.duedate) : '';
    } catch (error) {
      console.warn(`Error formatting due date for ${issue.key}:`, error);
    }
    
    // ── SEVERITY: route by issue type ──────────────────────────────────────
    // Bug              → customfield_19351 (Severity SODP)
    // Experience Defect → customfield_19632 (Severity of Defect — Dropdown)
    let severityValue = '';
    try {
      const sevRawField = isBug ? fields[severityField] : fields[severityAltField];
      if (sevRawField) {
        if (typeof sevRawField === 'string') {
          severityValue = sevRawField.trim();
        } else if (Array.isArray(sevRawField) && sevRawField.length > 0) {
          const first = sevRawField[0];
          severityValue = (typeof first === 'string' ? first : (first.value || first.name || String(first))).trim();
        } else if (sevRawField.value) {
          severityValue = String(sevRawField.value).trim();
        } else if (sevRawField.name) {
          severityValue = String(sevRawField.name).trim();
        }
      }
    } catch (error) {
      console.warn(`Error extracting severity for ${issue.key}:`, error);
    }

    // Update summary counters
    if (severityValue) countPrimaryFilled++;
    else countBothEmpty++;

    // Extract priority value with error handling
    let priorityValue = '';
    try {
      if (fields.priority) {
        priorityValue = fields.priority.name || fields.priority.value || '';
      }
    } catch (error) {
      console.warn(`Error extracting priority for ${issue.key}:`, error);
    }

    // ── APPLICATION: route by issue type ───────────────────────────────────
    // Bug              → customfield_19182 (Application)
    // Experience Defect → customfield_18937 (Application Name (Bug))
    let applicationValue = '';
    try {
      const appRawField = isExpDefect ? fields[applicationBugField] : fields[applicationField];
      if (appRawField) {
        if (typeof appRawField === 'string') {
          applicationValue = appRawField.trim();
        } else if (Array.isArray(appRawField) && appRawField.length > 0) {
          // Handles multi-select / array format: [{value: "..."}, ...] or ["...", ...]
          applicationValue = appRawField.map(a => (typeof a === 'string' ? a : (a.value || a.name || String(a)))).filter(Boolean).join(', ');
        } else if (appRawField.value) {
          applicationValue = String(appRawField.value).trim();
        } else if (appRawField.name) {
          applicationValue = String(appRawField.name).trim();
        }
      }
    } catch (error) {
      console.warn(`Error extracting application for ${issue.key}:`, error);
    }

    // ── ENVIRONMENT: route by issue type ───────────────────────────────────
    // Bug              → customfield_19453 (Environments (migrated) — text/array)
    // Experience Defect → customfield_13505 (Environment — Checkboxes array)
    let environmentValue = '';
    try {
      const envRawField = isBug ? fields[environmentBugField] : fields[environmentExpField];
      if (envRawField) {
        if (typeof envRawField === 'string') {
          environmentValue = envRawField.trim();
        } else if (Array.isArray(envRawField) && envRawField.length > 0) {
          // Handles plain string arrays ["ITN01", ...] and object arrays [{value: "Prod"}, ...]
          environmentValue = envRawField.map(e => (typeof e === 'string' ? e : (e.value || e.name || String(e)))).filter(Boolean).join(', ');
        } else if (envRawField.value) {
          environmentValue = String(envRawField.value).trim();
        } else if (envRawField.name) {
          environmentValue = String(envRawField.name).trim();
        }
      }
    } catch (error) {
      console.warn(`Error extracting environment for ${issue.key}:`, error);
    }
    
    // Extract assignee value with error handling
    let assigneeValue = 'Unassigned';
    try {
      if (fields.assignee) {
        assigneeValue = fields.assignee.displayName || fields.assignee.name || 'Unassigned';
      }
    } catch (error) {
      console.warn(`Error extracting assignee for ${issue.key}:`, error);
    }
    
    // Extract labels with enhanced debugging and error handling
    let labels = '';
    try {
      if (fields.labels) {
        console.log(`Issue ${issue.key} - Labels field raw:`, JSON.stringify(fields.labels));
        
        if (Array.isArray(fields.labels)) {
          if (fields.labels.length > 0) {
            labels = fields.labels.map(label => {
              // Handle both object and string formats
              if (typeof label === 'string') {
                return label;
              } else if (label && label.name) {
                return label.name;
              } else {
                console.log(`Unexpected label format for ${issue.key}:`, JSON.stringify(label));
                return String(label);
              }
            }).filter(l => l && l.trim() !== '').join(', ');
            
            console.log(`Issue ${issue.key} - Processed labels:`, labels);
          } else {
            console.log(`Issue ${issue.key} - Labels array is empty`);
          }
        } else {
          console.log(`Issue ${issue.key} - Labels is not an array:`, typeof fields.labels, JSON.stringify(fields.labels));
          // Handle case where labels might be a single value
          if (typeof fields.labels === 'string') {
            labels = fields.labels;
          }
        }
      } else {
        console.log(`Issue ${issue.key} - No labels field found`);
      }
    } catch (error) {
      console.warn(`Error extracting labels for ${issue.key}:`, error);
    }
    
    // Add the row — columns: Issue key, Summary, Severity, Priority, Status, Due date, Defect Age, Application, Environment, Labels, Assignee
    try {
      csvData.push([
        issue.key || '',
        fields.summary || '',
        severityValue || '',
        priorityValue || '',
        (fields.status && fields.status.name) ? fields.status.name : '',
        dueDate || '',
        defectAge.toString() || '0',
        applicationValue || '',
        environmentValue || '',
        labels || '',
        assigneeValue || 'Unassigned'
      ]);
    } catch (error) {
      console.warn(`Error adding CSV row for ${issue.key}:`, error);
      // Add a minimal row to prevent complete failure
      csvData.push([
        issue.key || 'UNKNOWN',
        'Error processing issue',
        '', '', '', '', '0', '', '', '', 'Unassigned'
      ]);
    }
  });
  
  console.log(`Successfully processed ${csvData.length - 1} issues (excluding header)`);
  console.log(`=== SEVERITY FIELD SUMMARY ===`);
  console.log(`Total issues processed: ${issues.length}`);
  console.log(`Issues with Severity SODP (customfield_19351) filled: ${countPrimaryFilled}`);
  console.log(`Issues with Severity of Defect (customfield_19632) filled: ${countAltFilled}`);
  console.log(`Issues where customfield_19632 was used as fallback for empty customfield_19351: ${countAltUsedAsFallback}`);
  console.log(`Issues where BOTH severity fields are empty: ${countBothEmpty}`);
  return csvData;
  
  } catch (transformError) {
    console.error('Critical error in transformJiraData:', transformError);
    console.error('Error stack:', transformError.stack);
    
    // Return a safe fallback with error information
    return [
      ['Issue key', 'Summary', 'Custom field (Severity SODP)', 'Priority', 'Status', 'Due date', 'Custom field (Defect Age)', 'Custom field (Application )', 'Labels', 'Assignee'],
      ['ERROR', `Transform failed: ${transformError.message}`, '', '', '', '', '0', '', '', 'System']
    ];
  }
}

/**
 * Fetch JIRA data and return as CSV string (called from HTML)
 */
function fetchJiraDataAsCsv() {
  try {
    const csvArray = fetchJiraData();
    return arrayToCsv(csvArray);
  } catch (error) {
    console.error('Error in fetchJiraDataAsCsv:', error);
    throw error;
  }
}

/**
 * Fetch GTM Launch Gating specific data from JIRA with token-based pagination
 */
function fetchGtmJiraData() {
  try {
    console.log('=== STARTING GTM JIRA DATA FETCH WITH TOKEN-BASED PAGINATION ===');
    
    // First, validate JIRA connection and configuration
    const validationResult = validateJiraConnection();
    if (!validationResult.success) {
      throw new Error(`JIRA connection validation failed: ${validationResult.error}`);
    }
    
    let allIssues = [];
    let startAt = 0;
    let nextPageToken = null;
    const maxResults = API_SETTINGS.maxResults;
    let total = 0;
    let pageCount = 0;
    
    const options = getJiraFetchOptions();
    
    do {
      pageCount++;
      const url = buildGtmJiraUrl(startAt, maxResults, nextPageToken);
      console.log(`=== GTM PAGINATION CALL ${pageCount} ===`);
      console.log(`Fetching GTM JIRA data from: ${url}`);
      
      let response;
      let success = false;
      
      // Retry logic with exponential backoff
      for (let attempt = 0; attempt <= API_SETTINGS.maxRetries; attempt++) {
        try {
          response = UrlFetchApp.fetch(url, options);
          const responseCode = response.getResponseCode();
          
          console.log(`GTM Attempt ${attempt + 1}: Response code ${responseCode}`);
          
          if (responseCode === 200) {
            success = true;
            break;
          } else if (responseCode === 401) {
            throw new Error('Authentication failed. Please check your JIRA email and API token in CONFIG.JIRA.');
          } else if (responseCode === 403) {
            throw new Error('Access forbidden. Please check if your account has permission to access the specified filter.');
          } else if (responseCode === 404) {
            throw new Error(`Filter not found. Please verify that filter ID ${CONFIG.FILTERS.gtm} exists and is accessible.`);
          } else if (responseCode === 410) {
            throw new Error('JIRA API endpoint has been removed. Please check the JIRA API documentation for the correct endpoint.');
          } else if (responseCode >= 500) {
            // Server error - retry
            if (attempt < API_SETTINGS.maxRetries) {
              const delay = Math.pow(2, attempt) * API_SETTINGS.retryDelayBase;
              console.log(`GTM Server error ${responseCode}, retrying in ${delay}ms...`);
              Utilities.sleep(delay);
              continue;
            } else {
              throw new Error(`JIRA server error ${responseCode}: ${response.getContentText()}`);
            }
          } else {
            // Other error codes
            const errorDetails = response.getContentText();
            console.error(`GTM JIRA API Error ${responseCode}:`, errorDetails);
            
            // Try to parse error details
            try {
              const errorData = JSON.parse(errorDetails);
              if (errorData.errorMessages && errorData.errorMessages.length > 0) {
                throw new Error(`JIRA API returned ${responseCode}: ${errorData.errorMessages.join(', ')}`);
              }
            } catch (parseError) {
              // If we can't parse the error, use the raw response
            }
            
            throw new Error(`JIRA API returned ${responseCode}: ${errorDetails}`);
          }
        } catch (error) {
          if (attempt === API_SETTINGS.maxRetries) {
            throw error;
          }
          console.log(`GTM Attempt ${attempt + 1} failed: ${error.message}`);
          const delay = Math.pow(2, attempt) * API_SETTINGS.retryDelayBase;
          Utilities.sleep(delay);
        }
      }
      
      if (!success) {
        throw new Error('Failed to fetch GTM data after maximum retries');
      }
      
      const data = JSON.parse(response.getContentText());
      
      // Log response structure for debugging
      console.log(`GTM Response keys: ${Object.keys(data).join(', ')}`);
      console.log(`GTM Issues returned: ${data.issues ? data.issues.length : 0}`);
      
      if (pageCount === 1) {
        total = data.total;
        console.log(`Total GTM issues in filter: ${total}`);
        
        if (total === 0) {
          console.warn('No GTM issues found in the specified filter');
          return [['Issue key', 'Summary', 'Custom field (Severity SODP)', 'Priority', 'Status', 'Due date', 'Custom field (Defect Age)', 'Custom field (Application )', 'Labels', 'Assignee']];
        }
      }
      
      // Add issues to our collection
      if (data.issues && data.issues.length > 0) {
        allIssues = allIssues.concat(data.issues);
        console.log(`Fetched ${allIssues.length} of ${total} GTM issues`);
      }
      
      // Extract nextPageToken for subsequent requests
      nextPageToken = data.nextPageToken || null;
      console.log(`GTM Next page token: ${nextPageToken ? nextPageToken : 'None (last page)'}`);
      
      // Update startAt for fallback (in case token-based pagination isn't available)
      startAt += maxResults;
      
      // Continue if we have more pages (either via token or if we haven't reached total)
    } while (nextPageToken || (allIssues.length < total && allIssues.length > 0));
    
    console.log(`=== GTM PAGINATION COMPLETE ===`);
    console.log(`Total GTM API calls made: ${pageCount}`);
    console.log(`Successfully fetched all ${allIssues.length} GTM issues from JIRA`);
    console.log(`GTM Expected: ${total}, Retrieved: ${allIssues.length}, Match: ${allIssues.length === total ? 'YES' : 'NO'}`);
    
    return transformJiraData(allIssues);
    
  } catch (error) {
    console.error('Error fetching GTM JIRA data:', error);
    
    // Enhanced error message for user
    let userFriendlyMessage = 'Failed to fetch GTM data from JIRA.';
    
    if (error.message.includes('Authentication failed')) {
      userFriendlyMessage += ' Please check your JIRA credentials.';
    } else if (error.message.includes('Access forbidden')) {
      userFriendlyMessage += ' Please check your JIRA permissions.';
    } else if (error.message.includes('Filter not found')) {
      userFriendlyMessage += ' Please verify the filter ID is correct.';
    } else if (error.message.includes('API endpoint has been removed')) {
      userFriendlyMessage += ' The JIRA API has been updated. Please contact your administrator.';
    } else {
      userFriendlyMessage += ` Error: ${error.message}`;
    }
    
    throw new Error(userFriendlyMessage);
  }
}

/**
 * Build the GTM-specific JIRA API URL with token-based pagination support
 */
function buildGtmJiraUrl(startAt = 0, maxResults = 100, nextPageToken = null) {
  // GTM Launch Gating specific JQL query
  const gtmJql = `filter=${CONFIG.FILTERS.gtm} AND status NOT IN (CLOSED, Cancelled) AND "Severity SODP" = Critical AND labels NOT IN (VoG_Beta_Launch_Blockers) ORDER BY resolution DESC, priority DESC, cf[19351] asc`;
  
  console.log(`Using GTM filter with JQL: ${gtmJql}`);
  
  // Add fields parameter to get the data we need for the dashboard
  // Include both severity fields: primary (customfield_19351) and fallback (customfield_19632)
  const fields = `summary,status,priority,assignee,labels,created,duedate,${FIELD_MAPPINGS.severity},${FIELD_MAPPINGS.severityAlt},${FIELD_MAPPINGS.application}`;
  
  // Build base URL
  let url = `${CONFIG.JIRA.baseUrl}/rest/api/3/search/jql?jql=${encodeURIComponent(gtmJql)}&fields=${encodeURIComponent(fields)}&maxResults=${maxResults}`;
  
  // Use nextPageToken if available (for subsequent requests), otherwise use startAt (for first request)
  if (nextPageToken) {
    url += `&nextPageToken=${encodeURIComponent(nextPageToken)}`;
    console.log(`Using nextPageToken for GTM pagination: ${nextPageToken}`);
  } else {
    url += `&startAt=${startAt}`;
    console.log(`Using startAt for GTM first request: ${startAt}`);
  }
  
  return url;
}

/**
 * Fetch GTM JIRA data and return as CSV string (called from HTML)
 */
function fetchGtmJiraDataAsCsv() {
  try {
    const csvArray = fetchGtmJiraData();
    return arrayToCsv(csvArray);
  } catch (error) {
    console.error('Error in fetchGtmJiraDataAsCsv:', error);
    throw error;
  }
}

// ===== TESTING STATUS INTEGRATION =====

/**
 * Fetch testing status data from Google Sheets
 */
function fetchTestingStatusData() {
  try {
    console.log('=== Starting fetchTestingStatusData function ===');
    
    // Open the specific spreadsheet by ID from CONFIG
    console.log(`Attempting to open spreadsheet with ID: ${CONFIG.SHEETS.testingStatusId}`);
    var ss = SpreadsheetApp.openById(CONFIG.SHEETS.testingStatusId);
    console.log('Spreadsheet opened successfully: ' + ss.getName());
    
    // Get all sheets and log their info
    var sheets = ss.getSheets();
    console.log('Total sheets found: ' + sheets.length);
    for (var j = 0; j < sheets.length; j++) {
      console.log('Sheet ' + j + ': Name="' + sheets[j].getName() + '", ID=' + sheets[j].getSheetId());
    }
    
    // Get the specific sheet by GID from CONFIG
    var targetSheet = null;
    
    // Find the sheet with the configured GID
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() == CONFIG.SHEETS.testingStatusGid) {
        targetSheet = sheets[i];
        console.log('Found target sheet by GID: ' + targetSheet.getName());
        break;
      }
    }
    
    // If not found by GID, try to get by name or use first sheet
    if (!targetSheet) {
      console.log('Target sheet not found by GID, trying alternatives...');
      targetSheet = ss.getSheetByName('Sheet1') || ss.getSheets()[0];
      console.log('Using alternative sheet: ' + (targetSheet ? targetSheet.getName() : 'NONE FOUND'));
    }
    
    if (!targetSheet) {
      throw new Error('No suitable sheet found in the spreadsheet');
    }
    
    console.log('Using sheet: ' + targetSheet.getName() + ' (ID: ' + targetSheet.getSheetId() + ')');
    
    // Get all data from the sheet - use getDisplayValues() to get actual values, not formulas
    var range = targetSheet.getDataRange();
    var values = range.getDisplayValues();
    
    if (values.length < 2) {
      throw new Error('No data found in the sheet');
    }
    
    // Get headers (first row)
    var headers = values[0];
    console.log('Headers found: ' + JSON.stringify(headers));
    
    // Find column indices
    var columnIndices = {
      teams: findTestingColumnIndex(headers, ['Teams', 'Team', 'team']),
      journey: findTestingColumnIndex(headers, ['Journey', 'journey']),
      totalTCs: findTestingColumnIndex(headers, ['# TC\'s', '# TCs', 'Total TCs', 'Total']),
      passed: findTestingColumnIndex(headers, ['# TC \nPassed', '# TC Passed', 'Passed']),
      failed: findTestingColumnIndex(headers, ['# TC \nFailed', '# TC Failed', 'Failed']),
      blocked: findTestingColumnIndex(headers, ['# TC \nBlocked', '# TC Blocked', 'Blocked']),
      notStarted: findTestingColumnIndex(headers, ['# Not Started and In Progress', 'Not Started', 'In Progress']),
      deferred: findTestingColumnIndex(headers, ['#TC\nDeferred', '#TC Deferred', 'Deferred'])
    };
    
    console.log('Column indices found: ' + JSON.stringify(columnIndices));
    
    // Check if we found the required columns
    if (columnIndices.teams === -1) {
      throw new Error('Teams column not found. Available headers: ' + JSON.stringify(headers));
    }
    
    // Process data rows
    var teamData = [];
    var lastTeamName = '';
    
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var teamName = row[columnIndices.teams];
      var journey = (columnIndices.journey !== -1) ? row[columnIndices.journey] : '';
      
      // Handle empty team names that might be Day 2 continuation rows
      if (!teamName || teamName.toString().trim() === '') {
        // Check if this might be a Day 2 row for the previous team
        if (journey && journey.toString().toLowerCase().includes('day 2') && lastTeamName) {
          teamName = lastTeamName;
          console.log('Row ' + i + ' - Using previous team name "' + teamName + '" for Day 2 row');
        } else {
          continue;
        }
      } else {
        // Update last team name for potential Day 2 rows
        if (!teamName.toString().toLowerCase().includes('total')) {
          lastTeamName = teamName.toString().trim();
        }
      }
      
      // Skip total rows
      if (teamName.toString().toLowerCase().includes('total')) {
        continue;
      }
      
      var displayName = teamName.toString().trim();
      
      // Check if this is a multi-day team that needs to be split
      if (journey && (journey.toString().toLowerCase().includes('day 1') || journey.toString().toLowerCase().includes('day 2'))) {
        var dayInfo = '';
        if (journey.toString().toLowerCase().includes('day 1')) {
          dayInfo = ' Day 1';
        } else if (journey.toString().toLowerCase().includes('day 2')) {
          dayInfo = ' Day 2';
        }
        displayName = teamName.toString().trim() + dayInfo;
      }
      
      // Get numeric values
      var passed = (columnIndices.passed !== -1) ? (parseInt(row[columnIndices.passed]) || 0) : 0;
      var failed = (columnIndices.failed !== -1) ? (parseInt(row[columnIndices.failed]) || 0) : 0;
      var blocked = (columnIndices.blocked !== -1) ? (parseInt(row[columnIndices.blocked]) || 0) : 0;
      var notStarted = (columnIndices.notStarted !== -1) ? (parseInt(row[columnIndices.notStarted]) || 0) : 0;
      var deferred = (columnIndices.deferred !== -1) ? (parseInt(row[columnIndices.deferred]) || 0) : 0;
      var total = (columnIndices.totalTCs !== -1) ? (parseInt(row[columnIndices.totalTCs]) || 0) : (passed + failed + blocked + notStarted + deferred);
      
      if (i <= 15) {
        console.log('Row ' + i + ' - Team: "' + teamName + '", Journey: "' + journey + '", Display Name: "' + displayName + '", Total: ' + total);
      }
      
      // Filter out unwanted teams
      var unwantedTeams = ['QA', 'Digital', 'Tax', 'IFRS', 'Finance', 'CIO Testing'];
      var isUnwanted = false;
      
      for (var j = 0; j < unwantedTeams.length; j++) {
        if (displayName.toLowerCase().includes(unwantedTeams[j].toLowerCase())) {
          isUnwanted = true;
          break;
        }
      }
      
      if (total > 0 && !isUnwanted) {
        teamData.push({
          name: displayName,
          passed: passed,
          failed: failed,
          blocked: blocked,
          notStarted: notStarted,
          deferred: deferred,
          total: total
        });
      }
    }
    
    console.log('Processed team data: ' + JSON.stringify(teamData));
    console.log('Number of teams found: ' + teamData.length);
    
    return {
      teams: teamData,
      lastUpdated: new Date().toLocaleString(),
      debug: {
        totalTeams: teamData.length,
        columnIndices: columnIndices
      }
    };
    
  } catch (error) {
    console.log('Error in fetchTestingStatusData: ' + error.toString());
    return {
      teams: [],
      error: error.toString(),
      lastUpdated: new Date().toLocaleString()
    };
  }
}

/**
 * Fetch testing status data and return as JSON string (called from HTML)
 */
function fetchTestingStatusDataAsJson() {
  try {
    const data = fetchTestingStatusData();
    return JSON.stringify(data);
  } catch (error) {
    console.error('Error in fetchTestingStatusDataAsJson:', error);
    return JSON.stringify({
      teams: [],
      error: error.toString(),
      lastUpdated: new Date().toLocaleString()
    });
  }
}

// ============================================================
// ===== DEBUG FUNCTIONS — Run directly from Apps Script  =====
// ============================================================

/**
 * Helper: POST to JIRA /search/jql with fields=['*all'] and return all issues.
 * Uses filter 35173 so it matches exactly what the Defects tab fetches.
 */
function _debugFetchAllRaw_() {
  console.log('🔍 _debugFetchAllRaw_: Starting POST to JIRA with fields=*all ...');

  const baseOptions = getJiraFetchOptions();
  const authHeader  = baseOptions.headers['Authorization'];
  const searchUrl   = CONFIG.JIRA.baseUrl + '/rest/api/3/search/jql';

  const body = {
    jql:        'filter=' + CONFIG.FILTERS.defects,
    maxResults: 200,
    fields:     ['*all']   // Request EVERY field so nothing is hidden
  };

  console.log('🔍 POST URL  : ' + searchUrl);
  console.log('🔍 POST body : ' + JSON.stringify(body));

  const response = UrlFetchApp.fetch(searchUrl, {
    method:           'POST',
    headers: {
      'Authorization': authHeader,
      'Accept':        'application/json',
      'Content-Type':  'application/json'
    },
    payload:          JSON.stringify(body),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  console.log('🔍 Response code: ' + code);

  if (code !== 200) {
    console.error('❌ JIRA returned ' + code + ': ' + response.getContentText().substring(0, 500));
    return [];
  }

  const data = JSON.parse(response.getContentText());
  console.log('🔍 Total issues returned: ' + (data.issues ? data.issues.length : 0));
  return data.issues || [];
}

/**
 * Helper: log all fields of a single issue in detail.
 */
function _logIssueFields_(issue) {
  const key    = issue.key || '???';
  const fields = issue.fields || {};

  console.log('');
  console.log('══════════════════════════════════════════════════════');
  console.log('📋 ISSUE: ' + key);
  console.log('   Summary   : ' + (fields.summary || '(none)'));
  console.log('   Issue Type: ' + (fields.issuetype ? fields.issuetype.name : '(none)'));
  console.log('   Status    : ' + (fields.status    ? fields.status.name    : '(none)'));
  console.log('   Priority  : ' + (fields.priority  ? fields.priority.name  : '(none)'));
  console.log('   Labels    : ' + JSON.stringify(fields.labels || []));
  console.log('');
  console.log('   ── KEY CUSTOM FIELDS ──────────────────────────────');
  console.log('   customfield_19351 (Severity SODP)          : ' + JSON.stringify(fields['customfield_19351']));
  console.log('   customfield_19632 (Severity of Defect)     : ' + JSON.stringify(fields['customfield_19632']));
  console.log('   customfield_19182 (Application)            : ' + JSON.stringify(fields['customfield_19182']));
  console.log('   customfield_18937 (Application Name Bug)   : ' + JSON.stringify(fields['customfield_18937']));
  console.log('   customfield_19453 (Environments migrated)  : ' + JSON.stringify(fields['customfield_19453']));
  console.log('   customfield_13505 (Environment checkboxes) : ' + JSON.stringify(fields['customfield_13505']));
  console.log('');
  console.log('   ── ALL FIELD KEYS RETURNED BY JIRA ────────────────');
  const allKeys = Object.keys(fields).sort();
  console.log('   Total fields: ' + allKeys.length);
  // Log all customfield_* keys with their values
  allKeys.forEach(function(k) {
    if (k.startsWith('customfield_')) {
      const val = fields[k];
      const valStr = val === null ? 'null' : (typeof val === 'object' ? JSON.stringify(val).substring(0, 120) : String(val));
      console.log('   ' + k + ' = ' + valStr);
    }
  });
  console.log('');
  console.log('   ── FULL RAW JSON (first 3000 chars) ───────────────');
  console.log('   ' + JSON.stringify(issue).substring(0, 3000));
  console.log('══════════════════════════════════════════════════════');
}

/**
 * DEBUG: Fetch all issues from filter 35173 and log detailed raw data
 * for GREEN-* issues only (Experience Defects).
 *
 * Run this directly from the Apps Script editor:
 *   1. Select "fetchGREEN" from the function dropdown
 *   2. Click Run
 *   3. Check the Execution Log for full raw field values
 */
function fetchGREEN() {
  console.log('');
  console.log('╔══════════════════════════════════════════════════════╗');
  console.log('║  fetchGREEN — DEBUG: Raw JIRA data for GREEN issues  ║');
  console.log('╚══════════════════════════════════════════════════════╝');
  console.log('Filter: ' + CONFIG.FILTERS.defects);
  console.log('Fields requested: *all (every field JIRA has)');
  console.log('');

  const issues = _debugFetchAllRaw_();

  if (issues.length === 0) {
    console.log('⚠️  No issues returned from JIRA. Check filter ID and credentials.');
    return;
  }

  const greenIssues = issues.filter(function(i) {
    return i.key && i.key.toUpperCase().startsWith('GREEN-');
  });

  console.log('Total issues in filter : ' + issues.length);
  console.log('GREEN issues found     : ' + greenIssues.length);

  if (greenIssues.length === 0) {
    console.log('⚠️  No GREEN-* issues found in filter ' + CONFIG.FILTERS.defects + '.');
    console.log('    All issue keys: ' + issues.map(function(i) { return i.key; }).join(', '));
    return;
  }

  console.log('GREEN issue keys: ' + greenIssues.map(function(i) { return i.key; }).join(', '));

  greenIssues.forEach(function(issue) {
    _logIssueFields_(issue);
  });

  console.log('');
  console.log('✅ fetchGREEN complete. ' + greenIssues.length + ' GREEN issue(s) logged above.');
}

/**
 * DEBUG: Fetch all issues from filter 35173 and log detailed raw data
 * for NON-GREEN issues only (Bugs, SODP-*, etc.).
 *
 * Run this directly from the Apps Script editor:
 *   1. Select "fetchNONGREEN" from the function dropdown
 *   2. Click Run
 *   3. Check the Execution Log for full raw field values
 */
function fetchNONGREEN() {
  console.log('');
  console.log('╔══════════════════════════════════════════════════════════╗');
  console.log('║  fetchNONGREEN — DEBUG: Raw JIRA data for non-GREEN issues ║');
  console.log('╚══════════════════════════════════════════════════════════╝');
  console.log('Filter: ' + CONFIG.FILTERS.defects);
  console.log('Fields requested: *all (every field JIRA has)');
  console.log('');

  const issues = _debugFetchAllRaw_();

  if (issues.length === 0) {
    console.log('⚠️  No issues returned from JIRA. Check filter ID and credentials.');
    return;
  }

  const nonGreenIssues = issues.filter(function(i) {
    return !(i.key && i.key.toUpperCase().startsWith('GREEN-'));
  });

  console.log('Total issues in filter    : ' + issues.length);
  console.log('Non-GREEN issues found    : ' + nonGreenIssues.length);

  if (nonGreenIssues.length === 0) {
    console.log('⚠️  All issues in filter are GREEN-* issues.');
    return;
  }

  console.log('Non-GREEN issue keys: ' + nonGreenIssues.map(function(i) { return i.key; }).join(', '));

  nonGreenIssues.forEach(function(issue) {
    _logIssueFields_(issue);
  });

  console.log('');
  console.log('✅ fetchNONGREEN complete. ' + nonGreenIssues.length + ' non-GREEN issue(s) logged above.');
}
