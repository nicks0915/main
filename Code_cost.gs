// ===== ASSESSMENTS-SPECIFIC FUNCTIONS =====
// This file contains all assessments-related data fetching and transformation logic
// Configuration and shared utilities are in Core.gs

/**
 * Fetch Assessments data from JIRA using the configured filter with token-based pagination
 * ENHANCED: Also fetches all Stories whose parent is one of the fetched Assessments
 */
function fetchAssessmentsData() {
  try {
    console.log('=== STARTING ASSESSMENTS & STORIES JIRA DATA FETCH WITH TOKEN-BASED PAGINATION ===');
    
    // Validate JIRA connection using shared function from Core.gs
    const validationResult = validateJiraConnection();
    if (!validationResult.success) {
      throw new Error(`JIRA connection validation failed: ${validationResult.error}`);
    }

    // PHASE 1: Fetch all Assessments from the filter
    console.log('=== PHASE 1: FETCHING ASSESSMENTS ===');
    const assessments = fetchAssessmentsFromFilter();
    
    if (assessments.length === 0) {
      console.warn('No Assessments found in filter');
      return [['Work Type', 'Work Item Key', 'Summary', 'Health', 'Status Details/Risk', 'Application', 'Gate 2 - Cost Estimate', 'Status', 'Due Date', 'Assignee', 'Parent Epic Key', 'Parent Epic Summary']];
    }
    
    console.log(`Successfully fetched ${assessments.length} Assessments`);
    
    // PHASE 2: Extract Assessment keys and create Assessment map for quick lookup
    console.log('=== PHASE 2: EXTRACTING ASSESSMENT KEYS ===');
    const assessmentKeys = assessments.map(assessment => assessment.key).filter(key => key);
    const assessmentsMap = createAssessmentsMap(assessments);
    
    console.log(`Extracted ${assessmentKeys.length} Assessment keys`);
    console.log(`Assessment keys: ${assessmentKeys.join(', ')}`);
    
    // PHASE 3: Fetch all Stories whose parent is one of these Assessments
    console.log('=== PHASE 3: FETCHING STORIES ===');
    const stories = fetchStoriesForAssessments(assessmentKeys);
    
    console.log(`Successfully fetched ${stories.length} Stories`);
    
    // PHASE 4: Enrich Stories with parent Assessment information
    console.log('=== PHASE 4: ENRICHING STORIES WITH PARENT INFO ===');
    const enrichedStories = enrichStoriesWithParentInfoAssessments(stories, assessmentsMap);
    
    // PHASE 5: Combine Assessments and Stories
    console.log('=== PHASE 5: COMBINING ASSESSMENTS AND STORIES ===');
    const allIssues = [...assessments, ...enrichedStories];
    console.log(`Total issues to transform: ${allIssues.length} (${assessments.length} Assessments + ${enrichedStories.length} Stories)`);
    
    return transformAssessmentsData(allIssues);

  } catch (error) {
    console.error('Error fetching Assessments & Stories JIRA data:', error);
    
    let userFriendlyMessage = 'Failed to fetch Assessments & Stories data from JIRA.';
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
 * Fetch Assessments from the configured filter
 * Returns array of Assessment issue objects
 */
function fetchAssessmentsFromFilter() {
  const maxResults = API_SETTINGS.maxResults;
  let allIssues = [];
  let startAt = 0;
  let nextPageToken = null;
  let total = 0;
  let pageCount = 0;

  // Assessments-specific fields using FIELD_MAPPINGS from Core.gs
  const fields = [
    'issuetype',                    // Work Type (JIRA Issue Type)
    'key',                          // Work Item Key
    'summary',                      // Summary
    FIELD_MAPPINGS.health,          // Health (customfield_19357)
    FIELD_MAPPINGS.statusDetails,   // Status Details/Risk (customfield_19361)
    FIELD_MAPPINGS.application,     // Application (customfield_19182)
    FIELD_MAPPINGS.costEstimate,    // Gate 2 - Cost Estimate (customfield_17265)
    'status',                       // Status
    FIELD_MAPPINGS.startDate,       // Start Date (customfield_21750)
    'duedate',                      // Due Date
    'assignee',                     // Assignee
    'labels'                        // Labels
  ].join(',');

  const options = getJiraFetchOptions();

  do {
    pageCount++;
    const url = buildAssessmentsJiraUrl(startAt, maxResults, nextPageToken, fields);
    console.log(`=== ASSESSMENTS PAGINATION CALL ${pageCount} ===`);
    console.log(`Fetching Assessments JIRA data from: ${url}`);
    
    let response;
    let success = false;

    // Retry logic with exponential backoff
    for (let attempt = 0; attempt <= API_SETTINGS.maxRetries; attempt++) {
      try {
        response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        
        console.log(`Assessments Attempt ${attempt + 1}: Response code ${responseCode}`);

        if (responseCode === 200) {
          success = true;
          break;
        } else if (responseCode === 401) {
          throw new Error('Authentication failed. Please check your JIRA email and API token in CONFIG.JIRA.');
        } else if (responseCode === 403) {
          throw new Error('Access forbidden. Please check if your account has permission to access the specified filter.');
        } else if (responseCode === 404) {
          throw new Error(`Filter not found. Please verify that filter ID ${CONFIG.FILTERS.assessments} exists and is accessible.`);
        } else if (responseCode === 410) {
          throw new Error('JIRA API endpoint has been removed. Please check the JIRA API documentation for the correct endpoint.');
        } else if (responseCode >= 500) {
          if (attempt < API_SETTINGS.maxRetries) {
            const delay = Math.pow(2, attempt) * API_SETTINGS.retryDelayBase;
            console.log(`Assessments Server error ${responseCode}, retrying in ${delay}ms...`);
            Utilities.sleep(delay);
            continue;
          } else {
            throw new Error(`JIRA server error ${responseCode}: ${response.getContentText()}`);
          }
        } else {
          const errorDetails = response.getContentText();
          console.error(`Assessments JIRA API Error ${responseCode}:`, errorDetails);
          throw new Error(`JIRA API returned ${responseCode}: ${errorDetails}`);
        }
      } catch (error) {
        if (attempt === API_SETTINGS.maxRetries) {
          throw error;
        }
        console.log(`Assessments Attempt ${attempt + 1} failed: ${error.message}`);
        const delay = Math.pow(2, attempt) * API_SETTINGS.retryDelayBase;
        Utilities.sleep(delay);
      }
    }

    if (!success) {
      throw new Error('Failed to fetch Assessments data after maximum retries');
    }

    const data = JSON.parse(response.getContentText());
    
    // Log response structure for debugging
    console.log(`Assessments Response keys: ${Object.keys(data).join(', ')}`);
    console.log(`Assessments Issues returned: ${data.issues ? data.issues.length : 0}`);

    if (pageCount === 1) {
      total = data.total;
      console.log(`Total Assessments issues in filter: ${total}`);
    }

    if (data.issues && data.issues.length > 0) {
      allIssues = allIssues.concat(data.issues);
      console.log(`Fetched ${allIssues.length} of ${total} Assessments issues`);
    }

    nextPageToken = data.nextPageToken || null;
    console.log(`Assessments Next page token: ${nextPageToken ? nextPageToken : 'None (last page)'}`);
    
    startAt += maxResults;
  } while (nextPageToken || (allIssues.length < total && allIssues.length > 0));

  console.log(`=== ASSESSMENTS PAGINATION COMPLETE ===`);
  console.log(`Total Assessments API calls made: ${pageCount}`);
  console.log(`Successfully fetched all ${allIssues.length} Assessments issues from JIRA`);
  console.log(`Assessments Expected: ${total}, Retrieved: ${allIssues.length}, Match: ${allIssues.length === total ? 'YES' : 'NO'}`);

  return allIssues;
}

/**
 * Create a map of Assessment key -> Assessment data for quick lookup
 */
function createAssessmentsMap(assessments) {
  const assessmentsMap = {};
  assessments.forEach(assessment => {
    if (assessment.key) {
      assessmentsMap[assessment.key] = {
        key: assessment.key,
        summary: assessment.fields.summary || ''
      };
    }
  });
  return assessmentsMap;
}

/**
 * Fetch all Stories whose parent is one of the provided Assessment keys
 */
function fetchStoriesForAssessments(assessmentKeys) {
  if (!assessmentKeys || assessmentKeys.length === 0) {
    console.log('No Assessment keys provided, skipping Stories fetch');
    return [];
  }

  const maxResults = API_SETTINGS.maxResults;
  let allStories = [];
  let startAt = 0;
  let nextPageToken = null;
  let total = 0;
  let pageCount = 0;

  // Build JQL query: parent in (KEY1, KEY2, ...) AND type = Story
  const jql = buildParentInJQLAssessments(assessmentKeys);
  
  // Stories need same fields as Assessments, plus 'parent' field
  const fields = [
    'issuetype',
    'key',
    'summary',
    FIELD_MAPPINGS.health,
    FIELD_MAPPINGS.statusDetails,
    FIELD_MAPPINGS.application,
    FIELD_MAPPINGS.costEstimate,
    'status',
    FIELD_MAPPINGS.startDate,
    'duedate',
    'assignee',
    'labels',
    'parent'  // To get parent Assessment key
  ].join(',');

  const options = getJiraFetchOptions();

  do {
    pageCount++;
    const url = buildStoriesJiraUrlAssessments(jql, startAt, maxResults, nextPageToken, fields);
    console.log(`=== STORIES PAGINATION CALL ${pageCount} ===`);
    console.log(`Fetching Stories JIRA data from: ${url}`);
    
    let response;
    let success = false;

    // Retry logic with exponential backoff
    for (let attempt = 0; attempt <= API_SETTINGS.maxRetries; attempt++) {
      try {
        response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        
        console.log(`Stories Attempt ${attempt + 1}: Response code ${responseCode}`);

        if (responseCode === 200) {
          success = true;
          break;
        } else if (responseCode >= 500) {
          if (attempt < API_SETTINGS.maxRetries) {
            const delay = Math.pow(2, attempt) * API_SETTINGS.retryDelayBase;
            console.log(`Stories Server error ${responseCode}, retrying in ${delay}ms...`);
            Utilities.sleep(delay);
            continue;
          } else {
            throw new Error(`JIRA server error ${responseCode}: ${response.getContentText()}`);
          }
        } else {
          const errorDetails = response.getContentText();
          console.error(`Stories JIRA API Error ${responseCode}:`, errorDetails);
          throw new Error(`JIRA API returned ${responseCode}: ${errorDetails}`);
        }
      } catch (error) {
        if (attempt === API_SETTINGS.maxRetries) {
          throw error;
        }
        console.log(`Stories Attempt ${attempt + 1} failed: ${error.message}`);
        const delay = Math.pow(2, attempt) * API_SETTINGS.retryDelayBase;
        Utilities.sleep(delay);
      }
    }

    if (!success) {
      throw new Error('Failed to fetch Stories data after maximum retries');
    }

    const data = JSON.parse(response.getContentText());
    
    console.log(`Stories Response keys: ${Object.keys(data).join(', ')}`);
    console.log(`Stories Issues returned: ${data.issues ? data.issues.length : 0}`);

    if (pageCount === 1) {
      total = data.total;
      console.log(`Total Stories issues: ${total}`);
    }

    if (data.issues && data.issues.length > 0) {
      allStories = allStories.concat(data.issues);
      console.log(`Fetched ${allStories.length} of ${total} Stories issues`);
    }

    nextPageToken = data.nextPageToken || null;
    console.log(`Stories Next page token: ${nextPageToken ? nextPageToken : 'None (last page)'}`);
    
    startAt += maxResults;
  } while (nextPageToken || (allStories.length < total && allStories.length > 0));

  console.log(`=== STORIES PAGINATION COMPLETE ===`);
  console.log(`Total Stories API calls made: ${pageCount}`);
  console.log(`Successfully fetched all ${allStories.length} Stories from JIRA`);

  return allStories;
}

/**
 * Build JQL query for Stories with parent in Assessment keys
 */
function buildParentInJQLAssessments(assessmentKeys) {
  // Build: project = SODP AND type = Story AND Parent in (KEY1, KEY2, ...) AND status != Cancelled ORDER BY created DESC
  const parentInClause = assessmentKeys.join(', ');
  return `project = SODP AND type = Story AND Parent in (${parentInClause}) AND status != Cancelled ORDER BY created DESC`;
}

/**
 * Build JIRA API URL for Stories query (Assessments)
 */
function buildStoriesJiraUrlAssessments(jql, startAt, maxResults, nextPageToken, fields) {
  console.log(`Using Stories JQL: ${jql}`);
  
  let url = `${CONFIG.JIRA.baseUrl}/rest/api/3/search/jql?jql=${encodeURIComponent(jql)}&fields=${encodeURIComponent(fields)}&maxResults=${maxResults}`;
  
  if (nextPageToken) {
    url += `&nextPageToken=${encodeURIComponent(nextPageToken)}`;
    console.log(`Using nextPageToken for Stories pagination: ${nextPageToken}`);
  } else {
    url += `&startAt=${startAt}`;
    console.log(`Using startAt for Stories first request: ${startAt}`);
  }
  
  return url;
}

/**
 * Enrich Stories with parent Assessment information
 */
function enrichStoriesWithParentInfoAssessments(stories, assessmentsMap) {
  return stories.map(story => {
    // Add parent Assessment info to the story's fields
    if (story.fields && story.fields.parent && story.fields.parent.key) {
      const parentKey = story.fields.parent.key;
      const parentAssessment = assessmentsMap[parentKey];
      
      if (parentAssessment) {
        story.fields.parentEpicKey = parentAssessment.key;
        story.fields.parentEpicSummary = parentAssessment.summary;
      } else {
        console.warn(`Parent Assessment ${parentKey} not found in assessmentsMap for Story ${story.key}`);
        story.fields.parentEpicKey = parentKey;
        story.fields.parentEpicSummary = '';
      }
    } else {
      // Story has no parent (shouldn't happen based on our query, but handle gracefully)
      story.fields.parentEpicKey = '';
      story.fields.parentEpicSummary = '';
    }
    
    return story;
  });
}

/**
 * Build the JIRA API URL for Assessments with token-based pagination support
 */
function buildAssessmentsJiraUrl(startAt, maxResults, nextPageToken, fields) {
  // Use the assessments filter from CONFIG
  const jql = `filter=${CONFIG.FILTERS.assessments}`;
  
  console.log(`Using Assessments filter ${CONFIG.FILTERS.assessments} with JQL: ${jql}`);
  
  let url = `${CONFIG.JIRA.baseUrl}/rest/api/3/search/jql?jql=${encodeURIComponent(jql)}&fields=${encodeURIComponent(fields)}&maxResults=${maxResults}`;
  
  // Use nextPageToken if available (for subsequent requests), otherwise use startAt (for first request)
  if (nextPageToken) {
    url += `&nextPageToken=${encodeURIComponent(nextPageToken)}`;
    console.log(`Using nextPageToken for Assessments pagination: ${nextPageToken}`);
  } else {
    url += `&startAt=${startAt}`;
    console.log(`Using startAt for Assessments first request: ${startAt}`);
  }
  
  return url;
}

/**
 * Transform Assessments & Stories JIRA data to CSV-like array
 * Includes "Status Details/Risk" (customfield_19361) after "Health" column
 * ENHANCED: Includes Parent Epic Key and Parent Epic Summary columns for Stories
 */
function transformAssessmentsData(issues) {
  try {
    console.log('=== STARTING TRANSFORM ASSESSMENTS & STORIES DATA ===');
    
    if (!issues || !Array.isArray(issues) || issues.length === 0) {
      console.warn('Assessments & Stories issues parameter is empty or invalid');
      return [['Work Type', 'Work Item Key', 'Summary', 'Health', 'Status Details/Risk', 'Application', 'Gate 2 - Cost Estimate', 'Status', 'Start Date', 'Due Date', 'Assignee', 'Parent Epic Key', 'Parent Epic Summary']];
    }

    console.log(`Processing ${issues.length} issues (Assessments + Stories)`);

    const csvData = [];
    csvData.push(['Work Type', 'Work Item Key', 'Summary', 'Health', 'Status Details/Risk', 'Application', 'Gate 2 - Cost Estimate', 'Status', 'Start Date', 'Due Date', 'Assignee', 'Labels', 'Parent Epic Key', 'Parent Epic Summary']);

    // Use field mappings from Core.gs
    const healthField = FIELD_MAPPINGS.health;
    const statusDetailsField = FIELD_MAPPINGS.statusDetails;
    const applicationField = FIELD_MAPPINGS.application;
    const startDateField = FIELD_MAPPINGS.startDate;

    issues.forEach(issue => {
      if (!issue || !issue.fields) {
        console.warn('Skipping malformed Assessments issue:', JSON.stringify(issue));
        return;
      }
      
      const fields = issue.fields;
      
      // Health value extraction with error handling
      let healthValue = '';
      try {
        if (fields[healthField]) {
          if (typeof fields[healthField] === 'object' && fields[healthField].id) {
            healthValue = fields[healthField].id;
          } else if (typeof fields[healthField] === 'object' && fields[healthField].value) {
            healthValue = fields[healthField].value;
          } else {
            healthValue = String(fields[healthField]);
          }
        }
      } catch (error) {
        console.warn(`Error extracting health for ${issue.key}:`, error);
      }
      
      // Status Details/Risk extraction (ADF-aware, show full plain text)
      // customfield_19361 comes from JIRA as Atlassian Document Format (ADF) — a nested JSON object.
      // We must recursively walk the ADF tree to extract plain text, NOT use String() which gives "[object Object]".
      let statusDetailsValue = '';
      try {
        if (fields[statusDetailsField]) {
          if (typeof fields[statusDetailsField] === 'string') {
            statusDetailsValue = fields[statusDetailsField];
          } else if (typeof fields[statusDetailsField] === 'object') {
            // Handle Atlassian Document Format (ADF)
            const extractTextFromADF = (node) => {
              let text = '';
              if (node.type === 'text') {
                text += node.text || '';
              }
              if (node.content && Array.isArray(node.content)) {
                node.content.forEach(child => {
                  text += extractTextFromADF(child);
                });
              }
              // Add line break after paragraphs for readability
              if (node.type === 'paragraph' && text) {
                text += '\n';
              }
              return text;
            };
            statusDetailsValue = extractTextFromADF(fields[statusDetailsField]).trim();
            if (!statusDetailsValue) statusDetailsValue = '';
          }
        }
      } catch (error) {
        console.warn(`Error extracting status details for ${issue.key}:`, error);
      }
      
      // Application value extraction
      let applicationValue = '';
      try {
        if (fields[applicationField]) {
          if (typeof fields[applicationField] === 'string') {
            applicationValue = fields[applicationField];
          } else if (fields[applicationField].value) {
            applicationValue = fields[applicationField].value;
          } else {
            applicationValue = String(fields[applicationField]);
          }
        }
      } catch (error) {
        console.warn(`Error extracting application for ${issue.key}:`, error);
      }
      
      // Gate 2 - Cost Estimate extraction
      let costEstimateValue = '';
      try {
        const costEstimateField = FIELD_MAPPINGS.costEstimate;
        if (fields[costEstimateField]) {
          if (typeof fields[costEstimateField] === 'string') {
            costEstimateValue = fields[costEstimateField];
          } else if (typeof fields[costEstimateField] === 'number') {
            costEstimateValue = String(fields[costEstimateField]);
          } else if (fields[costEstimateField].value) {
            costEstimateValue = fields[costEstimateField].value;
          } else {
            costEstimateValue = String(fields[costEstimateField]);
          }
        }
      } catch (error) {
        console.warn(`Error extracting cost estimate for ${issue.key}:`, error);
      }
      
      // Work Type (Issue Type)
      let workType = '';
      try {
        workType = (fields.issuetype && fields.issuetype.name) ? fields.issuetype.name : '';
      } catch (error) {
        console.warn(`Error extracting work type for ${issue.key}:`, error);
      }
      
      // Status
      let status = '';
      try {
        status = (fields.status && fields.status.name) ? fields.status.name : '';
      } catch (error) {
        console.warn(`Error extracting status for ${issue.key}:`, error);
      }
      
      // Assignee
      let assignee = 'Unassigned';
      try {
        assignee = (fields.assignee && fields.assignee.displayName) ? fields.assignee.displayName : 'Unassigned';
      } catch (error) {
        console.warn(`Error extracting assignee for ${issue.key}:`, error);
      }
      
      // Start Date extraction
      let startDateValue = '';
      try {
        startDateValue = fields[startDateField] || '';
      } catch (error) {
        console.warn(`Error extracting start date for ${issue.key}:`, error);
      }
      
      // Parent Epic info (only for Stories)
      const parentEpicKey = fields.parentEpicKey || '';
      const parentEpicSummary = fields.parentEpicSummary || '';
      
      // Labels extraction (array of strings)
      let labelsValue = '';
      try {
        if (fields.labels && Array.isArray(fields.labels)) {
          labelsValue = fields.labels.join(', ');
        }
      } catch (error) {
        console.warn(`Error extracting labels for ${issue.key}:`, error);
      }
      
      csvData.push([
        workType,
        issue.key || '',
        fields.summary || '',
        healthValue,
        statusDetailsValue,
        applicationValue,
        costEstimateValue,
        status,
        startDateValue,
        fields.duedate || '',
        assignee,
        labelsValue,
        parentEpicKey,
        parentEpicSummary
      ]);
    });

    console.log(`Successfully processed ${csvData.length - 1} issues (excluding header)`);
    return csvData;
    
  } catch (error) {
    console.error('Critical error in transformAssessmentsData:', error);
    console.error('Error stack:', error.stack);
    
    return [
      ['Work Type', 'Work Item Key', 'Summary', 'Health', 'Status Details/Risk', 'Application', 'Gate 2 - Cost Estimate', 'Status', 'Due Date', 'Assignee', 'Parent Epic Key', 'Parent Epic Summary'],
      ['ERROR', `Transform failed: ${error.message}`, '', '', '', '', '', '', '', '', '', '']
    ];
  }
}

/**
 * Return Assessments data as CSV string (called from HTML)
 */
function fetchAssessmentsDataAsCsv() {
  try {
    const csvArray = fetchAssessmentsData();
    return arrayToCsv(csvArray); // arrayToCsv is a shared function in Core.gs
  } catch (error) {
    console.error('Error in fetchAssessmentsDataAsCsv:', error);
    throw error;
  }
}

/**
 * Export filtered Assessments data to Google Sheet
 * @param {Array} filteredData - The filtered data array from the frontend (includes headers)
 * @returns {Object} - Success status and message
 */
function exportAssessmentsToSheet(filteredData) {
  try {
    console.log('=== STARTING ASSESSMENTS EXPORT TO GOOGLE SHEET ===');
    
    if (!filteredData || !Array.isArray(filteredData) || filteredData.length === 0) {
      throw new Error('No data provided for export');
    }
    
    console.log(`Received ${filteredData.length} rows to export (including header)`);
    
    // Get the target spreadsheet and sheet
    const spreadsheetId = CONFIG.SHEETS.assessmentsExportId;
    const sheetGid = CONFIG.SHEETS.assessmentsExportGid;
    
    console.log(`Opening spreadsheet: ${spreadsheetId}`);
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    // Get the sheet by GID
    const sheets = spreadsheet.getSheets();
    let targetSheet = null;
    
    for (let sheet of sheets) {
      if (sheet.getSheetId() === sheetGid) {
        targetSheet = sheet;
        break;
      }
    }
    
    if (!targetSheet) {
      // If sheet with GID not found, use the first sheet
      targetSheet = spreadsheet.getSheets()[0];
      console.log(`Sheet with GID ${sheetGid} not found, using first sheet: ${targetSheet.getName()}`);
    } else {
      console.log(`Found target sheet: ${targetSheet.getName()}`);
    }
    
    // Clear existing content
    targetSheet.clear();
    console.log('Cleared existing sheet content');
    
    // Write the data
    const numRows = filteredData.length;
    const numCols = filteredData[0] ? filteredData[0].length : 0;
    
    if (numCols === 0) {
      throw new Error('Invalid data structure: first row has no columns');
    }
    
    console.log(`Writing ${numRows} rows x ${numCols} columns to sheet`);
    const range = targetSheet.getRange(1, 1, numRows, numCols);
    range.setValues(filteredData);
    
    // Format the header row
    const headerRange = targetSheet.getRange(1, 1, 1, numCols);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    
    // Auto-resize columns
    for (let i = 1; i <= numCols; i++) {
      targetSheet.autoResizeColumn(i);
    }
    
    // Freeze the header row
    targetSheet.setFrozenRows(1);
    
    const recordCount = numRows - 1; // Exclude header
    console.log(`Successfully exported ${recordCount} records to Google Sheet`);
    
    return {
      success: true,
      message: `Successfully exported ${recordCount} record${recordCount !== 1 ? 's' : ''} to Google Sheet`,
      recordCount: recordCount
    };
    
  } catch (error) {
    console.error('Error exporting Assessments to Google Sheet:', error);
    console.error('Error stack:', error.stack);
    
    return {
      success: false,
      message: `Export failed: ${error.message}`,
      recordCount: 0
    };
  }
}
