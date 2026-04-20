  // ===== EPICS-SPECIFIC FUNCTIONS =====
  // This file contains all epics-related data fetching and transformation logic
  // Configuration and shared utilities are in Core.gs

  /**
  * Fetch Epics data from JIRA using the configured filter with token-based pagination
  * ENHANCED: Also fetches all Stories whose parent is one of the fetched Epics
  */
  function fetchEpicsData() {
    try {
      console.log('=== STARTING EPICS & STORIES JIRA DATA FETCH WITH TOKEN-BASED PAGINATION ===');
      
      // Validate JIRA connection using shared function from Core.gs
      const validationResult = validateJiraConnection();
      if (!validationResult.success) {
        throw new Error(`JIRA connection validation failed: ${validationResult.error}`);
      }

      // PHASE 1: Fetch all Epics from the filter
      console.log('=== PHASE 1: FETCHING EPICS ===');
      const epics = fetchEpicsFromFilter();
      
      if (epics.length === 0) {
        console.warn('No Epics found in filter');
        return [['Work Type', 'Work Item Key', 'Summary', 'Health', 'Status Details/Risk', 'Application', 'Status', 'Due Date', 'Assignee', 'Parent Epic Key', 'Parent Epic Summary']];
      }
      
      console.log(`Successfully fetched ${epics.length} Epics`);
      
      // PHASE 2: Extract Epic keys and create Epic map for quick lookup
      console.log('=== PHASE 2: EXTRACTING EPIC KEYS ===');
      const epicKeys = epics.map(epic => epic.key).filter(key => key);
      const epicsMap = createEpicsMap(epics);
      
      console.log(`Extracted ${epicKeys.length} Epic keys`);
      console.log(`Epic keys: ${epicKeys.join(', ')}`);
      
      // PHASE 3: Fetch all Stories whose parent is one of these Epics
      console.log('=== PHASE 3: FETCHING STORIES ===');
      const stories = fetchStoriesForEpics(epicKeys);
      
      console.log(`Successfully fetched ${stories.length} Stories`);
      
      // PHASE 4: Enrich Stories with parent Epic information
      console.log('=== PHASE 4: ENRICHING STORIES WITH PARENT INFO ===');
      const enrichedStories = enrichStoriesWithParentInfo(stories, epicsMap);
      
      // PHASE 5: Combine Epics and Stories
      console.log('=== PHASE 5: COMBINING EPICS AND STORIES ===');
      const allIssues = [...epics, ...enrichedStories];
      console.log(`Total issues to transform: ${allIssues.length} (${epics.length} Epics + ${enrichedStories.length} Stories)`);
      
      return transformEpicsData(allIssues);

    } catch (error) {
      console.error('Error fetching Epics & Stories JIRA data:', error);
      
      let userFriendlyMessage = 'Failed to fetch Epics & Stories data from JIRA.';
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
  * Fetch Epics from the configured filter
  * Returns array of Epic issue objects
  */
  function fetchEpicsFromFilter() {
    const maxResults = API_SETTINGS.maxResults;
    let allIssues = [];
    let startAt = 0;
    let nextPageToken = null;
    let total = 0;
    let pageCount = 0;

  // Epics-specific fields using FIELD_MAPPINGS from Core.gs
  const fields = [
    'issuetype',                    // Work Type (JIRA Issue Type)
    'key',                          // Work Item Key
    'summary',                      // Summary
    FIELD_MAPPINGS.health,          // Health (customfield_19357)
    FIELD_MAPPINGS.statusDetails,   // Status Details/Risk (customfield_19361)
    FIELD_MAPPINGS.application,     // Application (customfield_19182)
    'status',                       // Status
    FIELD_MAPPINGS.startDate,       // Start Date (customfield_21750)
    'duedate',                      // Due Date
    'assignee',                     // Assignee
    'labels'                        // Labels
  ].join(',');

    const options = getJiraFetchOptions();

    do {
      pageCount++;
      const url = buildEpicsJiraUrl(startAt, maxResults, nextPageToken, fields);
      console.log(`=== EPICS PAGINATION CALL ${pageCount} ===`);
      console.log(`Fetching Epics JIRA data from: ${url}`);
      
      let response;
      let success = false;

      // Retry logic with exponential backoff
      for (let attempt = 0; attempt <= API_SETTINGS.maxRetries; attempt++) {
        try {
          response = UrlFetchApp.fetch(url, options);
          const responseCode = response.getResponseCode();
          
          console.log(`Epics Attempt ${attempt + 1}: Response code ${responseCode}`);

          if (responseCode === 200) {
            success = true;
            break;
          } else if (responseCode === 401) {
            throw new Error('Authentication failed. Please check your JIRA email and API token in CONFIG.JIRA.');
          } else if (responseCode === 403) {
            throw new Error('Access forbidden. Please check if your account has permission to access the specified filter.');
          } else if (responseCode === 404) {
            throw new Error(`Filter not found. Please verify that filter ID ${CONFIG.FILTERS.epics} exists and is accessible.`);
          } else if (responseCode === 410) {
            throw new Error('JIRA API endpoint has been removed. Please check the JIRA API documentation for the correct endpoint.');
          } else if (responseCode >= 500) {
            if (attempt < API_SETTINGS.maxRetries) {
              const delay = Math.pow(2, attempt) * API_SETTINGS.retryDelayBase;
              console.log(`Epics Server error ${responseCode}, retrying in ${delay}ms...`);
              Utilities.sleep(delay);
              continue;
            } else {
              throw new Error(`JIRA server error ${responseCode}: ${response.getContentText()}`);
            }
          } else {
            const errorDetails = response.getContentText();
            console.error(`Epics JIRA API Error ${responseCode}:`, errorDetails);
            throw new Error(`JIRA API returned ${responseCode}: ${errorDetails}`);
          }
        } catch (error) {
          if (attempt === API_SETTINGS.maxRetries) {
            throw error;
          }
          console.log(`Epics Attempt ${attempt + 1} failed: ${error.message}`);
          const delay = Math.pow(2, attempt) * API_SETTINGS.retryDelayBase;
          Utilities.sleep(delay);
        }
      }

      if (!success) {
        throw new Error('Failed to fetch Epics data after maximum retries');
      }

      const data = JSON.parse(response.getContentText());
      
      // Log response structure for debugging
      console.log(`Epics Response keys: ${Object.keys(data).join(', ')}`);
      console.log(`Epics Issues returned: ${data.issues ? data.issues.length : 0}`);

      if (pageCount === 1) {
        total = data.total;
        console.log(`Total Epics issues in filter: ${total}`);
      }

      if (data.issues && data.issues.length > 0) {
        allIssues = allIssues.concat(data.issues);
        console.log(`Fetched ${allIssues.length} of ${total} Epics issues`);
      }

      nextPageToken = data.nextPageToken || null;
      console.log(`Epics Next page token: ${nextPageToken ? nextPageToken : 'None (last page)'}`);
      
      startAt += maxResults;
    } while (nextPageToken || (allIssues.length < total && allIssues.length > 0));

    console.log(`=== EPICS PAGINATION COMPLETE ===`);
    console.log(`Total Epics API calls made: ${pageCount}`);
    console.log(`Successfully fetched all ${allIssues.length} Epics issues from JIRA`);
    console.log(`Epics Expected: ${total}, Retrieved: ${allIssues.length}, Match: ${allIssues.length === total ? 'YES' : 'NO'}`);

    return allIssues;
  }

  /**
  * Create a map of Epic key -> Epic data for quick lookup
  */
  function createEpicsMap(epics) {
    const epicsMap = {};
    epics.forEach(epic => {
      if (epic.key) {
        epicsMap[epic.key] = {
          key: epic.key,
          summary: epic.fields.summary || ''
        };
      }
    });
    return epicsMap;
  }

  /**
  * Fetch all Stories/child issues whose parent is one of the provided Epic keys.
  * Uses POST /rest/api/3/search/jql to avoid URL length limits when there are many Epic keys.
  */
  function fetchStoriesForEpics(epicKeys) {
    if (!epicKeys || epicKeys.length === 0) {
      console.log('No Epic keys provided, skipping Stories fetch');
      return [];
    }

    const maxResults = API_SETTINGS.maxResults;
    let allStories = [];
    let nextPageToken = null;
    let pageCount = 0;

    // Build JQL query: Parent in (KEY1, KEY2, ...) AND status != Cancelled ORDER BY created DESC
    const jql = buildParentInJQL(epicKeys);

    // Stories need same fields as Epics, plus 'parent' field
    const fields = [
      'key',
      'summary',
      'issuetype',
      'status',
      'assignee',
      'duedate',
      'parent',
      'labels',
      FIELD_MAPPINGS.application,
      FIELD_MAPPINGS.health,
      FIELD_MAPPINGS.statusDetails,
      FIELD_MAPPINGS.startDate
    ];

    // Use POST to avoid "URLFetch URL Length" limit when many Epic keys are in the JQL
    // NOTE: POST /rest/api/3/search/jql only supports nextPageToken for pagination (no startAt)
    const postUrl = `${CONFIG.JIRA.baseUrl}/rest/api/3/search/jql`;
    const secret = getJiraSecretKey();
    const headers = {
      'Authorization': 'Bearer ' + secret,
      'Accept': 'application/json',
      'Content-Type': 'application/json',
      'User-Agent': API_SETTINGS.userAgent
    };

    do {
      pageCount++;
      console.log(`=== STORIES PAGINATION CALL ${pageCount} ===`);

      // Build POST body — only nextPageToken for pagination (POST endpoint does NOT support startAt)
      const body = {
        jql: jql,
        fields: fields,
        maxResults: maxResults
      };
      if (nextPageToken) {
        body.nextPageToken = nextPageToken;
        console.log(`Using nextPageToken for Stories pagination: ${nextPageToken}`);
      } else {
        console.log(`First Stories page — no pagination token`);
      }

      const options = {
        method: 'POST',
        headers: headers,
        payload: JSON.stringify(body),
        muteHttpExceptions: true
      };

      console.log(`Fetching Stories via POST to: ${postUrl}`);
      console.log(`Stories JQL: ${jql}`);

      let response;
      let success = false;

      // Retry logic with exponential backoff
      for (let attempt = 0; attempt <= API_SETTINGS.maxRetries; attempt++) {
        try {
          response = UrlFetchApp.fetch(postUrl, options);
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
      console.log(`Stories Total: ${data.total}`);

      if (data.issues && data.issues.length > 0) {
        allStories = allStories.concat(data.issues);
        console.log(`Fetched ${allStories.length} Stories so far`);
      }

      // POST /search/jql uses nextPageToken exclusively for pagination
      nextPageToken = data.nextPageToken || null;
      console.log(`Stories Next page token: ${nextPageToken ? nextPageToken : 'None (last page)'}`);

    } while (nextPageToken !== null);

    console.log(`=== STORIES PAGINATION COMPLETE ===`);
    console.log(`Total Stories API calls made: ${pageCount}`);
    console.log(`Successfully fetched all ${allStories.length} Stories from JIRA`);

    return allStories;
  }

  /**
  * Build JQL query for child issues with parent in Epic keys
  */
  function buildParentInJQL(epicKeys) {
    // Build: Parent in (KEY1, KEY2, ...) AND status != Cancelled ORDER BY created DESC
    const parentInClause = epicKeys.join(', ');
    return `Parent in (${parentInClause}) AND status != Cancelled ORDER BY created DESC`;
  }

  /**
  * Enrich Stories with parent Epic information
  */
  function enrichStoriesWithParentInfo(stories, epicsMap) {
    return stories.map(story => {
      // Add parent Epic info to the story's fields
      if (story.fields && story.fields.parent && story.fields.parent.key) {
        const parentKey = story.fields.parent.key;
        const parentEpic = epicsMap[parentKey];
        
        if (parentEpic) {
          story.fields.parentEpicKey = parentEpic.key;
          story.fields.parentEpicSummary = parentEpic.summary;
        } else {
          console.warn(`Parent Epic ${parentKey} not found in epicsMap for Story ${story.key}`);
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
  * Build the JIRA API URL for Epics with token-based pagination support
  */
  function buildEpicsJiraUrl(startAt, maxResults, nextPageToken, fields) {
    // Use the epics filter from CONFIG
    const jql = `filter=${CONFIG.FILTERS.epics}`;
    
    console.log(`Using Epics filter ${CONFIG.FILTERS.epics} with JQL: ${jql}`);
    
    let url = `${CONFIG.JIRA.baseUrl}/rest/api/3/search/jql?jql=${encodeURIComponent(jql)}&fields=${encodeURIComponent(fields)}&maxResults=${maxResults}`;
    
    // Use nextPageToken if available (for subsequent requests), otherwise use startAt (for first request)
    if (nextPageToken) {
      url += `&nextPageToken=${encodeURIComponent(nextPageToken)}`;
      console.log(`Using nextPageToken for Epics pagination: ${nextPageToken}`);
    } else {
      url += `&startAt=${startAt}`;
      console.log(`Using startAt for Epics first request: ${startAt}`);
    }
    
    return url;
  }

  /**
  * Transform Epics & Stories JIRA data to CSV-like array
  * Includes "Status Details/Risk" (customfield_19361) after "Health" column
  * ENHANCED: Includes Parent Epic Key and Parent Epic Summary columns for Stories
  */
  function transformEpicsData(issues) {
    try {
      console.log('=== STARTING TRANSFORM EPICS & STORIES DATA ===');
      
      if (!issues || !Array.isArray(issues) || issues.length === 0) {
        console.warn('Epics & Stories issues parameter is empty or invalid');
        return [['Work Type', 'Work Item Key', 'Summary', 'Health', 'Status Details/Risk', 'Application', 'Status', 'Start Date', 'Due Date', 'Assignee', 'Parent Epic Key', 'Parent Epic Summary']];
      }

      console.log(`Processing ${issues.length} issues (Epics + Stories)`);

      const csvData = [];
      csvData.push(['Work Type', 'Work Item Key', 'Summary', 'Health', 'Status Details/Risk', 'Application', 'Status', 'Start Date', 'Due Date', 'Assignee', 'Labels', 'Parent Epic Key', 'Parent Epic Summary']);

      // Use field mappings from Core.gs
      const healthField = FIELD_MAPPINGS.health;
      const statusDetailsField = FIELD_MAPPINGS.statusDetails;
      const applicationField = FIELD_MAPPINGS.application;
      const startDateField = FIELD_MAPPINGS.startDate;

      issues.forEach(issue => {
        if (!issue || !issue.fields) {
          console.warn('Skipping malformed Epics issue:', JSON.stringify(issue));
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
      console.error('Critical error in transformEpicsData:', error);
      console.error('Error stack:', error.stack);
      
      return [
        ['Work Type', 'Work Item Key', 'Summary', 'Health', 'Status Details/Risk', 'Application', 'Status', 'Due Date', 'Assignee', 'Parent Epic Key', 'Parent Epic Summary'],
        ['ERROR', `Transform failed: ${error.message}`, '', '', '', '', '', '', '', '', '']
      ];
    }
  }

  /**
  * Return Epics data as CSV string (called from HTML)
  */
  function fetchEpicsDataAsCsv() {
    try {
      const csvArray = fetchEpicsData();
      return arrayToCsv(csvArray); // arrayToCsv is a shared function in Core.gs
    } catch (error) {
      console.error('Error in fetchEpicsDataAsCsv:', error);
      throw error;
    }
  }
