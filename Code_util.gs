// ===== UTILITY FUNCTIONS AND APP INFRASTRUCTURE =====
// This file contains web app entry points, configuration getters, and shared utility functions
// Configuration constants are in Core.gs

// ===== WEB APP ENTRY POINTS =====

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===== CONFIGURATION GETTERS =====

/**
 * Get Feature Flags configuration (called from HTML to enable/disable features).
 * Returns the FEATURE_FLAGS object from Core.gs CONFIG.
 * Each flag: true = enabled (interactive), false = disabled (greyed out, still visible).
 */
function getFeatureFlags() {
  return CONFIG.FEATURE_FLAGS;
}

/**
 * Send summary to Slack with user-selected options.
 * Called from the frontend popup modal.
 * @param {string} summaryText   - The pre-generated summary text shown in the textbox
 * @param {string} summaryType   - 'defects', 'epics', or 'both'
 * @param {Array}  channelIds    - Array of Slack channel IDs selected by the user (multi-select)
 * @returns {object} Success/failure status with message
 */
function sendToSlackWithOptions(summaryText, summaryType, channelIds) {
  try {
    // Normalise channelIds: accept a single string (legacy) or an array
    const channels = Array.isArray(channelIds)
      ? channelIds.filter(id => id && id.trim())
      : (channelIds ? [channelIds] : []);

    console.log(`=== SEND TO SLACK WITH OPTIONS: type=${summaryType}, channels=[${channels.join(', ')}] ===`);

    if (!summaryText || !summaryType || channels.length === 0) {
      return { success: false, message: 'Missing required parameters (summaryText, summaryType, or at least one channelId).' };
    }

    // Helper: send one message (or two for 'both') to a single channel
    function sendToOneChannel(channelId) {
      if (summaryType === 'both') {
        // Split the combined text at the divider line if present
        const divider = '\n' + '='.repeat(60) + '\n';
        const parts = summaryText.split(divider);

        let defectsText = '';
        let epicsText = '';

        if (parts.length >= 2) {
          parts.forEach(part => {
            if (part.trim().toUpperCase().includes('DEFECT')) {
              defectsText = part.trim();
            } else {
              epicsText = part.trim();
            }
          });
        } else {
          epicsText = summaryText;
        }

        const results = [];
        if (defectsText) results.push(sendToSlack(defectsText, 'defects', channelId));
        if (epicsText)   results.push(sendToSlack(epicsText,   'epics',   channelId));
        return results;
      }

      // Single type
      return [sendToSlack(summaryText, summaryType, channelId)];
    }

    // Send to every selected channel and collect results
    const allResults = [];
    channels.forEach(channelId => {
      sendToOneChannel(channelId).forEach(r => allResults.push(r));
    });

    const allSuccess = allResults.every(r => r.success);
    const channelWord = channels.length === 1 ? 'channel' : 'channels';

    return {
      success: allSuccess,
      message: allSuccess
        ? `Summary sent to ${channels.length} ${channelWord} successfully!`
        : `Some messages failed to send to one or more channels. Check logs for details.`
    };

  } catch (error) {
    console.error('sendToSlackWithOptions error:', error);
    return { success: false, message: `Error: ${error.message}` };
  }
}

/**
 * Get Director Team Mapping configuration
 */
function getDirectorTeamMapping() {
  return DIRECTOR_TEAM_MAPPING;
}

/**
 * Get Project Links configuration
 */
function getProjectLinks() {
  return PROJECT_LINKS;
}

/**
 * Get Dashboard Name from configuration
 */
function getDashboardName() {
  return CONFIG.DASHBOARD.name;
}

/**
 * Get JIRA Secret Key from Script Properties (replaces old API Token)
 * @returns {string} JIRA Secret Key
 * @throws {Error} If secret is not configured
 */
function getJiraSecretKey() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const secret = scriptProperties.getProperty('JIRA_SECRET_KEY');
  
  if (!secret) {
    throw new Error('JIRA_SECRET_KEY not found in Script Properties. Please configure it in Project Settings > Script Properties.');
  }
  
  return secret;
}

/**
 * Get Slack Bot Token from Script Properties
 * @returns {string} Slack Bot Token
 * @throws {Error} If token is not configured
 */
function getSlackBotToken() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const token = scriptProperties.getProperty('SLACK_BOT_TOKEN');
  
  if (!token) {
    throw new Error('SLACK_BOT_TOKEN not found in Script Properties. Please configure it in Project Settings > Script Properties.');
  }
  
  return token;
}

// ===== UTILITY FUNCTIONS =====

/**
 * Convert CSV array to CSV string format
 */
function arrayToCsv(csvArray) {
  return csvArray.map(row => 
    row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
  ).join('\n');
}

/**
 * Format JIRA date to match expected format (UTC-based to avoid timezone issues)
 */
function formatJiraDate(jiraDate) {
  if (!jiraDate) return '';
  
  try {
    console.log(`Raw JIRA date: ${jiraDate}`);
    
    // Handle different JIRA date formats
    let dateToProcess = jiraDate;
    
    // If it's just a date string (YYYY-MM-DD), ensure it's treated as UTC
    if (jiraDate.match(/^\d{4}-\d{2}-\d{2}$/)) {
      dateToProcess = jiraDate + 'T00:00:00.000Z';
      console.log(`Added UTC timezone to date: ${dateToProcess}`);
    }
    
    const date = new Date(dateToProcess);
    console.log(`Parsed Date object: ${date.toISOString()}`);
    
    // Use UTC methods to avoid timezone conversion issues
    const utcYear = date.getUTCFullYear();
    const utcMonth = date.getUTCMonth() + 1;
    const utcDay = date.getUTCDate();
    
    console.log(`UTC date components: ${utcYear}-${utcMonth}-${utcDay}`);
    
    // Format as MM/DD/YYYY to match existing CSV format
    const formatted = `${utcMonth.toString().padStart(2, '0')}/${utcDay.toString().padStart(2, '0')}/${utcYear}`;
    console.log(`Formatted result: ${formatted}`);
    
    return formatted;
  } catch (error) {
    console.error('Error formatting date:', jiraDate, error);
    return '';
  }
}

/**
 * Validate JIRA connection and configuration
 * 
 * NOTE: Uses POST /rest/api/3/search/jql instead of GET /rest/api/3/myself
 * because service account Bearer tokens have 'read:jira-work' scope only —
 * they do NOT have 'read:me' scope required by the /myself endpoint.
 * This matches the approach used in the GREEN dashboard (green_core.gs).
 *
 * FALLBACK: If JIRA_SECRET_KEY is not set, also tries JIRA_API_TOKEN
 * (the property name used by the GREEN dashboard) so both dashboards
 * can share the same Script Property without duplication.
 */
function validateJiraConnection() {
  try {
    console.log('=== VALIDATING JIRA CONNECTION ===');
    
    // Check configuration
    if (!CONFIG.JIRA.baseUrl) {
      return {
        success: false,
        error: 'JIRA configuration is incomplete. Please check baseUrl in CONFIG.JIRA.'
      };
    }
    
    // Validate base URL format
    if (!CONFIG.JIRA.baseUrl.startsWith('https://')) {
      return {
        success: false,
        error: 'JIRA baseUrl must start with https://'
      };
    }
    
    // Get Secret Key from Script Properties (with fallback to JIRA_API_TOKEN)
    const secret = getJiraSecretKey();
    
    // Test connectivity using /search/jql (POST) — service account tokens support read:jira-work scope.
    // The /myself endpoint requires read:me scope which service account tokens do NOT have,
    // causing a 401/403 even when the token is perfectly valid.
    //
    // Use a broad JQL that any valid token can execute (not project-specific),
    // so validation succeeds even if the token has limited project access.
    const testUrl = `${CONFIG.JIRA.baseUrl}/rest/api/3/search/jql`;
    const options = {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + secret,
        'Accept': 'application/json',
        'Content-Type': 'application/json',
        'User-Agent': API_SETTINGS.userAgent
      },
      payload: JSON.stringify({ jql: 'ORDER BY created DESC', maxResults: 1, fields: ['key'] }),
      muteHttpExceptions: true
    };
    
    console.log('Testing JIRA connectivity with /search/jql endpoint (service account compatible)...');
    const response = UrlFetchApp.fetch(testUrl, options);
    const responseCode = response.getResponseCode();
    
    console.log(`JIRA validation response code: ${responseCode}`);
    
    if (responseCode === 200) {
      console.log('✅ JIRA connection validated successfully via search/jql endpoint');
      return {
        success: true
      };
    } else if (responseCode === 400) {
      // 400 can mean the JQL is valid but returns no results — still means auth worked
      // Try a simpler fallback JQL
      console.log('Got 400 on broad JQL, trying filter-based validation...');
      const fallbackOptions = {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + secret,
          'Accept': 'application/json',
          'Content-Type': 'application/json',
          'User-Agent': API_SETTINGS.userAgent
        },
        payload: JSON.stringify({ jql: `filter=${CONFIG.FILTERS.defects}`, maxResults: 1, fields: ['key'] }),
        muteHttpExceptions: true
      };
      const fallbackResponse = UrlFetchApp.fetch(testUrl, fallbackOptions);
      const fallbackCode = fallbackResponse.getResponseCode();
      console.log(`Fallback validation response code: ${fallbackCode}`);
      if (fallbackCode === 200) {
        console.log('✅ JIRA connection validated successfully via filter fallback');
        return { success: true };
      } else if (fallbackCode === 401) {
        return { success: false, error: 'Authentication failed. Please verify your JIRA_SECRET_KEY in Script Properties is correct.' };
      } else if (fallbackCode === 403) {
        return { success: false, error: 'Access forbidden. The Service Account may not have sufficient permissions.' };
      } else {
        const errorText = fallbackResponse.getContentText();
        console.error(`JIRA fallback validation failed with code ${fallbackCode}:`, errorText);
        return { success: false, error: `JIRA API returned ${fallbackCode}. Please check your JIRA configuration.` };
      }
    } else if (responseCode === 401) {
      return {
        success: false,
        error: 'Authentication failed. Please verify your JIRA_SECRET_KEY in Script Properties is correct.'
      };
    } else if (responseCode === 403) {
      return {
        success: false,
        error: 'Access forbidden. The Service Account may not have sufficient permissions.'
      };
    } else {
      const errorText = response.getContentText();
      console.error(`JIRA validation failed with code ${responseCode}:`, errorText);
      return {
        success: false,
        error: `JIRA API returned ${responseCode}. Please check your JIRA configuration.`
      };
    }
    
  } catch (error) {
    console.error('JIRA validation error:', error);
    return {
      success: false,
      error: `Connection validation failed: ${error.message}`
    };
  }
}

/**
 * Validate filter access
 */
function validateFilterAccess(filterId) {
  try {
    console.log(`Testing filter access for filter ID: ${filterId}`);
    
    // Get Secret Key from Script Properties
    const secret = getJiraSecretKey();
    
    const filterUrl = `${CONFIG.JIRA.baseUrl}/rest/api/3/filter/${filterId}`;
    const headers = {
      'Authorization': 'Bearer ' + secret,
      'Accept': 'application/json',
      'Content-Type': 'application/json'
    };
    
    const options = {
      method: 'GET',
      headers: headers,
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(filterUrl, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const filterData = JSON.parse(response.getContentText());
      console.log(`✅ Filter access successful. Filter: "${filterData.name}" by ${filterData.owner.displayName}`);
      return {
        success: true,
        filterName: filterData.name,
        filterOwner: filterData.owner.displayName
      };
    } else if (responseCode === 404) {
      return {
        success: false,
        error: `Filter ${filterId} not found. Please verify the filter ID is correct and accessible.`
      };
    } else if (responseCode === 403) {
      return {
        success: false,
        error: `Access denied to filter ${filterId}. Please ensure you have permission to view this filter.`
      };
    } else {
      const errorText = response.getContentText();
      console.error(`Filter validation failed with code ${responseCode}:`, errorText);
      return {
        success: false,
        error: `Filter validation failed with code ${responseCode}. Please check the filter ID.`
      };
    }
    
  } catch (error) {
    console.error('Filter validation error:', error);
    return {
      success: false,
      error: `Filter validation failed: ${error.message}`
    };
  }
}

/**
 * Get standard JIRA API headers (Bearer Token auth)
 */
function getJiraHeaders() {
  const secret = getJiraSecretKey();
  return {
    'Authorization': 'Bearer ' + secret,
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    'User-Agent': API_SETTINGS.userAgent
  };
}

/**
 * Get standard fetch options for JIRA API calls
 */
function getJiraFetchOptions() {
  return {
    method: 'GET',
    headers: getJiraHeaders(),
    muteHttpExceptions: true
  };
}

/**
 * Helper function to find column index by possible header names for testing data
 * Handles exact matching and line breaks in headers
 */
function findTestingColumnIndex(headers, possibleNames) {
  for (var i = 0; i < headers.length; i++) {
    var header = headers[i].toString().trim();
    
    for (var j = 0; j < possibleNames.length; j++) {
      var searchName = possibleNames[j];
      
      // Try exact match first
      if (header === searchName) {
        console.log('Found exact match for "' + searchName + '" at column ' + i);
        return i;
      }
      
      // Try case-insensitive match
      if (header.toLowerCase() === searchName.toLowerCase()) {
        console.log('Found case-insensitive match for "' + searchName + '" at column ' + i);
        return i;
      }
      
      // Try contains match (for partial matches)
      if (header.toLowerCase().includes(searchName.toLowerCase())) {
        console.log('Found partial match for "' + searchName + '" in "' + header + '" at column ' + i);
        return i;
      }
    }
  }
  
  console.log('No match found for any of: ' + JSON.stringify(possibleNames));
  return -1; // Not found
}
