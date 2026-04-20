/**
 * GREEN Dashboard - Google Apps Script Backend
 * JIRA Integration for GREEN/GMAP Projects
 * Two-Dropdown Filter: Idea (GMAP) + Phase Label
 * Author: Cline AI Assistant
 * Date: December 2025
 */

// JIRA Configuration
const JIRA_CONFIG = {
  baseUrl: 'https://telus-cio.atlassian.net',
  gmapProject: 'GMAP',      // Ideas/Initiatives project
  greenProject: 'GREEN',     // Capabilities project
  gssProject: 'GSS'          // Green Sample Squad (Epics/Tasks)
};

// JIRA OVERVIEW CONFIGURATION - Easy to update if location changes
const JIRA_OVERVIEW_CONFIG = {
  spreadsheetId: '1Kzeg9cutpMekr_pBCiJBoNpARinZabQI7s3IiFCouE4',
  sheetName: 'Registry'
};

/**
 * Get Ideas list from Script Properties
 * Parses GREEN_IDEAS_LIST and returns array of objects with value and display text
 * 
 * Script Property Format: GMAP-14:Start.ca;GMAP-15:Optik-5;GMAP-16:Project X
 * Returns: [{ value: "GMAP-14", display: "GMAP-14: Start.ca" }, ...]
 */
function getIdeasList() {
  try {
    // Get from Script Properties (like JIRA_API_TOKEN)
    const ideasListString = PropertiesService.getScriptProperties().getProperty('GREEN_IDEAS_LIST');
    
    if (!ideasListString) {
      Logger.log('⚠️ GREEN_IDEAS_LIST not found in Script Properties');
      return [];
    }
    
    Logger.log('📋 GREEN_IDEAS_LIST: ' + ideasListString);
    
    // Split by semicolon and process each idea
    const ideas = ideasListString.split(';').map(item => item.trim()).filter(item => item);
    
    return ideas.map(idea => {
      // Split by colon to separate GMAP-XX from description
      const parts = idea.split(':');
      const value = parts[0].trim();  // "GMAP-14"
      const description = parts[1] ? parts[1].trim() : '';  // "Start.ca"
      
      return {
        value: value,                                    // For JQL query: "GMAP-14"
        display: description ? `${value}: ${description}` : value  // Display: "GMAP-14: Start.ca"
      };
    });
  } catch (error) {
    Logger.log('❌ Error parsing GREEN_IDEAS_LIST: ' + error.message);
    return [];
  }
}

/**
 * Get Labels list from Script Properties
 * Parses GREEN_LABELS_LIST and returns array of label values
 * 
 * Script Property Format: Phase1;Phase2;Phase3
 * Returns: ["Phase1", "Phase2", "Phase3"]
 */
function getLabelsList() {
  try {
    // Get from Script Properties (like JIRA_API_TOKEN)
    const labelsListString = PropertiesService.getScriptProperties().getProperty('GREEN_LABELS_LIST');
    
    if (!labelsListString) {
      Logger.log('⚠️ GREEN_LABELS_LIST not found in Script Properties');
      return [];
    }
    
    Logger.log('🏷️ GREEN_LABELS_LIST: ' + labelsListString);
    
    // Split by semicolon and return array
    return labelsListString.split(';').map(item => item.trim()).filter(item => item);
  } catch (error) {
    Logger.log('❌ Error parsing GREEN_LABELS_LIST: ' + error.message);
    return [];
  }
}

/**
 * TEST FUNCTION - Verify GREEN configuration properties
 * Run this in Apps Script editor to check if properties are set correctly
 */
function TEST_GREEN_CONFIG() {
  Logger.clear();
  Logger.log('='.repeat(80));
  Logger.log('🧪 TESTING GREEN CONFIGURATION PROPERTIES');
  Logger.log('='.repeat(80));
  
  // Test IDEAS_LIST
  Logger.log('\n📋 Testing GREEN_IDEAS_LIST...');
  const ideas = getIdeasList();
  Logger.log('✅ Found ' + ideas.length + ' ideas');
  ideas.forEach(idea => {
    Logger.log('   • Value: "' + idea.value + '" → Display: "' + idea.display + '"');
  });
  
  // Test LABELS_LIST
  Logger.log('\n🏷️ Testing GREEN_LABELS_LIST...');
  const labels = getLabelsList();
  Logger.log('✅ Found ' + labels.length + ' labels');
  labels.forEach(label => {
    Logger.log('   • ' + label);
  });
  
  Logger.log('\n' + '='.repeat(80));
  Logger.log('✅ TEST COMPLETE');
  Logger.log('='.repeat(80));
}

/**
 * Serve the HTML page for the GREEN dashboard
 */
function doGet(e) {
  const htmlContent = HtmlService.createHtmlOutputFromFile('GREEN').getContent();
  return HtmlService.createHtmlOutput(htmlContent)
    .setTitle('GREEN Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * TOKEN DIAGNOSTIC TEST - Tests the exact JQL queries the dashboard uses for GMAP-14
 * Run this to verify the service account has access to the ideas in GREEN_IDEAS_LIST
 */
function TEST_AUTH_DEBUG() {
  Logger.clear();
  Logger.log('='.repeat(80));
  Logger.log('🔬 DASHBOARD ACCESS DIAGNOSTIC TEST');
  Logger.log('='.repeat(80));
  
  const props = PropertiesService.getScriptProperties();
  const apiToken = props.getProperty('JIRA_API_TOKEN');
  const cloudId = props.getProperty('JIRA_CLOUD_ID');
  
  if (!apiToken || !cloudId) {
    Logger.log('❌ JIRA_API_TOKEN or JIRA_CLOUD_ID not set!');
    return;
  }
  
  const cleanToken = apiToken.trim();
  const baseUrl = 'https://api.atlassian.com/ex/jira/' + cloudId.trim();
  Logger.log('Base URL: ' + baseUrl);
  Logger.log('Token length: ' + cleanToken.length + ' chars');
  
  // Helper to run a JQL query and log results
  function runJQL(label, jql, fields) {
    Logger.log('\n' + '─'.repeat(60));
    Logger.log('📋 ' + label);
    Logger.log('   JQL: ' + jql);
    try {
      var r = UrlFetchApp.fetch(baseUrl + '/rest/api/3/search/jql', {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + cleanToken,
          'Accept': 'application/json',
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({ jql: jql, maxResults: 5, fields: fields || ['key', 'summary', 'status', 'issuetype'] }),
        muteHttpExceptions: true
      });
      var code = r.getResponseCode();
      var body = r.getContentText();
      Logger.log('   Response Code: ' + code);
      if (code === 200) {
        var d = JSON.parse(body);
        Logger.log('   Total found: ' + d.total);
        if (d.issues && d.issues.length > 0) {
          d.issues.forEach(function(issue) {
            Logger.log('   ✅ ' + issue.key + ' | ' + (issue.fields.issuetype ? issue.fields.issuetype.name : '?') + ' | ' + issue.fields.summary);
          });
        } else {
          Logger.log('   ❌ 0 results — service account cannot see these issues');
        }
      } else {
        Logger.log('   ❌ Error: ' + body.substring(0, 300));
      }
    } catch (e) {
      Logger.log('   ❌ Exception: ' + e.message);
    }
  }
  
  // ── TEST 1: GMAP-5 (known to work) ───────────────────────────────────────
  runJQL('TEST 1: key = GMAP-5 (known working issue)', 'key = GMAP-5');
  
  // ── TEST 2: GMAP-14 (dashboard idea) ─────────────────────────────────────
  runJQL('TEST 2: key = GMAP-14 (dashboard idea)', 'key = GMAP-14');
  
  // ── TEST 3: Exact dashboard Step 1 JQL for GMAP-14 ───────────────────────
  runJQL('TEST 3: Exact dashboard Step 1 JQL (GMAP-14)', 'project = GMAP AND key = GMAP-14');
  
  // ── TEST 4: Exact dashboard Step 2 JQL — Capabilities linked to GMAP-14 ──
  runJQL('TEST 4: Capabilities linked to GMAP-14 (dashboard Step 2)', 
    'project = GREEN AND type = Capability AND issue in linkedIssues(GMAP-14) ORDER BY created DESC');
  
  // ── TEST 5: All ideas from GREEN_IDEAS_LIST ───────────────────────────────
  Logger.log('\n' + '─'.repeat(60));
  Logger.log('📋 TEST 5: All ideas from GREEN_IDEAS_LIST Script Property');
  var ideasListString = props.getProperty('GREEN_IDEAS_LIST');
  if (!ideasListString) {
    Logger.log('   ❌ GREEN_IDEAS_LIST not set in Script Properties');
  } else {
    Logger.log('   GREEN_IDEAS_LIST: ' + ideasListString);
    var ideaKeys = ideasListString.split(';').map(function(item) {
      return item.trim().split(':')[0].trim();
    }).filter(function(k) { return k; });
    Logger.log('   Found ' + ideaKeys.length + ' ideas: ' + ideaKeys.join(', '));
    
    // Test each idea key
    ideaKeys.forEach(function(ideaKey) {
      runJQL('  → Testing access to ' + ideaKey, 'key = ' + ideaKey);
    });
  }
  
  // ── TEST 6: Specific child-level issues (GREEN, SODP, VOGS) ──────────────
  Logger.log('\n' + '─'.repeat(60));
  Logger.log('📋 TEST 6: Child-level issue access (GREEN Capabilities, SODP & VOGS Epics)');
  
  // 6a: GREEN-177 (Capability)
  runJQL('  6a: key = GREEN-177 (Capability)', 'key = GREEN-177');
  
  // 6b: SODP-12152 (Epic)
  runJQL('  6b: key = SODP-12152 (SODP Epic)', 'key = SODP-12152');
  
  // 6c: VOGS-812 (Epic)
  runJQL('  6c: key = VOGS-812 (VOGS Epic)', 'key = VOGS-812');
  
  // 6d: Children of GREEN-177 (Epics/Tasks under a Capability)
  runJQL('  6d: parent in ("GREEN-177") — child Epics/Tasks under Capability', 'parent in ("GREEN-177")');
  
  // 6e: Children of SODP-12152 (Stories under a SODP Epic)
  runJQL('  6e: parent in ("SODP-12152") — Stories under SODP Epic', 'parent in ("SODP-12152")');
  
  // 6f: Children of VOGS-812 (Stories under a VOGS Epic)
  runJQL('  6f: parent in ("VOGS-812") — Stories under VOGS Epic', 'parent in ("VOGS-812")');
  
  // 6g: All GREEN Capabilities (broad project access check)
  runJQL('  6g: project = GREEN AND type = Capability (broad access check)', 'project = GREEN AND type = Capability ORDER BY created DESC');
  
  // 6h: All SODP Epics (broad project access check)
  runJQL('  6h: project = SODP AND type = Epic (broad access check)', 'project = SODP AND type = Epic ORDER BY created DESC');
  
  // 6i: All VOGS Epics (broad project access check)
  runJQL('  6i: project = VOGS AND type = Epic (broad access check)', 'project = VOGS AND type = Epic ORDER BY created DESC');
  
  Logger.log('\n' + '='.repeat(80));
  Logger.log('📊 SUMMARY');
  Logger.log('   If TEST 1 (GMAP-5) works but TEST 2 (GMAP-14) fails:');
  Logger.log('   → Service account needs project-level access to GMAP and GREEN projects');
  Logger.log('   → Contact Admin team to grant the service account access to GMAP and GREEN projects');
  Logger.log('   TEST 6 results show if service account can see child Epics/Stories in GREEN, SODP, VOGS');
  Logger.log('='.repeat(80));
}

/**
 * TEST FUNCTION - Calls fetchGreenData() directly to reproduce browser crashes
 * Change ideaKey and label to match the failing combination
 */
function TEST_FETCH_GREEN_DATA() {
  Logger.clear();
  Logger.log('='.repeat(80));
  Logger.log('🧪 TEST_FETCH_GREEN_DATA - Simulating browser call to fetchGreenData()');
  Logger.log('='.repeat(80));
  
  // ── CHANGE THESE TO MATCH THE FAILING COMBINATION ──────────────────────
  var ideaKey = 'GMAP-5';          // e.g. 'GMAP-5', 'GMAP-14'
  var phaseLabel = 'Optik_It1.2';  // e.g. 'Optik_It1.2', 'Optik_IT1.4', ''
  // ────────────────────────────────────────────────────────────────────────
  
  Logger.log('ideaKey: ' + ideaKey);
  Logger.log('phaseLabel: ' + phaseLabel);
  Logger.log('');
  
  try {
    var result = fetchGreenData(ideaKey, phaseLabel);
    
    Logger.log('');
    Logger.log('='.repeat(80));
    Logger.log('📊 RESULT:');
    Logger.log('  success: ' + result.success);
    Logger.log('  message: ' + result.message);
    if (result.data) {
      Logger.log('  data items: ' + result.data.length);
    }
    if (result.summary) {
      Logger.log('  summary: ' + JSON.stringify(result.summary));
    }
    Logger.log('');
    Logger.log('📋 DEBUG LOG (' + (result.debugLog ? result.debugLog.length : 0) + ' entries):');
    if (result.debugLog) {
      result.debugLog.forEach(function(line) { Logger.log(line); });
    }
    Logger.log('='.repeat(80));
    
  } catch (e) {
    Logger.log('');
    Logger.log('💥 UNCAUGHT EXCEPTION:');
    Logger.log('  Message: ' + e.message);
    Logger.log('  Stack: ' + e.stack);
    Logger.log('='.repeat(80));
  }
}

/**
 * TEST FUNCTION - Run this in Apps Script editor to verify JIRA connection
 */
function TEST_GREEN_CONNECTION() {
  Logger.clear();
  Logger.log('='.repeat(80));
  Logger.log('🧪 TESTING GREEN JIRA CONNECTION');
  Logger.log('='.repeat(80));
  
  const debugLog = [];
  
  // Get JIRA configuration
  const config = getJiraConfig(debugLog);
  if (!config.success) {
    Logger.log('❌ Configuration failed: ' + config.message);
    debugLog.forEach(log => Logger.log(log));
    return;
  }
  
  Logger.log('✅ Configuration successful');
  
  // Test 1: Fetch a specific Idea
  Logger.log('\n📋 TEST 1: Fetch Idea GMAP-14');
  const ideaJQL = 'project = GMAP AND key = GMAP-14';
  const ideaResult = testQuery(ideaJQL, config.baseUrl, config.authHeader);
  Logger.log('Result: ' + (ideaResult.success ? '✅ SUCCESS' : '❌ FAILED'));
  Logger.log('Issues found: ' + ideaResult.count);
  
  // Test 2: Fetch Capabilities linked to GMAP-14
  Logger.log('\n📋 TEST 2: Fetch Capabilities linked to GMAP-14');
  const capJQL = 'project = GREEN AND type = Capability AND issue in linkedIssues(GMAP-14)';
  const capResult = testQuery(capJQL, config.baseUrl, config.authHeader);
  Logger.log('Result: ' + (capResult.success ? '✅ SUCCESS' : '❌ FAILED'));
  Logger.log('Capabilities found: ' + capResult.count);
  if (capResult.count > 0) {
    Logger.log('Capability keys: ' + capResult.issues.join(', '));
  }
  
  // Test 3: Fetch Capabilities with Phase1 label
  Logger.log('\n📋 TEST 3: Fetch Capabilities with Phase1 label');
  const phaseJQL = 'project = GREEN AND type = Capability AND issue in linkedIssues(GMAP-14) AND labels = "Phase1"';
  const phaseResult = testQuery(phaseJQL, config.baseUrl, config.authHeader);
  Logger.log('Result: ' + (phaseResult.success ? '✅ SUCCESS' : '❌ FAILED'));
  Logger.log('Phase1 Capabilities found: ' + phaseResult.count);
  
  Logger.log('\n' + '='.repeat(80));
  Logger.log('✅ TEST COMPLETE');
  Logger.log('='.repeat(80));
}

/**
 * ENHANCED DIAGNOSTIC TEST - Run this to diagnose JIRA connection issues
 * This provides detailed information about what's working and what's not
 */
function TEST_GREEN_CONNECTION_DETAILED() {
  Logger.clear();
  Logger.log('='.repeat(80));
  Logger.log('🔬 ENHANCED GREEN JIRA CONNECTION DIAGNOSTIC');
  Logger.log('='.repeat(80));
  
  // Step 1: Check Script Properties
  Logger.log('\n📋 STEP 1: Checking Script Properties...');
  const props      = PropertiesService.getScriptProperties();
  const apiToken   = props.getProperty('JIRA_API_TOKEN');
  const cloudId    = props.getProperty('JIRA_CLOUD_ID');

  Logger.log('   JIRA_API_TOKEN: ' + (apiToken ? '✅ SET (length: ' + apiToken.length + ')' : '❌ NOT SET'));
  Logger.log('   JIRA_CLOUD_ID:  ' + (cloudId  ? '✅ SET → ' + cloudId                     : '❌ NOT SET'));
  Logger.log('   Auth mode: Bearer token (service account via Atlassian API Gateway)');

  if (!apiToken) {
    Logger.log('\n❌ CRITICAL: JIRA_API_TOKEN is not set!');
    Logger.log('Set it in Project Settings → Script Properties (value = secret from go/secureshare)');
    return;
  }
  if (!cloudId) {
    Logger.log('\n❌ CRITICAL: JIRA_CLOUD_ID is not set!');
    Logger.log('Get it from https://telus-cio.atlassian.net/_edge/tenant_info and set it in Script Properties');
    return;
  }

  const actualBaseUrl = `https://api.atlassian.com/ex/jira/${cloudId}`;
  const authHeader    = 'Bearer ' + apiToken;
  Logger.log('   Base URL: ' + actualBaseUrl);
  
  // Step 2: Test API Authentication using /search/jql (service account token has read:jira-work scope, not read:me)
  Logger.log('\n📋 STEP 2: Testing API Authentication via /search/jql (NOT /myself - service account tokens do not support /myself)...');
  try {
    const authTestUrl = actualBaseUrl + '/rest/api/3/search/jql';
    const authTestResponse = UrlFetchApp.fetch(authTestUrl, {
      method: 'POST',
      headers: {
        'Authorization': authHeader,
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ jql: 'key = GMAP-5', maxResults: 1, fields: ['key', 'summary'] }),
      muteHttpExceptions: true
    });
    
    const authTestCode = authTestResponse.getResponseCode();
    Logger.log('   Response Code: ' + authTestCode);
    
    if (authTestCode === 200) {
      Logger.log('   ✅ Authentication SUCCESS');
      Logger.log('   Token is valid and has read:jira-work scope');
      Logger.log('   Note: /myself endpoint is not used (service account token scope is read:jira-work, not read:me)');
    } else {
      Logger.log('   ❌ Authentication FAILED');
      Logger.log('   Response: ' + authTestResponse.getContentText());
      return;
    }
  } catch (error) {
    Logger.log('   ❌ Error: ' + error.message);
    return;
  }
  
  // Step 3: Can I see ANY issues at all?
  Logger.log('\n📋 STEP 3: Testing if I can see ANY issues in JIRA...');
  try {
    // Use a valid JQL query - get issues created in last 2 years
    const anyIssueJQL = 'created >= -730d order by created DESC';
    const anyIssueResult = testQueryDetailed(anyIssueJQL, actualBaseUrl, authHeader);
    
    Logger.log('   Response Code: ' + anyIssueResult.code);
    Logger.log('   Total issues I can access: ' + anyIssueResult.total);
    
    if (anyIssueResult.code !== 200) {
      Logger.log('   ❌ API request failed');
      Logger.log('   Error: ' + anyIssueResult.error);
    } else if (anyIssueResult.total > 0) {
      Logger.log('   ✅ I CAN see issues in JIRA');
      Logger.log('   Sample issues: ' + anyIssueResult.sampleKeys.join(', '));
    } else {
      Logger.log('   ❌ I CANNOT see ANY issues in JIRA');
      Logger.log('   This suggests a permission problem with the API token');
    }
  } catch (error) {
    Logger.log('   ❌ Error: ' + error.message);
  }
  
  // Step 4: Can I access the GMAP project?
  Logger.log('\n📋 STEP 4: Testing access to GMAP project...');
  try {
    const gmapProjectJQL = 'project = GMAP order by created DESC';
    const gmapResult = testQueryDetailed(gmapProjectJQL, actualBaseUrl, authHeader);
    
    Logger.log('   Response Code: ' + gmapResult.code);
    Logger.log('   GMAP issues found: ' + gmapResult.total);
    
    if (gmapResult.total > 0) {
      Logger.log('   ✅ I CAN access GMAP project');
      Logger.log('   Sample GMAP issues: ' + gmapResult.sampleKeys.join(', '));
    } else {
      Logger.log('   ❌ I CANNOT access GMAP project (or it has no issues)');
      Logger.log('   Possible reasons:');
      Logger.log('      • GMAP project doesn\'t exist in this JIRA instance');
      Logger.log('      • No permission to view GMAP project');
      Logger.log('      • GMAP project is empty');
    }
  } catch (error) {
    Logger.log('   ❌ Error: ' + error.message);
  }
  
  // Step 5: Can I see GMAP-14 specifically?
  Logger.log('\n📋 STEP 5: Testing access to GMAP-14 specifically...');
  try {
    const gmap14JQL = 'key = GMAP-14';
    const gmap14Result = testQueryDetailed(gmap14JQL, actualBaseUrl, authHeader);
    
    Logger.log('   Response Code: ' + gmap14Result.code);
    Logger.log('   GMAP-14 found: ' + (gmap14Result.total > 0 ? 'YES' : 'NO'));
    
    if (gmap14Result.total > 0) {
      Logger.log('   ✅ I CAN see GMAP-14');
      Logger.log('   Summary: ' + gmap14Result.summary);
    } else {
      Logger.log('   ❌ I CANNOT see GMAP-14');
      Logger.log('   Possible reasons:');
      Logger.log('      • GMAP-14 doesn\'t exist in this JIRA instance');
      Logger.log('      • No permission to view GMAP-14');
    }
  } catch (error) {
    Logger.log('   ❌ Error: ' + error.message);
  }
  
  // Step 6: Can I access the GREEN project?
  Logger.log('\n📋 STEP 6: Testing access to GREEN project...');
  try {
    const greenProjectJQL = 'project = GREEN order by created DESC';
    const greenResult = testQueryDetailed(greenProjectJQL, actualBaseUrl, authHeader);
    
    Logger.log('   Response Code: ' + greenResult.code);
    Logger.log('   GREEN issues found: ' + greenResult.total);
    
    if (greenResult.total > 0) {
      Logger.log('   ✅ I CAN access GREEN project');
      Logger.log('   Sample GREEN issues: ' + greenResult.sampleKeys.join(', '));
    } else {
      Logger.log('   ❌ I CANNOT access GREEN project (or it has no issues)');
    }
  } catch (error) {
    Logger.log('   ❌ Error: ' + error.message);
  }
  
  // Summary
  Logger.log('\n' + '='.repeat(80));
  Logger.log('📊 DIAGNOSTIC SUMMARY');
  Logger.log('='.repeat(80));
  Logger.log('Run this test and share the FULL log output to diagnose the issue.');
  Logger.log('='.repeat(80));
}

/**
 * Helper function for detailed query testing
 */
function testQueryDetailed(jql, baseUrl, authHeader) {
  try {
    const searchUrl = baseUrl + '/rest/api/3/search/jql';
    const requestBody = {
      jql: jql,
      maxResults: 5,
      fields: ['key', 'summary']
    };
    
    const response = UrlFetchApp.fetch(searchUrl, {
      method: 'POST',
      headers: {
        'Authorization': authHeader,
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    });
    
    const code = response.getResponseCode();
    const responseText = response.getContentText();
    
    // Try to parse response
    let data;
    try {
      data = JSON.parse(responseText);
    } catch (parseError) {
      return {
        code: code,
        total: 0,
        sampleKeys: [],
        summary: 'N/A',
        error: 'Failed to parse response: ' + responseText.substring(0, 200)
      };
    }
    
    // Check for error messages in response
    if (code !== 200) {
      const errorMsg = data.errorMessages ? data.errorMessages.join(', ') : responseText.substring(0, 200);
      return {
        code: code,
        total: 0,
        sampleKeys: [],
        summary: 'N/A',
        error: errorMsg
      };
    }
    
    return {
      code: code,
      total: data.total || 0,
      sampleKeys: (data.issues || []).map(i => i.key),
      summary: (data.issues && data.issues.length > 0) ? data.issues[0].fields.summary : 'N/A',
      error: null
    };
  } catch (error) {
    return {
      code: 0,
      total: 0,
      sampleKeys: [],
      summary: 'N/A',
      error: error.message
    };
  }
}

/**
 * Helper function to test a single query
 */
function testQuery(jql, baseUrl, authHeader) {
  try {
    const searchUrl = `${baseUrl}/rest/api/3/search/jql`;
    const requestBody = {
      jql: jql,
      maxResults: 100,
      fields: ['key', 'summary']
    };
    
    const requestOptions = {
      method: 'POST',
      headers: {
        'Authorization': authHeader,
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(searchUrl, requestOptions);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      return { success: false, count: 0, issues: [] };
    }
    
    const data = JSON.parse(response.getContentText());
    const issues = data.issues || [];
    
    return {
      success: true,
      count: issues.length,
      issues: issues.map(i => i.key)
    };
    
  } catch (error) {
    Logger.log('Error: ' + error.message);
    return { success: false, count: 0, issues: [] };
  }
}

/**
 * Get JIRA configuration and setup authentication.
 * Uses a service account Bearer API token via the Atlassian API Gateway.
 *
 * Required Script Properties:
 *   JIRA_API_TOKEN - Service account Bearer token (secret from go/secureshare)
 *   JIRA_CLOUD_ID  - Atlassian Cloud ID (from https://telus-cio.atlassian.net/_edge/tenant_info)
 *
 * API Gateway base URL: https://api.atlassian.com/ex/jira/{cloudId}
 */
function getJiraConfig(debugLog) {
  try {
    debugLog.push(`[${new Date().toISOString()}] 🔧 Configuring JIRA connection...`);

    const props      = PropertiesService.getScriptProperties();
    const apiToken   = props.getProperty('JIRA_API_TOKEN');
    const cloudId    = props.getProperty('JIRA_CLOUD_ID');

    if (!apiToken) {
      debugLog.push(`[${new Date().toISOString()}] ❌ JIRA_API_TOKEN not set in Script Properties`);
      return {
        success: false,
        message: 'JIRA_API_TOKEN not configured. Set it in Project Settings → Script Properties.',
        debugLog: debugLog
      };
    }

    if (!cloudId) {
      debugLog.push(`[${new Date().toISOString()}] ❌ JIRA_CLOUD_ID not set in Script Properties`);
      return {
        success: false,
        message: 'JIRA_CLOUD_ID not configured. Get it from https://telus-cio.atlassian.net/_edge/tenant_info and set it in Script Properties.',
        debugLog: debugLog
      };
    }

    // Atlassian API Gateway URL — required for enterprise service account tokens
    const baseUrl = `https://api.atlassian.com/ex/jira/${cloudId}`;
    debugLog.push(`[${new Date().toISOString()}] 🌐 Using base URL: ${baseUrl}`);
    debugLog.push(`[${new Date().toISOString()}] 🔑 Auth mode: Bearer token (service account)`);

    const authHeader = 'Bearer ' + apiToken;
    debugLog.push(`[${new Date().toISOString()}] ✅ Configuration successful (Bearer auth)`);

    return {
      success: true,
      baseUrl: baseUrl,
      authHeader: authHeader
    };

  } catch (error) {
    debugLog.push(`[${new Date().toISOString()}] 💥 Error: ${error.message}`);
    return {
      success: false,
      message: 'Configuration error: ' + error.message,
      debugLog: debugLog
    };
  }
}

/**
 * Build team mapping from Google Sheet data
 * Returns a map of project prefix to team name
 */
function buildTeamMapping(debugLog) {
  try {
    debugLog.push(`[${new Date().toISOString()}] 📊 Fetching team mapping from Google Sheet...`);
    
    const sheetResult = fetchJiraOverviewData();
    
    if (!sheetResult.success) {
      debugLog.push(`[${new Date().toISOString()}] ⚠️ Failed to fetch Google Sheet: ${sheetResult.message}`);
      debugLog.push(`[${new Date().toISOString()}] ⚠️ Will use 'Unknown' for all teams`);
      return {}; // Return empty map - will default to 'Unknown'
    }
    
    const teamMap = {};
    const data = sheetResult.data || [];
    
    debugLog.push(`[${new Date().toISOString()}] 📊 Processing ${data.length} rows from Google Sheet...`);
    
    // Build mapping: Extract project prefix from "Jira Issue Key" → map to "Team"
    data.forEach(row => {
      const jiraKey = row['Jira Issue Key'] || '';
      const teamName = row['Team'] || '';
      
      if (jiraKey && teamName) {
        // Extract project prefix (e.g., 'GSS' from 'GSS-123')
        const prefix = jiraKey.split('-')[0];
        if (prefix) {
          teamMap[prefix] = teamName;
        }
      }
    });
    
    debugLog.push(`[${new Date().toISOString()}] ✅ Built team mapping with ${Object.keys(teamMap).length} project prefixes`);
    debugLog.push(`[${new Date().toISOString()}] 📋 Sample mappings: ${JSON.stringify(Object.fromEntries(Object.entries(teamMap).slice(0, 5)))}`);
    
    return teamMap;
    
  } catch (error) {
    debugLog.push(`[${new Date().toISOString()}] ⚠️ Exception building team mapping: ${error.message}`);
    debugLog.push(`[${new Date().toISOString()}] ⚠️ Will use 'Unknown' for all teams`);
    return {}; // Return empty map on error
  }
}

/**
 * Build team leadership mapping from Google Sheet data
 * Returns a map of team name to { director, manager }
 */
function buildTeamLeadershipMapping(debugLog) {
  try {
    debugLog.push(`[${new Date().toISOString()}] 👔 Building team leadership mapping from Google Sheet...`);
    
    const sheetResult = fetchJiraOverviewData();
    
    if (!sheetResult.success) {
      debugLog.push(`[${new Date().toISOString()}] ⚠️ Failed to fetch Google Sheet: ${sheetResult.message}`);
      debugLog.push(`[${new Date().toISOString()}] ⚠️ Will use 'Unknown' for all leadership`);
      return {}; // Return empty map - will default to 'Unknown'
    }
    
    const leadershipMap = {};
    const data = sheetResult.data || [];
    
    debugLog.push(`[${new Date().toISOString()}] 📊 Processing ${data.length} rows from Google Sheet...`);
    
    // Build mapping: Team → { director, manager }
    data.forEach(row => {
      const teamName = row['Team'] || '';
      const director = row['Engineering Director'] || '';
      const manager = row['Engineering Manager'] || '';
      
      if (teamName) {
        // Store the first occurrence of each team's leadership
        if (!leadershipMap[teamName]) {
          leadershipMap[teamName] = {
            director: director || 'Unknown',
            manager: manager || 'Unknown'
          };
        }
      }
    });
    
    debugLog.push(`[${new Date().toISOString()}] ✅ Built leadership mapping for ${Object.keys(leadershipMap).length} teams`);
    debugLog.push(`[${new Date().toISOString()}] 📋 Sample mappings: ${JSON.stringify(Object.fromEntries(Object.entries(leadershipMap).slice(0, 3)))}`);
    
    return leadershipMap;
    
  } catch (error) {
    debugLog.push(`[${new Date().toISOString()}] ⚠️ Exception building leadership mapping: ${error.message}`);
    debugLog.push(`[${new Date().toISOString()}] ⚠️ Will use 'Unknown' for all leadership`);
    return {}; // Return empty map on error
  }
}

/**
 * Main function to fetch GREEN data based on Idea and Phase Label
 * Called from GREEN.html
 */
function fetchGreenData(ideaKey, phaseLabel) {
  const debugLog = [];
  
  try {
    debugLog.push(`[${new Date().toISOString()}] 🚀 Starting GREEN data fetch...`);
    debugLog.push(`[${new Date().toISOString()}] ${'='.repeat(80)}`);
    
    // DEBUG: Log raw parameters received from HTML
    debugLog.push(`[${new Date().toISOString()}] 📥 RAW PARAMETERS RECEIVED FROM HTML:`);
    debugLog.push(`[${new Date().toISOString()}]    ideaKey = "${ideaKey}"`);
    debugLog.push(`[${new Date().toISOString()}]    ideaKey type = ${typeof ideaKey}`);
    debugLog.push(`[${new Date().toISOString()}]    ideaKey length = ${ideaKey ? ideaKey.length : 'N/A'}`);
    debugLog.push(`[${new Date().toISOString()}]    ideaKey is null? ${ideaKey === null}`);
    debugLog.push(`[${new Date().toISOString()}]    ideaKey is undefined? ${ideaKey === undefined}`);
    debugLog.push(`[${new Date().toISOString()}]    phaseLabel = "${phaseLabel}"`);
    debugLog.push(`[${new Date().toISOString()}]    phaseLabel type = ${typeof phaseLabel}`);
    debugLog.push(`[${new Date().toISOString()}]    phaseLabel length = ${phaseLabel ? phaseLabel.length : 'N/A'}`);
    debugLog.push(`[${new Date().toISOString()}] ${'='.repeat(80)}`);
    
    debugLog.push(`[${new Date().toISOString()}] 📋 Idea: ${ideaKey}, Phase: ${phaseLabel || 'All Phases'}`);
    
    // Step 0: Read Google Sheet ONCE and cache it, then build all mappings from the cached data
    debugLog.push(`[${new Date().toISOString()}] ⏱️ START: Reading Google Sheet (single cached read)...`);
    const sheetStartTime = new Date();
    const cachedSheetData = fetchJiraOverviewData();
    const sheetElapsed = ((new Date() - sheetStartTime) / 1000).toFixed(1);
    debugLog.push(`[${new Date().toISOString()}] ⏱️ END: Google Sheet read took ${sheetElapsed}s — success: ${cachedSheetData.success}`);

    const teamMapping = buildTeamMappingFromCache(cachedSheetData, debugLog);
    const leadershipMapping = buildTeamLeadershipMappingFromCache(cachedSheetData, debugLog);
    const applicationToTeamMapping = buildApplicationToTeamMappingFromCache(cachedSheetData, debugLog);
    
    // Get JIRA configuration
    const config = getJiraConfig(debugLog);
    if (!config.success) {
      return config;
    }
    
    const { baseUrl, authHeader } = config;
    
    // Step 1: Fetch the selected Idea from GMAP
    debugLog.push(`[${new Date().toISOString()}] 📋 STEP 1: Fetching Idea ${ideaKey}...`);
    // Use variable from HTML dropdown (no quotes around the variable)
    const ideaJQL = `project = GMAP AND key = ${ideaKey}`;
    debugLog.push(`[${new Date().toISOString()}] 📝 Idea JQL: ${ideaJQL}`);
    const ideaResult = fetchJiraIssues(ideaJQL, baseUrl, authHeader, debugLog);
    
    if (!ideaResult.success || ideaResult.data.length === 0) {
      return {
        success: false,
        message: `Idea ${ideaKey} not found`,
        debugLog: debugLog
      };
    }
    
    const idea = ideaResult.data[0];
    debugLog.push(`[${new Date().toISOString()}] ✅ Found Idea: ${idea.key} - ${idea.fields.summary}`);
    
    // Step 2: Fetch Capabilities linked to this Idea
    debugLog.push(`[${new Date().toISOString()}] 📋 STEP 2: Fetching Capabilities linked to ${ideaKey}...`);
    
    let capabilityJQL;
    // phaseLabel can now be a single string OR an array of strings (multi-select)
    const phaseLabels = Array.isArray(phaseLabel) ? phaseLabel.filter(l => l && l !== '') : (phaseLabel && phaseLabel !== '' ? [phaseLabel] : []);
    
    if (phaseLabels.length > 0) {
      // Filter by one or more phase labels using JQL "labels in (...)" — OR logic
      const labelList = phaseLabels.map(l => `"${l}"`).join(', ');
      capabilityJQL = `project = GREEN AND type = Capability AND issue in linkedIssues(${ideaKey}) AND labels in (${labelList}) ORDER BY created DESC`;
      debugLog.push(`[${new Date().toISOString()}] 🏷️ Filtering by Phase(s): ${phaseLabels.join(', ')}`);
    } else {
      // Get all capabilities (no label filter)
      capabilityJQL = `project = GREEN AND type = Capability AND issue in linkedIssues(${ideaKey}) ORDER BY created DESC`;
      debugLog.push(`[${new Date().toISOString()}] 🏷️ No phase filter - fetching all capabilities`);
    }
    
    debugLog.push(`[${new Date().toISOString()}] 📝 Capability JQL: ${capabilityJQL}`);
    const capabilityResult = fetchJiraIssues(capabilityJQL, baseUrl, authHeader, debugLog);
    
    if (!capabilityResult.success) {
      return capabilityResult;
    }
    
    const capabilities = capabilityResult.data || [];
    debugLog.push(`[${new Date().toISOString()}] ✅ Found ${capabilities.length} Capabilities`);
    
    // Step 3: For each Capability, fetch child Epics and Tasks
    let allChildren = [];
    
    if (capabilities.length > 0) {
      debugLog.push(`[${new Date().toISOString()}] 📋 STEP 3: Fetching child Epics and Tasks...`);
      
      const capabilityKeys = capabilities.map(cap => cap.key);
      
      // Fetch all children using parent field
      const childJQL = `parent in (${capabilityKeys.map(key => `"${key}"`).join(',')}) ORDER BY created DESC`;
      debugLog.push(`[${new Date().toISOString()}] 📝 Child JQL: ${childJQL}`);
      
      const childResult = fetchJiraIssues(childJQL, baseUrl, authHeader, debugLog);
      
      if (childResult.success) {
        allChildren = childResult.data || [];
        debugLog.push(`[${new Date().toISOString()}] ✅ Found ${allChildren.length} child items (Epics/Tasks)`);
      } else {
        debugLog.push(`[${new Date().toISOString()}] ⚠️ Failed to fetch children: ${childResult.message}`);
      }
    }
    
    // Step 4: Fetch Stories attached to Epics
    let allStories = [];
    
    if (allChildren.length > 0) {
      debugLog.push(`[${new Date().toISOString()}] 📋 STEP 4: Fetching Stories attached to Epics...`);
      
      // Get all Epic keys from children
      const epicKeys = allChildren
        .filter(child => child.fields.issuetype && child.fields.issuetype.name === 'Epic')
        .map(epic => epic.key);
      
      debugLog.push(`[${new Date().toISOString()}] 📊 Found ${epicKeys.length} Epics to fetch Stories from`);
      
      if (epicKeys.length > 0) {
        const storyJQL = `parent in (${epicKeys.map(key => `"${key}"`).join(',')}) ORDER BY created DESC`;
        debugLog.push(`[${new Date().toISOString()}] 📝 Story JQL: ${storyJQL}`);
        
        const storyResult = fetchJiraIssues(storyJQL, baseUrl, authHeader, debugLog);
        
        if (storyResult.success) {
          allStories = storyResult.data || [];
          debugLog.push(`[${new Date().toISOString()}] ✅ Found ${allStories.length} Stories attached to Epics`);
        } else {
          debugLog.push(`[${new Date().toISOString()}] ⚠️ Failed to fetch Stories: ${storyResult.message}`);
        }
      } else {
        debugLog.push(`[${new Date().toISOString()}] ℹ️ No Epics found, skipping Story fetch`);
      }
    }
    
    // Combine all issues
    const allIssues = [idea, ...capabilities, ...allChildren, ...allStories];
    debugLog.push(`[${new Date().toISOString()}] ✅ Total issues: ${allIssues.length}`);
    
    // Process the data with team mapping, leadership mapping, application-to-team mapping, and cached sheet data
    const processedData = processGreenData(allIssues, teamMapping, leadershipMapping, applicationToTeamMapping, debugLog, cachedSheetData);
    
    debugLog.push(`[${new Date().toISOString()}] 🎉 Successfully completed GREEN data fetch!`);
    
    // ── PAYLOAD SIZE REDUCTION ──────────────────────────────────────────────
    // google.script.run has a ~50MB serialization limit.
    // Large datasets (Optik_It1.2: 94 items) can exceed this limit due to:
    //   • description field (full ADF JSON, can be 10-50KB per item)
    //   • Status/Risk field (ADF JSON)
    //   • debugLog array (hundreds of timestamped strings)
    // Fix: strip description from items (not used in dashboard views),
    //      truncate Status/Risk to plain text only, and drop debugLog from return.
    const slimItems = processedData.items.map(item => {
      const slim = Object.assign({}, item);
      // Remove description — large ADF JSON, not rendered in any dashboard view
      delete slim.description;
      // Truncate Status/Risk to 500 chars max
      if (slim['Status/Risk'] && slim['Status/Risk'].length > 500) {
        slim['Status/Risk'] = slim['Status/Risk'].substring(0, 500) + '…';
      }
      return slim;
    });
    
    // Keep only last 30 debug log entries to avoid payload bloat
    const slimDebugLog = debugLog.slice(-30);
    // ────────────────────────────────────────────────────────────────────────
    
    return {
      success: true,
      message: `Successfully fetched ${slimItems.length} items`,
      data: slimItems,
      summary: processedData.summary,
      leadershipMapping: leadershipMapping,  // Send leadership mapping to frontend
      debugLog: slimDebugLog
    };
    
  } catch (error) {
    debugLog.push(`[${new Date().toISOString()}] 💥 Exception: ${error.message}`);
    return {
      success: false,
      message: 'Exception occurred: ' + error.message,
      debugLog: debugLog,
      error: error.toString()
    };
  }
}

/**
 * Fetch JIRA issues using JQL query
 */
function fetchJiraIssues(jql, baseUrl, authHeader, debugLog) {
  try {
    debugLog.push(`[${new Date().toISOString()}] 🔍 Executing JQL query...`);
    
    const searchUrl = `${baseUrl}/rest/api/3/search/jql`;
    
    const requestBody = {
      jql: jql,
      maxResults: 1000,
      fields: [
        'key',
        'summary',
        'description',
        'status',
        'assignee',
        'created',
        'duedate',
        'issuetype',
        'labels',
        'parent',
        'customfield_19357',  // Health
        'customfield_17265',  // Gate 2 - Cost Estimate
        'customfield_24241',  // Person Days
        'customfield_21190',  // GTM Date
        'customfield_26082',  // Impacted Digital Products
        'customfield_30833',  // Impacted Applications
        'customfield_19361',  // Status/Risk
        'customfield_19257',  // T-shirt
        'customfield_16070',  // Start date (Dev Start Date)
        'customfield_24242',  // E2E Test Ready Date (Dev Completion)
        'customfield_21050',  // UAT Start Date
        'customfield_24243',  // UAT End Date
        'priority',           // Standard Jira Priority (Critical/High/Medium/Low)
        'customfield_19632'   // Severity of Defect (Critical/High/Medium/Low) — Experience Defects
      ]
    };
    
    const requestOptions = {
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
    
    const response = UrlFetchApp.fetch(searchUrl, requestOptions);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    debugLog.push(`[${new Date().toISOString()}] 📡 Response code: ${responseCode}`);
    
    if (responseCode !== 200) {
      debugLog.push(`[${new Date().toISOString()}] ❌ API request failed`);
      return {
        success: false,
        message: `JIRA API request failed with HTTP ${responseCode}`,
        debugLog: debugLog
      };
    }
    
    const responseData = JSON.parse(responseText);
    debugLog.push(`[${new Date().toISOString()}] 📊 Found ${responseData.issues ? responseData.issues.length : 0} issues`);
    
    // Log sample of customfield_30833 from first issue for verification
    if (responseData.issues && responseData.issues.length > 0) {
      const firstIssue = responseData.issues[0];
      debugLog.push(`[${new Date().toISOString()}] 🔍 SAMPLE - First issue: ${firstIssue.key}`);
      debugLog.push(`[${new Date().toISOString()}] 🔍 SAMPLE - customfield_30833 (Impacted Applications) raw value: ${JSON.stringify(firstIssue.fields.customfield_30833)}`);
    }
    
    return {
      success: true,
      data: responseData.issues || []
    };
    
  } catch (error) {
    debugLog.push(`[${new Date().toISOString()}] 💥 Error: ${error.message}`);
    return {
      success: false,
      message: 'JIRA fetch error: ' + error.message,
      debugLog: debugLog
    };
  }
}

/**
 * Project prefix to application name mapping
 * Based on the comprehensive application mapping table
 */
function getApplicationFromProjectKey(issueKey) {
  const PROJECT_TO_APP_MAPPING = {
    // Buy & Checkout Experience
    'BYOD': 'Plans and Addons',
    'ATCC': 'Hardware Shell + MFEs',
    'ATM': 'Manage Shell + MFEs (Reciepts, Cancel, PSO)',
    'VOGS': 'All Shells + MFEs pertaining to wireline journies',
    'CUXMFE': 'Credit Assessment MFE Digital Signature MFE',
    'UOST': 'Universal MFE (UOST), Green Checkout',
    'GXY': 'Universal MFE (UOST), Green Checkout',
    
    // Offer & Catalog Management
    'NCB': 'CloudBSS',
    'GEM': 'GEM',
    
    // Customer Care Experience
    'DJR': 'My TELUS Billing/Payments',
    'HOMERUN': 'My TELUS Appointments/Purple Manage Apps',
    'HFPF': 'My TELUS Overview',
    'BEAN': 'overview-api (My TELUS Overview BFF)/My TELUS B2B WLN',
    'GLOB': 'My TELUS Navigation',
    'SMS': 'My TELUS App (Expo - Front end)',
    'BFF2': 'My TELUS App (BFF)',
    'SSUP': 'My TELUS Plans & Devices',
    'FAST': 'My TELUS B2B WLS',
    
    // Customer Communications
    'ECOMM': 'Enterprise Communication Platform (ECP)',
    'C3': 'C30C',
    
    // Bespoke Solutions
    'OFNT': 'PaaS',
    
    // Field Workforce Management
    'TREE': 'SWORM',
    
    // Contact Centre Solutions
    'CCEP': 'Casa Partner',
    'GLAD': 'Casa TBS',
    'BEES': 'Casa CE',
    'TECH': 'Casa CE',
    'SMRPN': 'Casa CE',
    'SVN': 'Casa CE',
    
    // Design
    'PXD': 'Experience Design (Figma)'
  };
  
  if (!issueKey) return 'N/A';
  
  // Extract project prefix (everything before the dash)
  const prefix = issueKey.split('-')[0];
  
  // Return mapped application name or N/A if not found
  return PROJECT_TO_APP_MAPPING[prefix] || 'N/A';
}

// ─────────────────────────────────────────────────────────────────────────────
// CACHED MAPPING FUNCTIONS — use pre-fetched sheet data instead of re-reading
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Build team mapping from already-fetched sheet data (no extra sheet read)
 */
function buildTeamMappingFromCache(sheetResult, debugLog) {
  try {
    if (!sheetResult.success) {
      debugLog.push(`[${new Date().toISOString()}] ⚠️ [CACHE] Sheet data unavailable for team mapping — using empty map`);
      return {};
    }
    const teamMap = {};
    const data = sheetResult.data || [];
    data.forEach(row => {
      const jiraKey = row['Jira Issue Key'] || '';
      const teamName = row['Team'] || '';
      if (jiraKey && teamName) {
        const prefix = jiraKey.split('-')[0];
        if (prefix) teamMap[prefix] = teamName;
      }
    });
    debugLog.push(`[${new Date().toISOString()}] ✅ [CACHE] Team mapping built: ${Object.keys(teamMap).length} project prefixes`);
    return teamMap;
  } catch (error) {
    debugLog.push(`[${new Date().toISOString()}] ⚠️ [CACHE] Exception building team mapping: ${error.message}`);
    return {};
  }
}

/**
 * Build team leadership mapping from already-fetched sheet data (no extra sheet read)
 */
function buildTeamLeadershipMappingFromCache(sheetResult, debugLog) {
  try {
    if (!sheetResult.success) {
      debugLog.push(`[${new Date().toISOString()}] ⚠️ [CACHE] Sheet data unavailable for leadership mapping — using empty map`);
      return {};
    }
    const leadershipMap = {};
    const data = sheetResult.data || [];
    data.forEach(row => {
      const teamName = row['Team'] || '';
      const director = row['Engineering Director'] || '';
      const manager = row['Engineering Manager'] || '';
      if (teamName && !leadershipMap[teamName]) {
        leadershipMap[teamName] = {
          director: director || 'Unknown',
          manager: manager || 'Unknown'
        };
      }
    });
    debugLog.push(`[${new Date().toISOString()}] ✅ [CACHE] Leadership mapping built: ${Object.keys(leadershipMap).length} teams`);
    return leadershipMap;
  } catch (error) {
    debugLog.push(`[${new Date().toISOString()}] ⚠️ [CACHE] Exception building leadership mapping: ${error.message}`);
    return {};
  }
}

/**
 * Build Application-to-Team mapping from already-fetched sheet data (no extra sheet read)
 */
function buildApplicationToTeamMappingFromCache(sheetResult, debugLog) {
  try {
    if (!sheetResult.success) {
      debugLog.push(`[${new Date().toISOString()}] ⚠️ [CACHE] Sheet data unavailable for app-to-team mapping — using empty map`);
      return {};
    }
    const appToTeamMap = {};
    const data = sheetResult.data || [];
    data.forEach(row => {
      const applications = row['Application'] || '';
      const teamName = row['Team'] || '';
      if (applications && teamName) {
        applications.split(',').map(app => app.trim()).forEach(appName => {
          if (appName && !appToTeamMap[appName]) appToTeamMap[appName] = teamName;
        });
      }
    });
    debugLog.push(`[${new Date().toISOString()}] ✅ [CACHE] App-to-Team mapping built: ${Object.keys(appToTeamMap).length} entries`);
    return appToTeamMap;
  } catch (error) {
    debugLog.push(`[${new Date().toISOString()}] ⚠️ [CACHE] Exception building app-to-team mapping: ${error.message}`);
    return {};
  }
}

/**
 * Build Application-to-Project mapping from already-fetched sheet data (no extra sheet read)
 */
function buildApplicationToProjectMappingFromCache(sheetResult, debugLog) {
  try {
    if (!sheetResult.success) {
      debugLog.push(`[${new Date().toISOString()}] ⚠️ [CACHE] Sheet data unavailable for app-to-project mapping — using empty map`);
      return {};
    }
    const appToProjectMap = {};
    const data = sheetResult.data || [];
    data.forEach(row => {
      const applications = row['Application'] || '';
      const jiraKey = row['Jira Issue Key'] || '';
      if (applications && jiraKey) {
        const prefix = jiraKey.split('-')[0];
        applications.split(',').map(app => app.trim()).forEach(appName => {
          if (appName && prefix) appToProjectMap[appName] = prefix;
        });
      }
    });
    debugLog.push(`[${new Date().toISOString()}] ✅ [CACHE] App-to-Project mapping built: ${Object.keys(appToProjectMap).length} applications`);
    return appToProjectMap;
  } catch (error) {
    debugLog.push(`[${new Date().toISOString()}] ⚠️ [CACHE] Exception building app-to-project mapping: ${error.message}`);
    return {};
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// ORIGINAL MAPPING FUNCTIONS (kept for backward compatibility / standalone use)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Build Application Name → Team Name mapping from Google Sheet
 * Used as a fallback when an Epic's project prefix is not found in teamMapping
 * (e.g., CloudBSS teams that create Epics in the GREEN board instead of their own board)
 */
function buildApplicationToTeamMapping(debugLog) {
  try {
    debugLog.push(`[${new Date().toISOString()}] 🗺️ Building Application-to-Team mapping from Google Sheet...`);
    
    const sheetResult = fetchJiraOverviewData();
    
    if (!sheetResult.success) {
      debugLog.push(`[${new Date().toISOString()}] ⚠️ Failed to fetch Google Sheet for app-to-team mapping: ${sheetResult.message}`);
      return {};
    }
    
    const appToTeamMap = {};
    const data = sheetResult.data || [];
    
    data.forEach(row => {
      const applications = row['Application'] || '';
      const teamName = row['Team'] || '';
      
      if (applications && teamName) {
        // Split by comma in case multiple apps are listed in one cell
        const appList = applications.split(',').map(app => app.trim());
        appList.forEach(appName => {
          if (appName && !appToTeamMap[appName]) {
            appToTeamMap[appName] = teamName;
          }
        });
      }
    });
    
    debugLog.push(`[${new Date().toISOString()}] ✅ Built Application-to-Team mapping with ${Object.keys(appToTeamMap).length} entries`);
    debugLog.push(`[${new Date().toISOString()}] 📋 Sample: ${JSON.stringify(Object.fromEntries(Object.entries(appToTeamMap).slice(0, 5)))}`);
    
    return appToTeamMap;
    
  } catch (error) {
    debugLog.push(`[${new Date().toISOString()}] ⚠️ Exception building Application-to-Team mapping: ${error.message}`);
    return {};
  }
}

/**
 * Build reverse mapping: Application Name → Project Prefix
 * This is used to check if Epics exist for impacted applications
 * NOW READS DYNAMICALLY FROM GOOGLE SHEET instead of hardcoded values
 */
function buildApplicationToProjectMapping(debugLog) {
  try {
    debugLog.push(`[${new Date().toISOString()}] 📊 Building Application-to-Project mapping from Google Sheet...`);
    
    const sheetResult = fetchJiraOverviewData();
    
    if (!sheetResult.success) {
      debugLog.push(`[${new Date().toISOString()}] ⚠️ Failed to fetch Google Sheet: ${sheetResult.message}`);
      debugLog.push(`[${new Date().toISOString()}] ⚠️ Will use empty mapping`);
      return {}; // Return empty map - will show warnings for missing mappings
    }
    
    const appToProjectMap = {};
    const data = sheetResult.data || [];
    
    debugLog.push(`[${new Date().toISOString()}] 📊 Processing ${data.length} rows from Google Sheet...`);
    
    // Build mapping: Extract application names → project prefix
    data.forEach(row => {
      const applications = row['Application'] || '';  // Changed from 'Application(s)' to 'Application'
      const jiraKey = row['Jira Issue Key'] || '';
      
      if (applications && jiraKey) {
        // Extract project prefix (e.g., 'C3X' from 'C3X-123')
        const prefix = jiraKey.split('-')[0];
        
        // Split applications by comma (in case multiple apps listed)
        const appList = applications.split(',').map(app => app.trim());
        
        appList.forEach(appName => {
          if (appName && prefix) {
            appToProjectMap[appName] = prefix;
            debugLog.push(`[${new Date().toISOString()}]    📋 Mapped: "${appName}" → "${prefix}"`);
          }
        });
      }
    });
    
    debugLog.push(`[${new Date().toISOString()}] ✅ Built application mapping with ${Object.keys(appToProjectMap).length} applications`);
    debugLog.push(`[${new Date().toISOString()}] 📋 Sample mappings: ${JSON.stringify(Object.fromEntries(Object.entries(appToProjectMap).slice(0, 5)))}`);
    
    return appToProjectMap;
    
  } catch (error) {
    debugLog.push(`[${new Date().toISOString()}] ⚠️ Exception building application mapping: ${error.message}`);
    debugLog.push(`[${new Date().toISOString()}] ⚠️ Will use empty mapping`);
    return {}; // Return empty map on error
  }
}

/**
 * Process GREEN data and extract custom fields
 * applicationToTeamMapping: fallback map of Application Name → Team Name
 * (used for GREEN-board Epics whose project prefix is not in teamMapping)
 */
function processGreenData(issues, teamMapping, leadershipMapping, applicationToTeamMapping, debugLog, cachedSheetData) {
  // Handle backward-compat: if applicationToTeamMapping is actually the debugLog (old callers)
  if (Array.isArray(applicationToTeamMapping)) {
    debugLog = applicationToTeamMapping;
    applicationToTeamMapping = {};
  }
  applicationToTeamMapping = applicationToTeamMapping || {};
  cachedSheetData = cachedSheetData || null;
  debugLog.push(`[${new Date().toISOString()}] 📊 Processing ${issues.length} issues...`);
  debugLog.push(`[${new Date().toISOString()}] 📊 Team mapping has ${Object.keys(teamMapping).length} entries`);
  debugLog.push(`[${new Date().toISOString()}] 📊 Leadership mapping has ${Object.keys(leadershipMapping).length} entries`);
  debugLog.push(`[${new Date().toISOString()}] 📊 Application-to-Team mapping has ${Object.keys(applicationToTeamMapping).length} entries`);
  
  const items = [];
  let ideaCount = 0;
  let capabilityCount = 0;
  let epicCount = 0;
  let storyCount = 0;
  let taskCount = 0;
  
  // Capability status breakdown
  let capabilityInProgress = 0;
  let capabilityReadyForTesting = 0;
  let capabilityCioReady = 0;
  let capabilityDone = 0;
  let capabilityBlocked = 0;
  let capabilityRefinement = 0;
  
  // Epic status breakdown
  let epicBacklog = 0;
  let epicInProgress = 0;
  let epicDone = 0;
  let epicBlocked = 0;
  let epicRestOfStatus = 0;
  
  // Story status breakdown
  let storyBacklog = 0;
  let storyInProgress = 0;
  let storyDone = 0;
  let storyBlocked = 0;
  let storyRestOfStatus = 0;
  
  issues.forEach(issue => {
    const fields = issue.fields;
    const issueType = fields.issuetype ? fields.issuetype.name : 'Unknown';
    const status = fields.status ? fields.status.name : 'Unknown';
    const statusLower = status.toLowerCase();
    
    // Count by type and status
    if (issueType === 'Idea' || issueType === 'Initiative') {
      ideaCount++;
    } else if (issueType === 'Capability') {
      capabilityCount++;
      
      // Count capability by status
      if (statusLower.includes('in progress')) {
        capabilityInProgress++;
      } else if (statusLower.includes('ready for testing')) {
        capabilityReadyForTesting++;
      } else if (statusLower.includes('cio ready')) {
        capabilityCioReady++;
      } else if (statusLower.includes('done')) {
        capabilityDone++;
      } else if (statusLower.includes('blocked')) {
        capabilityBlocked++;
      } else if (statusLower.includes('refinement')) {
        capabilityRefinement++;
      }
    } else if (issueType === 'Epic') {
      epicCount++;
      // Count epic by status
      if (statusLower.includes('backlog')) {
        epicBacklog++;
      } else if (statusLower.includes('in progress') || 
          statusLower.includes('ready for testing') || 
          statusLower.includes('cio ready')) {
        epicInProgress++;
      } else if (statusLower.includes('done')) {
        epicDone++;
      } else if (statusLower.includes('blocked')) {
        epicBlocked++;
      } else {
        // All other statuses go here
        epicRestOfStatus++;
      }
    } else if (issueType === 'Story') {
      storyCount++;
      // Count story by status
      if (statusLower.includes('backlog')) {
        storyBacklog++;
      } else if (statusLower.includes('in progress') || 
          statusLower.includes('ready for testing') || 
          statusLower.includes('cio ready')) {
        storyInProgress++;
      } else if (statusLower.includes('done')) {
        storyDone++;
      } else if (statusLower.includes('blocked')) {
        storyBlocked++;
      } else {
        // All other statuses go here
        storyRestOfStatus++;
      }
    } else if (issueType.includes('Task')) {
      taskCount++;
    }
    
    // Extract Health (customfield_19357)
    let health = 'N/A';
    if (fields.customfield_19357) {
      if (typeof fields.customfield_19357 === 'object' && fields.customfield_19357.value) {
        health = fields.customfield_19357.value;
      } else if (typeof fields.customfield_19357 === 'string') {
        health = fields.customfield_19357;
      }
      
      // Convert emoji to text
      if (health.includes('🔴') || health.toLowerCase().includes('red')) {
        health = 'Red';
      } else if (health.includes('🟡') || health.toLowerCase().includes('yellow')) {
        health = 'Yellow';
      } else if (health.includes('🟢') || health.toLowerCase().includes('green')) {
        health = 'Green';
      } else if (health.includes('⚪') || health.toLowerCase().includes('grey') || health.toLowerCase().includes('gray')) {
        health = 'Grey';
      }
    }
    
    // Extract Gate 2 - Cost Estimate (customfield_17265)
    let gate2Cost = null;
    if (fields.customfield_17265 !== null && fields.customfield_17265 !== undefined) {
      gate2Cost = parseFloat(fields.customfield_17265);
      if (isNaN(gate2Cost)) {
        gate2Cost = null;
      }
    }
    
    // Extract Person Days (customfield_24241)
    let personDays = null;
    if (fields.customfield_24241 !== null && fields.customfield_24241 !== undefined) {
      personDays = parseFloat(fields.customfield_24241);
      if (isNaN(personDays)) {
        personDays = null;
      }
    }
    
    // Extract GTM Date (customfield_21190)
    let gtmDate = 'Not set';
    if (fields.customfield_21190) {
      gtmDate = fields.customfield_21190;
    }
    
    // Extract Dev Start Date (customfield_16070)
    let devStartDate = 'Not set';
    if (fields.customfield_16070) {
      devStartDate = fields.customfield_16070;
    }
    
    // Extract E2E Test Ready Date / Dev Completion (customfield_24242)
    let devCompletionDate = 'Not set';
    if (fields.customfield_24242) {
      devCompletionDate = fields.customfield_24242;
    }
    
    // Extract UAT Start Date (customfield_21050)
    let uatStartDate = 'Not set';
    if (fields.customfield_21050) {
      uatStartDate = fields.customfield_21050;
    }
    
    // Extract UAT End Date (customfield_24243)
    let uatEndDate = 'Not set';
    if (fields.customfield_24243) {
      uatEndDate = fields.customfield_24243;
    }
    
    // Extract Impacted Digital Products (customfield_26082)
    let impactedProducts = [];
    if (fields.customfield_26082 && Array.isArray(fields.customfield_26082)) {
      impactedProducts = fields.customfield_26082.map(product => {
        if (typeof product === 'object' && product.value) {
          return product.value;
        }
        return product;
      });
    }
    
    // Extract Impacted Applications (customfield_30833)
    let impactedApplications = [];
    if (fields.customfield_30833 && Array.isArray(fields.customfield_30833)) {
      impactedApplications = fields.customfield_30833.map(app => {
        if (typeof app === 'object' && app.value) return app.value.trim();
        return String(app).trim();
      }).filter(app => app);
    }
    
    // Extract parent
    let parent = 'None';
    if (fields.parent && fields.parent.key) {
      parent = fields.parent.key;
    }
    
    // Extract labels
    let labels = [];
    if (fields.labels && Array.isArray(fields.labels)) {
      labels = fields.labels;
    }
    
    // Extract Status/Risk (customfield_19361)
    // This field uses Atlassian Document Format (ADF) - need to extract text from JSON structure
    let statusRisk = 'N/A';
    if (fields.customfield_19361) {
      if (typeof fields.customfield_19361 === 'string') {
        statusRisk = fields.customfield_19361;
      } else if (typeof fields.customfield_19361 === 'object') {
        // Handle Atlassian Document Format (ADF)
        try {
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
            // Add line break after paragraphs
            if (node.type === 'paragraph' && text) {
              text += '\n';
            }
            return text;
          };
          
          statusRisk = extractTextFromADF(fields.customfield_19361).trim();
          if (!statusRisk) {
            statusRisk = 'N/A';
          }
        } catch (e) {
          statusRisk = 'N/A';
        }
      }
    }
    
    // Extract T-shirt (customfield_19257)
    let tshirt = 'N/A';
    if (fields.customfield_19257) {
      if (typeof fields.customfield_19257 === 'object' && fields.customfield_19257.value) {
        tshirt = fields.customfield_19257.value;
      } else if (typeof fields.customfield_19257 === 'string') {
        tshirt = fields.customfield_19257;
      }
    }
    
    // Extract description
    let description = '';
    if (fields.description) {
      description = fields.description;
    }
    
    // Derive team from project key for Epics and Stories only
    // Capabilities, Tasks, and Initiatives get 'N/A' (no meaningful team association)
    let team = 'N/A';
    let director = 'N/A';
    let manager = 'N/A';
    
    if (issueType === 'Epic' || issueType === 'Story') {
      const prefix = issue.key.split('-')[0];
      if (teamMapping[prefix]) {
        team = teamMapping[prefix];
        if (leadershipMapping[team]) {
          director = leadershipMapping[team].director;
          manager = leadershipMapping[team].manager;
        } else {
          director = 'Unknown';
          manager = 'Unknown';
        }
      } else {
        let resolvedTeam = '';
        for (let i = 0; i < impactedApplications.length; i++) {
          if (applicationToTeamMapping[impactedApplications[i]]) {
            resolvedTeam = applicationToTeamMapping[impactedApplications[i]];
            break;
          }
        }
        team = resolvedTeam || 'Unknown';
        director = 'Unknown';
        manager = 'Unknown';
        if (resolvedTeam && leadershipMapping[resolvedTeam]) {
          director = leadershipMapping[resolvedTeam].director;
          manager = leadershipMapping[resolvedTeam].manager;
        }
      }
    }
    
    // Extract Jira Priority (standard field) — used for Experience Defects
    let jiraPriority = 'N/A';
    if (fields.priority) {
      if (typeof fields.priority === 'object' && fields.priority.name) {
        jiraPriority = fields.priority.name;
      } else if (typeof fields.priority === 'string') {
        jiraPriority = fields.priority;
      }
    }

    // Extract Severity of Defect (customfield_19632) — Experience Defects
    let severityOfDefect = 'N/A';
    if (fields.customfield_19632) {
      if (typeof fields.customfield_19632 === 'object' && fields.customfield_19632.value) {
        severityOfDefect = fields.customfield_19632.value.trim();
      } else if (typeof fields.customfield_19632 === 'string') {
        severityOfDefect = fields.customfield_19632.trim();
      }
    }

    const item = {
      key: issue.key,
      summary: fields.summary || '',
      description: description,
      issueType: issueType,
      parent: parent,
      health: health,
      gate2Cost: gate2Cost,
      personDays: personDays,
      gtmDate: gtmDate,
      devStartDate: devStartDate,
      devCompletionDate: devCompletionDate,
      uatStartDate: uatStartDate,
      uatEndDate: uatEndDate,
      impactedProducts: impactedProducts,
      impactedApplications: impactedApplications,
      status: fields.status ? fields.status.name : 'Unknown',
      assignee: fields.assignee ? fields.assignee.displayName : 'Unassigned',
      created: fields.created || '',
      dueDate: fields.duedate || 'Not set',
      labels: labels,
      'Status/Risk': statusRisk,
      tshirt: tshirt,
      team: team,
      director: director,
      manager: manager,
      jiraPriority: jiraPriority,
      severityOfDefect: severityOfDefect
    };
    
    items.push(item);
  });
  
  // Calculate summary statistics
  const summary = {
    ideaCount: ideaCount,
    capabilityCount: capabilityCount,
    capabilityInProgress: capabilityInProgress,
    capabilityReadyForTesting: capabilityReadyForTesting,
    capabilityCioReady: capabilityCioReady,
    capabilityDone: capabilityDone,
    capabilityBlocked: capabilityBlocked,
    capabilityRefinement: capabilityRefinement,
    epicCount: epicCount,
    epicBacklog: epicBacklog,
    epicInProgress: epicInProgress,
    epicDone: epicDone,
    epicBlocked: epicBlocked,
    epicRestOfStatus: epicRestOfStatus,
    storyCount: storyCount,
    storyBacklog: storyBacklog,
    storyInProgress: storyInProgress,
    storyDone: storyDone,
    storyBlocked: storyBlocked,
    storyRestOfStatus: storyRestOfStatus,
    taskCount: taskCount,
    totalItems: items.length
  };
  
  debugLog.push(`[${new Date().toISOString()}] 📊 Summary: ${ideaCount} ideas, ${capabilityCount} capabilities (IP:${capabilityInProgress}, RFT:${capabilityReadyForTesting}, CIO:${capabilityCioReady}, Done:${capabilityDone}, Blocked:${capabilityBlocked}, Refinement:${capabilityRefinement}), ${epicCount} epics (Backlog:${epicBacklog}, IP:${epicInProgress}, Done:${epicDone}, Blocked:${epicBlocked}, Other:${epicRestOfStatus}), ${storyCount} stories (Backlog:${storyBacklog}, IP:${storyInProgress}, Done:${storyDone}, Blocked:${storyBlocked}, Other:${storyRestOfStatus}), ${taskCount} tasks`);
  
  // Log Impacted Applications summary
  debugLog.push(`[${new Date().toISOString()}] 📊 IMPACTED APPLICATIONS SUMMARY:`);
  const appCounts = {};
  let itemsWithApps = 0;
  
  items.forEach(item => {
    if (item.impactedApplications && item.impactedApplications.length > 0) {
      itemsWithApps++;
      item.impactedApplications.forEach(app => {
        appCounts[app] = (appCounts[app] || 0) + 1;
      });
    }
  });
  
  debugLog.push(`[${new Date().toISOString()}]    Items with applications: ${itemsWithApps} / ${items.length}`);
  debugLog.push(`[${new Date().toISOString()}]    Total unique applications: ${Object.keys(appCounts).length}`);
  
  if (Object.keys(appCounts).length > 0) {
    Object.entries(appCounts).sort((a, b) => b[1] - a[1]).forEach(([app, count]) => {
      debugLog.push(`[${new Date().toISOString()}]       • ${app}: ${count} issue(s)`);
    });
  } else {
    debugLog.push(`[${new Date().toISOString()}]    ⚠️ No applications found in any issues`);
  }
  
  // Calculate "Epics Missing" for each Capability
  debugLog.push(`[${new Date().toISOString()}] 📊 CALCULATING EPICS MISSING FOR CAPABILITIES...`);
  
  // Build Application → Project Prefix mapping — use cached sheet data if available (avoids 4th sheet read)
  const appToProjectMap = cachedSheetData
    ? buildApplicationToProjectMappingFromCache(cachedSheetData, debugLog)
    : buildApplicationToProjectMapping(debugLog);
  debugLog.push(`[${new Date().toISOString()}]    Application-to-Project mapping has ${Object.keys(appToProjectMap).length} entries`);
  
  // Get all Capabilities and Epics
  const capabilities = items.filter(item => item.issueType === 'Capability');
  const epics = items.filter(item => item.issueType === 'Epic');
  
  debugLog.push(`[${new Date().toISOString()}]    Found ${capabilities.length} Capabilities to check`);
  debugLog.push(`[${new Date().toISOString()}]    Found ${epics.length} Epics total`);
  
  let totalMissingEpics = 0;
  
  capabilities.forEach(capability => {
    const impactedApps = capability.impactedApplications || [];
    const missingEpics = [];
    const childEpics = epics.filter(epic => epic.parent === capability.key);
    const greenChildEpics = childEpics.filter(epic => epic.key.startsWith('GREEN-'));

    impactedApps.forEach(appName => {
      const expectedPrefix = appToProjectMap[appName];
      if (!expectedPrefix) {
        const covered = greenChildEpics.some(ge =>
          (ge.impactedApplications || []).some(a => a.trim().toLowerCase() === appName.trim().toLowerCase())
        );
        if (!covered) {
          missingEpics.push({ team: 'Unknown', appName, epicKey: 'N/A', message: `${appName} for ${capability.key}` });
        }
        return;
      }
      const hasEpic = childEpics.some(epic => epic.key.startsWith(expectedPrefix + '-'));
      if (!hasEpic) {
        const covered = greenChildEpics.some(ge =>
          (ge.impactedApplications || []).some(a => a.trim().toLowerCase() === appName.trim().toLowerCase())
        );
        if (!covered) {
          const teamName = teamMapping[expectedPrefix] || 'Unknown';
          missingEpics.push({ team: teamName, appName, epicKey: expectedPrefix + '-', message: `${appName} for ${capability.key}` });
          totalMissingEpics++;
        }
      }
    });

    capability.epicsMissing = missingEpics;
  });
  
  debugLog.push(`[${new Date().toISOString()}] 📊 EPICS MISSING SUMMARY:`);
  debugLog.push(`[${new Date().toISOString()}]    Total missing Epics across all Capabilities: ${totalMissingEpics}`);
  debugLog.push(`[${new Date().toISOString()}]    Capabilities with missing Epics: ${capabilities.filter(c => c.epicsMissing && c.epicsMissing.length > 0).length} / ${capabilities.length}`);
  
  return {
    items: items,
    summary: summary
  };
}

/**
 * TEST FUNCTION - Force authorization for Google Sheets access
 * Run this once in Apps Script editor to grant permissions
 */
function TEST_AUTHORIZE_SHEETS() {
  try {
    Logger.log('🔐 Testing Google Sheets authorization...');
    
    // This will trigger the authorization prompt
    const spreadsheetId = JIRA_OVERVIEW_CONFIG.spreadsheetId;
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(JIRA_OVERVIEW_CONFIG.sheetName);
    
    if (sheet) {
      const rowCount = sheet.getLastRow();
      Logger.log('✅ Successfully authorized and accessed sheet: ' + JIRA_OVERVIEW_CONFIG.sheetName);
      Logger.log('✅ Sheet has ' + rowCount + ' rows');
      Logger.log('✅ Authorization complete! You can now use the Jira Overview feature.');
      return 'Authorization successful! Sheet has ' + rowCount + ' rows.';
    } else {
      Logger.log('❌ Sheet not found: ' + JIRA_OVERVIEW_CONFIG.sheetName);
      return 'Sheet not found: ' + JIRA_OVERVIEW_CONFIG.sheetName;
    }
  } catch (error) {
    Logger.log('❌ Error: ' + error.message);
    Logger.log('💡 If you see a permission error, click "Review Permissions" and authorize the script.');
    return 'Error: ' + error.message;
  }
}

/**
 * TEST FUNCTION - Verify fetchJiraOverviewData works after authorization
 */
function TEST_FETCH_JIRA_OVERVIEW() {
  Logger.log('🧪 Testing fetchJiraOverviewData function...');
  
  const result = fetchJiraOverviewData();
  
  Logger.log('Success: ' + result.success);
  Logger.log('Message: ' + result.message);
  
  if (result.success) {
    Logger.log('✅ Data rows fetched: ' + result.data.length);
    Logger.log('✅ Headers: ' + result.headers.join(', '));
    if (result.data.length > 0) {
      Logger.log('✅ Sample row: ' + JSON.stringify(result.data[0]));
    }
  } else {
    Logger.log('❌ Error: ' + result.message);
  }
  
  return result;
}

/**
 * Fetch Jira Overview data from Google Sheet
 * Reads organizational/team structure data from the Registry sheet
 * Called from GREEN.html when "Green Engineering - Jira Overview" button is clicked
 */
function fetchJiraOverviewData() {
  try {
    // Get configuration
    const spreadsheetId = JIRA_OVERVIEW_CONFIG.spreadsheetId;
    const sheetName = JIRA_OVERVIEW_CONFIG.sheetName;
    
    // Open the spreadsheet
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      return {
        success: false,
        message: `Sheet "${sheetName}" not found in spreadsheet`
      };
    }
    
    // Get all data from the sheet
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    Logger.log('DEBUG: Total rows in sheet: ' + values.length);
    
    if (values.length < 2) {
      return {
        success: false,
        message: 'Not enough data in sheet (need at least 2 rows)'
      };
    }
    
    // Row 1 is the title row (merged cells) - skip it
    // Row 2 contains the actual column headers
    const headers = values[1];
    Logger.log('DEBUG: Headers found (from row 2): ' + JSON.stringify(headers));
    Logger.log('DEBUG: Number of headers: ' + headers.length);
    
    const data = [];
    
    // Process each row starting from row 3 (index 2)
    for (let i = 2; i < values.length; i++) {
      const row = values[i];
      const rowData = {};
      
      // Map each column to its header
      for (let j = 0; j < headers.length; j++) {
        const headerName = headers[j];
        const cellValue = row[j] || '';
        rowData[headerName] = cellValue;
      }
      
      // Log first data row for debugging
      if (i === 2) {
        Logger.log('DEBUG: First data row object: ' + JSON.stringify(rowData));
        Logger.log('DEBUG: First data row object keys: ' + Object.keys(rowData).join(', '));
      }
      
      // Filter out completely empty rows (all cells are empty)
      const isEmptyRow = Object.values(rowData).every(value => 
        value === '' || value === null || value === undefined
      );
      
      // Only add non-empty rows
      if (!isEmptyRow) {
        data.push(rowData);
      }
    }
    
    Logger.log('DEBUG: Total data rows created: ' + data.length);
    
    return {
      success: true,
      message: `Successfully fetched ${data.length} rows`,
      data: data,
      headers: headers
    };
    
  } catch (error) {
    Logger.log('ERROR in fetchJiraOverviewData: ' + error.message);
    Logger.log('ERROR stack: ' + error.stack);
    return {
      success: false,
      message: 'Error fetching Jira Overview data: ' + error.message,
      error: error.toString()
    };
  }
}
/**
 * DEBUG FUNCTION - Diagnose why large label (Optik_It1.2) fails for GMAP-5
 * Run this in Apps Script editor and paste the full log output back to Cline
 */
function DEBUG_LARGE_LABEL() {
  Logger.clear();
  Logger.log('='.repeat(80));
  Logger.log('🔬 DEBUG: Large Label Failure Diagnosis - GMAP-5 / Optik_It1.2');
  Logger.log('='.repeat(80));
  
  var startTime = new Date();
  
  var props = PropertiesService.getScriptProperties();
  var apiToken = props.getProperty('JIRA_API_TOKEN');
  var cloudId = props.getProperty('JIRA_CLOUD_ID');
  var baseUrl = 'https://api.atlassian.com/ex/jira/' + cloudId.trim();
  var authHeader = 'Bearer ' + apiToken.trim();
  
  var ideaKey = 'GMAP-5';
  var failingLabel = 'Optik_It1.2';
  
  function elapsed() {
    return ((new Date() - startTime) / 1000).toFixed(1) + 's';
  }
  
  function jqlPost(label, jql, maxResults) {
    maxResults = maxResults || 1000;
    Logger.log('\n[' + elapsed() + '] 📋 ' + label);
    Logger.log('   JQL: ' + jql);
    Logger.log('   JQL length: ' + jql.length + ' chars');
    try {
      var r = UrlFetchApp.fetch(baseUrl + '/rest/api/3/search/jql', {
        method: 'POST',
        headers: {
          'Authorization': authHeader,
          'Accept': 'application/json',
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
          jql: jql,
          maxResults: maxResults,
          fields: ['key', 'summary', 'issuetype', 'status', 'parent', 'labels']
        }),
        muteHttpExceptions: true
      });
      var code = r.getResponseCode();
      var body = r.getContentText();
      Logger.log('   [' + elapsed() + '] Response Code: ' + code);
      if (code === 200) {
        var d = JSON.parse(body);
        Logger.log('   ✅ Total in JIRA: ' + d.total + ' | Returned in this call: ' + (d.issues ? d.issues.length : 0));
        Logger.log('   isLast: ' + d.isLast + ' | startAt: ' + d.startAt);
        if (d.total > 1000) {
          Logger.log('   ⚠️ WARNING: More than 1000 results! Only first 1000 returned. Need pagination!');
        }
        return { success: true, total: d.total, count: d.issues ? d.issues.length : 0, issues: d.issues || [], isLast: d.isLast };
      } else {
        Logger.log('   ❌ FAILED Code=' + code + ': ' + body.substring(0, 500));
        return { success: false, total: 0, count: 0, issues: [], code: code, error: body.substring(0, 500) };
      }
    } catch (e) {
      Logger.log('   💥 EXCEPTION: ' + e.message);
      return { success: false, total: 0, count: 0, issues: [], error: e.message };
    }
  }
  
  // STEP 1: Auth check
  Logger.log('\n[' + elapsed() + '] ── STEP 1: Auth check ──');
  var authCheck = jqlPost('Auth check (key = GMAP-5)', 'key = GMAP-5', 1);
  if (!authCheck.success) {
    Logger.log('❌ Auth failed - stopping. Check JIRA_API_TOKEN in Script Properties.');
    return;
  }
  Logger.log('✅ Auth OK');
  
  // STEP 2: Count capabilities for this label
  Logger.log('\n[' + elapsed() + '] ── STEP 2: Capabilities for label ' + failingLabel + ' ──');
  var capResult = jqlPost(
    'Capabilities with label ' + failingLabel,
    'project = GREEN AND type = Capability AND issue in linkedIssues(' + ideaKey + ') AND labels in ("' + failingLabel + '") ORDER BY created DESC',
    1000
  );
  Logger.log('   Capabilities found: ' + capResult.total);
  
  if (!capResult.success) {
    Logger.log('❌ Capability fetch FAILED - this is the failure point!');
    return;
  }
  if (capResult.total === 0) {
    Logger.log('⚠️ 0 capabilities found for this label - label may not exist or no capabilities tagged with it');
    return;
  }
  
  var capKeys = capResult.issues.map(function(i) { return i.key; });
  Logger.log('   Capability keys: ' + capKeys.join(', '));
  
  // STEP 3: Count child Epics/Tasks
  Logger.log('\n[' + elapsed() + '] ── STEP 3: Child Epics/Tasks for ' + capKeys.length + ' capabilities ──');
  var childJQL = 'parent in (' + capKeys.map(function(k) { return '"' + k + '"'; }).join(',') + ') ORDER BY created DESC';
  var childResult = jqlPost('Children of capabilities', childJQL, 1000);
  Logger.log('   Children found: ' + childResult.total);
  
  if (!childResult.success) {
    Logger.log('❌ Child fetch FAILED - this is the failure point!');
    return;
  }
  
  // STEP 4: Identify Epics
  var epicKeys = childResult.issues
    .filter(function(i) { return i.fields.issuetype && i.fields.issuetype.name === 'Epic'; })
    .map(function(i) { return i.key; });
  var taskCount = childResult.issues.length - epicKeys.length;
  Logger.log('\n[' + elapsed() + '] ── STEP 4: Epic/Task breakdown ──');
  Logger.log('   Epics: ' + epicKeys.length);
  Logger.log('   Tasks/Other: ' + taskCount);
  Logger.log('   Epic keys (first 20): ' + epicKeys.slice(0, 20).join(', '));
  
  if (epicKeys.length === 0) {
    Logger.log('ℹ️ No Epics found - no Story fetch needed. Data should load fine.');
    Logger.log('✅ No obvious failure point found. Issue may be in data processing, not fetching.');
  } else {
    // STEP 5: Check Story JQL size and count
    Logger.log('\n[' + elapsed() + '] ── STEP 5: Stories under ' + epicKeys.length + ' Epics ──');
    var storyJQL = 'parent in (' + epicKeys.map(function(k) { return '"' + k + '"'; }).join(',') + ') ORDER BY created DESC';
    Logger.log('   Story JQL length: ' + storyJQL.length + ' chars');
    
    if (storyJQL.length > 30000) {
      Logger.log('   🚨 JQL TOO LONG (' + storyJQL.length + ' chars > 30000 limit)!');
      Logger.log('   🚨 THIS IS THE FAILURE POINT - need to batch epic keys into chunks');
    } else {
      var storyResult = jqlPost('Stories under Epics', storyJQL, 1000);
      Logger.log('   Stories found: ' + storyResult.total);
      
      if (!storyResult.success) {
        Logger.log('❌ Story fetch FAILED - this is the failure point!');
        Logger.log('   Error: ' + storyResult.error);
      } else if (storyResult.total > 1000) {
        Logger.log('🚨 MORE THAN 1000 STORIES (' + storyResult.total + ')!');
        Logger.log('🚨 THIS IS THE FAILURE POINT - only got ' + storyResult.count + ' of ' + storyResult.total + ' stories');
        Logger.log('🚨 Need pagination to fetch all stories in batches');
      } else {
        Logger.log('✅ Story fetch OK - ' + storyResult.total + ' stories returned');
      }
    }
  }
  
  Logger.log('\n' + '='.repeat(80));
  Logger.log('📊 FINAL SUMMARY');
  Logger.log('   Total elapsed: ' + elapsed());
  Logger.log('   Capabilities: ' + capResult.total);
  Logger.log('   Children (Epics+Tasks): ' + childResult.total);
  Logger.log('   Epics: ' + epicKeys.length);
  Logger.log('='.repeat(80));
}

/**
 * Fetches Executive Summary data (Goals, Risks, Decisions) from a configured Google Sheet
 * for the given idea key. Sheet URL is stored in Script Property GREEN_EXEC_SUMMARY_SHEETS.
 * Format: GMAP-5:https://docs.google.com/spreadsheets/d/ID/edit;GMAP-14:https://...
 */
function fetchExecSummaryForIdea(ideaKey) {
  try {
    Logger.log('📊 fetchExecSummaryForIdea called for: ' + ideaKey);
    
    var mapping = PropertiesService.getScriptProperties().getProperty('GREEN_EXEC_SUMMARY_SHEETS');
    
    if (!mapping) {
      Logger.log('⚠️ GREEN_EXEC_SUMMARY_SHEETS not set in Script Properties');
      return { success: false, message: 'No sheet mapping configured' };
    }
    
    Logger.log('📋 Raw mapping: ' + mapping.substring(0, 100));
    
    // Parse entries separated by semicolons
    var entries = mapping.split(';').map(function(e) { return e.trim(); }).filter(function(e) { return e.length > 0; });
    Logger.log('📋 Found ' + entries.length + ' entries in mapping');
    
    // Find the URL for this idea key
    var sheetUrl = null;
    for (var i = 0; i < entries.length; i++) {
      var colonIdx = entries[i].indexOf(':');
      if (colonIdx === -1) continue;
      var key = entries[i].substring(0, colonIdx).trim();
      var url = entries[i].substring(colonIdx + 1).trim();
      Logger.log('  Comparing key [' + key + '] with ideaKey [' + ideaKey + ']');
      if (key.toLowerCase() === ideaKey.toLowerCase()) {
        sheetUrl = url;
        Logger.log('  ✅ Match found! URL: ' + url.substring(0, 60));
        break;
      }
    }
    
    if (!sheetUrl) {
      Logger.log('ℹ️ No sheet configured for idea: ' + ideaKey);
      return { success: false, message: 'No sheet configured for ' + ideaKey };
    }
    
    // Open the spreadsheet
    Logger.log('📂 Opening spreadsheet...');
    var ss = SpreadsheetApp.openByUrl(sheetUrl);
    Logger.log('✅ Spreadsheet opened: ' + ss.getName());
    
    var goals = '';
    var risks = '';
    var decisions = '';
    var overallStatus = '';
    
    // Read Goals from "Automated Project Goal" tab
    var goalsSheet = ss.getSheetByName('Automated Project Goal');
    if (goalsSheet) {
      Logger.log('📄 Reading Automated Project Goal tab...');
      var goalsData = goalsSheet.getDataRange().getValues();
      var goalLines = [];
      for (var r = 1; r < goalsData.length; r++) { // skip header row
        var colB = (goalsData[r][1] || '').toString().trim();
        var colC = (goalsData[r][2] || '').toString().trim();
        if (colB) goalLines.push(colB);
        if (colC && !overallStatus) overallStatus = colC; // first non-empty C = overall status
      }
      goals = goalLines.join('\n');
      Logger.log('✅ Goals read: ' + goals.length + ' chars, status: ' + overallStatus);
    } else {
      Logger.log('⚠️ Tab "Automated Project Goal" not found');
    }
    
    // Read Risks and Decisions from "Automated Risks" tab
    var risksSheet = ss.getSheetByName('Automated Risks');
    if (risksSheet) {
      Logger.log('📄 Reading Automated Risks tab...');
      var risksData = risksSheet.getDataRange().getValues();
      var riskLines = [];
      var decisionLines = [];
      for (var r = 1; r < risksData.length; r++) { // skip header row
        var status = (risksData[r][0] || '').toString().trim().toLowerCase();
        if (status === 'closed') continue; // skip closed items
        var colB = (risksData[r][1] || '').toString().trim();
        var colC = (risksData[r][2] || '').toString().trim();
        var colF = (risksData[r][5] || '').toString().trim(); // "Ask" column
        if (colB || colC) riskLines.push([colB, colC].filter(Boolean).join(' — '));
        if (colF) decisionLines.push(colF);
      }
      risks = riskLines.join('\n');
      decisions = decisionLines.join('\n');
      Logger.log('✅ Risks read: ' + risks.length + ' chars, Decisions: ' + decisions.length + ' chars');
    } else {
      Logger.log('⚠️ Tab "Automated Risks" not found');
    }
    
    Logger.log('✅ fetchExecSummaryForIdea complete for ' + ideaKey);
    return {
      success: true,
      goals: goals,
      risks: risks,
      decisions: decisions,
      overallStatus: overallStatus
    };
    
  } catch (e) {
    Logger.log('❌ Error in fetchExecSummaryForIdea: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// DEFECTS — Backend functions (called from GREEN.html via google.script.run)
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Fetch Experience Defects for a selected Idea (GMAP key)
 * Called from GREEN.html via google.script.run.fetchDefectsForIdea(ideaKey)
 *
 * Logic:
 *   1. Find all Capabilities linked to the Idea (same JQL as fetchGreenData Step 2)
 *   2. Fetch all Experience Defects that are children of those Capabilities
 *   3. Extract the 12 fields needed for display + AI analysis
 *   4. Return array of defect objects to HTML
 */
function fetchDefectsForIdea(ideaKey) {
  const debugLog = [];

  try {
    debugLog.push(`[${new Date().toISOString()}] 🐛 fetchDefectsForIdea called for: ${ideaKey}`);

    if (!ideaKey) {
      return { success: false, message: 'No Idea key provided', data: [], debugLog };
    }

    // ── STEP 1: Get JIRA config (same pattern as fetchGreenData) ─────────────
    const config = getJiraConfig(debugLog);
    if (!config.success) return config;
    const { baseUrl, authHeader } = config;

    // ── STEP 2: Fetch Capabilities linked to this Idea ────────────────────────
    debugLog.push(`[${new Date().toISOString()}] 📋 STEP 2: Fetching Capabilities for ${ideaKey}...`);
    const capJQL = `project = GREEN AND type = Capability AND issue in linkedIssues(${ideaKey}) ORDER BY created DESC`;
    debugLog.push(`[${new Date().toISOString()}] 📝 Capability JQL: ${capJQL}`);

    const capResult = fetchJiraIssues(capJQL, baseUrl, authHeader, debugLog);
    if (!capResult.success || capResult.data.length === 0) {
      debugLog.push(`[${new Date().toISOString()}] ⚠️ No capabilities found for ${ideaKey}`);
      return { success: true, data: [], message: 'No capabilities found for this Idea', debugLog };
    }

    const capKeys = capResult.data.map(c => c.key);
    debugLog.push(`[${new Date().toISOString()}] ✅ Found ${capKeys.length} capabilities: ${capKeys.join(', ')}`);

    // ── STEP 3: Fetch Experience Defects under those Capabilities ─────────────
    debugLog.push(`[${new Date().toISOString()}] 📋 STEP 3: Fetching Experience Defects...`);

    const capKeysJQL = capKeys.map(k => `"${k}"`).join(',');
    const defectJQL  = `parent in (${capKeysJQL}) AND issuetype = "Experience Defect" ORDER BY priority DESC, created DESC`;
    debugLog.push(`[${new Date().toISOString()}] 📝 Defect JQL: ${defectJQL}`);

    const defectResult = fetchDefectIssues(defectJQL, baseUrl, authHeader, debugLog);
    if (!defectResult.success) return defectResult;

    const rawDefects = defectResult.data;
    debugLog.push(`[${new Date().toISOString()}] ✅ Found ${rawDefects.length} Experience Defects`);

    // ── STEP 4: Extract the fields we need ────────────────────────────────────
    const defects = rawDefects.map(issue => extractDefectFields(issue, debugLog));

    debugLog.push(`[${new Date().toISOString()}] 🎉 fetchDefectsForIdea complete — returning ${defects.length} defects`);

    return {
      success: true,
      message: `Found ${defects.length} Experience Defects`,
      data: defects,
      debugLog: debugLog.slice(-20)   // keep last 20 lines to avoid payload bloat
    };

  } catch (error) {
    debugLog.push(`[${new Date().toISOString()}] 💥 Exception: ${error.message}`);
    return { success: false, message: 'Exception: ' + error.message, data: [], debugLog };
  }
}


/**
 * Fetch defect issues using a targeted field list.
 * Separate from fetchJiraIssues() so we request only defect-relevant fields.
 */
function fetchDefectIssues(jql, baseUrl, authHeader, debugLog) {
  try {
    debugLog.push(`[${new Date().toISOString()}] 🔍 Fetching defects with targeted field list...`);

    const searchUrl = `${baseUrl}/rest/api/3/search/jql`;

    const requestBody = {
      jql: jql,
      maxResults: 500,
      fields: [
        'key', 'summary', 'description', 'issuetype', 'parent',
        'status', 'assignee', 'reporter', 'created', 'updated',
        'priority',
        'customfield_19632',   // Severity of Defect
        'customfield_12913',   // Defect Severity
        'customfield_12912',   // Defect Priority
        'customfield_18937',   // Application Name (Bug)
        'customfield_26316',   // Targeted Channels
        'customfield_19781',   // UI Identity team
        'customfield_19904',   // Environment
        'customfield_15601',   // Customer Facing
        'customfield_12400',   // Testing Environment
        'customfield_20660',   // Steps to reproduce
        'customfield_16631',   // Issue Description
        'labels'
      ]
    };

    const response = UrlFetchApp.fetch(searchUrl, {
      method: 'POST',
      headers: {
        'Authorization': authHeader,
        'Accept': 'application/json',
        'Content-Type': 'application/json',
        'User-Agent': 'GoogleAppsScript/1.0'
      },
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    debugLog.push(`[${new Date().toISOString()}] 📡 Defect fetch response code: ${code}`);

    if (code !== 200) {
      const errorText = response.getContentText().substring(0, 300);
      debugLog.push(`[${new Date().toISOString()}] ❌ Defect fetch failed: ${errorText}`);
      return { success: false, message: `JIRA API error ${code}: ${errorText}`, data: [] };
    }

    const data = JSON.parse(response.getContentText());
    debugLog.push(`[${new Date().toISOString()}] ✅ Defect fetch OK — total: ${data.total}, returned: ${(data.issues || []).length}`);

    return { success: true, data: data.issues || [] };

  } catch (error) {
    debugLog.push(`[${new Date().toISOString()}] 💥 fetchDefectIssues exception: ${error.message}`);
    return { success: false, message: error.message, data: [] };
  }
}


/**
 * Extract and normalise the fields we need from a raw Jira defect issue object.
 * Returns a clean flat object ready to be used in HTML table + AI prompt.
 */
function extractDefectFields(issue, debugLog) {
  const f = issue.fields;

  // ── Helper: extract plain text from Atlassian Document Format (ADF) ────────
  function adfToText(node) {
    if (!node) return '';
    if (typeof node === 'string') return node;
    let text = '';
    if (node.type === 'text') text += node.text || '';
    if (node.content && Array.isArray(node.content)) {
      node.content.forEach(child => { text += adfToText(child); });
    }
    if (node.type === 'paragraph' && text) text += '\n';
    return text;
  }

  // ── Helper: safely get a select-field value ──────────────────────────────
  function selectVal(field) {
    if (!field) return '';
    if (typeof field === 'string') return field.trim();
    if (field.value) return field.value.trim();
    return '';
  }

  // ── Helper: extract multi-select values as array ─────────────────────────
  function multiSelectVals(field) {
    if (!Array.isArray(field)) return [];
    return field.map(v => (typeof v === 'object' ? v.value || '' : String(v))).filter(Boolean);
  }

  // ── Severity — prefer customfield_19632, fallback to priority ────────────
  const severityRaw   = selectVal(f.customfield_19632) || selectVal(f.customfield_12913) || '';
  const priorityRaw   = selectVal(f.priority) || '';

  function normaliseSeverity(raw) {
    if (!raw) return '';
    const r = raw.toLowerCase();
    if (r.includes('critical'))                      return 'Critical';
    if (r.includes('high'))                          return 'High';
    if (r.includes('medium') || r.includes('moderate')) return 'Medium';
    if (r.includes('low'))                           return 'Low';
    return raw;
  }

  const severity = normaliseSeverity(severityRaw) || normaliseSeverity(priorityRaw) || 'Medium';

  // ── Description ───────────────────────────────────────────────────────────
  const desc1 = adfToText(f.customfield_20660).trim();
  const desc2 = adfToText(f.customfield_16631).trim();
  const descMain = adfToText(f.description).trim();

  let description = desc1 || desc2 || descMain || '';
  description = description
    .replace(/TODO/gi, '')
    .replace(/\*Test account\/data\/address\*/gi, '')
    .replace(/\*Steps to reproduce the issue\*/gi, '')
    .replace(/\*Actual results\*/gi, '')
    .trim()
    .substring(0, 600);

  // ── Application / Channel / Environment ──────────────────────────────────
  const appName        = selectVal(f.customfield_18937);
  const channels       = multiSelectVals(f.customfield_26316).join(', ');
  const uiTeamHint     = selectVal(f.customfield_19781);
  const environment    = selectVal(f.customfield_19904) || selectVal(f.customfield_12400);
  const customerFacing = selectVal(f.customfield_15601);
  const labels         = (f.labels || []).join(', ');

  return {
    key:          issue.key,
    summary:      f.summary || '',
    issueType:    f.issuetype ? f.issuetype.name : 'Experience Defect',
    parent:       f.parent ? f.parent.key : '',
    status:       f.status ? f.status.name : 'Unknown',
    assignee:     f.assignee ? f.assignee.displayName : '',
    reporter:     f.reporter ? f.reporter.displayName : '',
    created:      f.created ? f.created.substring(0, 10) : '',
    updated:      f.updated ? f.updated.substring(0, 10) : '',
    severity:     severity,
    severityRaw:  severityRaw,
    priority:     priorityRaw,
    app:          appName,
    channel:      channels,
    uiTeamHint:   uiTeamHint,
    environment:  environment,
    customerFacing: customerFacing,
    labels:       labels,
    description:  description
  };
}


/**
 * TEST FUNCTION — Run in Apps Script editor to verify defect fetch.
 * Change ideaKey to match a real idea in your GREEN_IDEAS_LIST.
 */
function TEST_FETCH_DEFECTS() {
  Logger.clear();
  Logger.log('='.repeat(80));
  Logger.log('🐛 TEST_FETCH_DEFECTS');
  Logger.log('='.repeat(80));

  const ideaKey = 'GMAP-5';   // ← change to your idea key

  const result = fetchDefectsForIdea(ideaKey);

  Logger.log('success: ' + result.success);
  Logger.log('message: ' + result.message);
  Logger.log('defect count: ' + (result.data ? result.data.length : 0));

  if (result.data && result.data.length > 0) {
    Logger.log('\nFirst defect:');
    Logger.log(JSON.stringify(result.data[0], null, 2));

    Logger.log('\nAll defect keys + severity + app:');
    result.data.forEach(d => {
      Logger.log(`  ${d.key} | ${d.severity} | ${d.status} | ${d.assignee || 'UNASSIGNED'} | App: ${d.app || '—'} | Channel: ${d.channel || '—'}`);
    });
  }

  if (result.debugLog) {
    Logger.log('\nDebug log:');
    result.debugLog.forEach(l => Logger.log(l));
  }

  Logger.log('='.repeat(80));
}
