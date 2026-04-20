// ===== SLACK INTEGRATION FUNCTIONS =====
// This file contains all Slack-related functionality for sending summaries
// Configuration is in Core.gs

/**
 * Get the list of configured Slack channels from Core.gs
 * Called from the frontend to populate the channel dropdown in the popup.
 * @returns {Array} Array of {id, name} channel objects
 */
function getSlackChannels() {
  if (!CONFIG.SLACK || !CONFIG.SLACK.channels || CONFIG.SLACK.channels.length === 0) {
    return [];
  }
  return CONFIG.SLACK.channels;
}

/**
 * Send summary to Slack
 * @param {string} summaryText - The formatted summary text
 * @param {string} summaryType - Either 'defects' or 'epics'
 * @param {string} channelId   - The Slack channel ID to send to
 * @returns {object} Success/failure status
 */
function sendToSlack(summaryText, summaryType, channelId) {
  try {
    console.log('=== SENDING TO SLACK ===');
    console.log(`Summary Type: ${summaryType}`);
    console.log(`Summary Length: ${summaryText ? summaryText.length : 0} characters`);
    
    // Validate inputs
    if (!summaryText || !summaryType) {
      throw new Error('Missing summary text or type');
    }
    
    // Resolve channel: use passed channelId, or fall back to first configured channel
    let resolvedChannelId = channelId;
    if (!resolvedChannelId) {
      if (!CONFIG.SLACK || !CONFIG.SLACK.channels || CONFIG.SLACK.channels.length === 0) {
        throw new Error('No Slack channels configured in Core.gs');
      }
      resolvedChannelId = CONFIG.SLACK.channels[0].id;
    }
    
    // Get Bot Token from Script Properties
    const botToken = getSlackBotToken();
    
    // Format message for Slack
    const slackMessage = formatSlackMessage(summaryText, summaryType);
    console.log('Slack message formatted successfully');
    
    // Send to Slack using Web API
    const slackApiUrl = 'https://slack.com/api/chat.postMessage';
    console.log(`Sending to channel: ${resolvedChannelId}`);
    
    const response = UrlFetchApp.fetch(slackApiUrl, {
      method: 'POST',
      contentType: 'application/json',
      headers: {
        'Authorization': `Bearer ${botToken}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        channel: resolvedChannelId,
        ...slackMessage
      }),
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    const responseData = JSON.parse(responseText);
    
    console.log(`Slack API Response Code: ${responseCode}`);
    console.log(`Slack API Response: ${responseText}`);
    
    if (responseCode === 200 && responseData.ok) {
      const successMessage = `${summaryType === 'defects' ? 'Defects' : 'Epic'} summary sent to Slack successfully!`;
      console.log(`✅ ${successMessage}`);
      return {
        success: true,
        message: successMessage
      };
    } else {
      const errorMsg = responseData.error || 'Unknown error';
      throw new Error(`Slack API error: ${errorMsg}`);
    }
    
  } catch (error) {
    console.error('Error sending to Slack:', error);
    return {
      success: false,
      message: `Failed to send to Slack: ${error.message}`
    };
  }
}

/**
 * Format message for Slack with proper formatting
 * Uses Slack Block Kit for rich formatting
 */
function formatSlackMessage(summaryText, summaryType) {
  const emoji = summaryType === 'defects' ? '🐛' : '📊';
  const title = summaryType === 'defects' ? 'DEFECTS STATUS REPORT' : summaryType === 'both' ? 'DEFECTS & WORK ITEMS STATUS REPORT' : 'WORK ITEMS STATUS REPORT';
  const timestamp = new Date().toLocaleString('en-US', { 
    timeZone: 'America/Vancouver',
    dateStyle: 'medium',
    timeStyle: 'short'
  });
  
  return {
    text: `${emoji} ${title}`,
    blocks: [
      {
        type: 'header',
        text: {
          type: 'plain_text',
          text: `${emoji} ${title}`,
          emoji: true
        }
      },
      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text: '```\n' + summaryText + '\n```'
        }
      },
      {
        type: 'context',
        elements: [
          {
            type: 'mrkdwn',
            text: `Sent from Optik on Green Dashboard | ${timestamp}`
          }
        ]
      }
    ]
  };
}

/**
 * Send summary to one or more Slack channels
 * Called from the frontend modal when user clicks "Send 🚀"
 * @param {string} summaryText  - The formatted summary text
 * @param {string} summaryType  - 'defects', 'epics', or 'both'
 * @param {Array}  channelIds   - Array of Slack channel IDs to send to
 * @returns {object} Success/failure status
 */
function sendToSlackWithOptions(summaryText, summaryType, channelIds) {
  try {
    console.log('=== SENDING TO SLACK (multi-channel) ===');
    console.log(`Summary Type: ${summaryType}`);
    console.log(`Channels: ${JSON.stringify(channelIds)}`);

    if (!summaryText || !summaryType) {
      throw new Error('Missing summary text or type');
    }

    if (!channelIds || channelIds.length === 0) {
      throw new Error('No channels selected');
    }

    const botToken = getSlackBotToken();
    const slackMessage = formatSlackMessage(summaryText, summaryType);
    const slackApiUrl = 'https://slack.com/api/chat.postMessage';

    const failures = [];

    channelIds.forEach(function(channelId) {
      try {
        const response = UrlFetchApp.fetch(slackApiUrl, {
          method: 'POST',
          contentType: 'application/json',
          headers: {
            'Authorization': `Bearer ${botToken}`,
            'Content-Type': 'application/json'
          },
          payload: JSON.stringify({
            channel: channelId,
            ...slackMessage
          }),
          muteHttpExceptions: true
        });

        const responseData = JSON.parse(response.getContentText());
        if (!responseData.ok) {
          failures.push(`${channelId}: ${responseData.error || 'unknown error'}`);
        } else {
          console.log(`✅ Sent to channel ${channelId}`);
        }
      } catch (e) {
        failures.push(`${channelId}: ${e.message}`);
      }
    });

    if (failures.length === 0) {
      const label = channelIds.length === 1 ? '1 channel' : `${channelIds.length} channels`;
      return {
        success: true,
        message: `Summary sent to ${label} successfully!`
      };
    } else if (failures.length < channelIds.length) {
      return {
        success: true,
        message: `Sent to ${channelIds.length - failures.length} of ${channelIds.length} channels. Failed: ${failures.join(', ')}`
      };
    } else {
      throw new Error(`Failed to send to all channels: ${failures.join(', ')}`);
    }

  } catch (error) {
    console.error('Error in sendToSlackWithOptions:', error);
    return {
      success: false,
      message: `Failed to send to Slack: ${error.message}`
    };
  }
}

/**
 * Test Slack integration (for debugging)
 * Call this function manually from Apps Script editor to test
 */
function testSlackIntegration() {
  console.log('=== TESTING SLACK INTEGRATION ===');
  
  // Test with a sample defects summary
  const testSummary = `DEFECTS STATUS REPORT
Generated: ${new Date().toLocaleString()}

📊 By Severity:
  • Critical: 5
  • High: 12
  • Medium: 8
  • Low: 3

📈 By Status:
  • Open: 15
  • In Progress: 10
  • Resolved: 3

Total Defects: 28`;
  
  const result = sendToSlack(testSummary, 'defects');
  
  if (result.success) {
    console.log('✅ Test successful!');
    console.log(result.message);
  } else {
    console.log('❌ Test failed!');
    console.log(result.message);
  }
  
  return result;
}
