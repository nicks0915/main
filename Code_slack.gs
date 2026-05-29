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

/**
 * Fetch the most recent Entertainment 5.0 status message from #prod-green-commerce.
 * Searches up to the last 50 messages for keywords; falls back to the most recent message.
 * Requires the bot token to have channels:history scope.
 * @returns {Object} { success, text, postedAt } or { success: false, message }
 */
function fetchExecutiveStatusFromSlack() {
  try {
    var botToken = getSlackBotToken();
    var channelId = 'C03UMGV7DDE'; // prod-green-commerce

    var response = UrlFetchApp.fetch(
      'https://slack.com/api/conversations.history?channel=' + channelId + '&limit=50',
      {
        method: 'GET',
        headers: { 'Authorization': 'Bearer ' + botToken },
        muteHttpExceptions: true
      }
    );

    var data = JSON.parse(response.getContentText());

    if (!data.ok) {
      var hint = data.error === 'missing_scope'
        ? ' — add channels:history scope to your Slack app at api.slack.com/apps'
        : '';
      return { success: false, message: 'Slack API error: ' + data.error + hint };
    }

    var messages = data.messages || [];
    var keywords = ['entertainment 5.0', 'optik tv on green', 'entertainment on green', 'optik tv'];

    // Find most recent human message matching any keyword
    for (var i = 0; i < messages.length; i++) {
      var msg = messages[i];
      if (msg.type !== 'message' || msg.subtype) continue;
      var lower = (msg.text || '').toLowerCase();
      if (keywords.some(function(k) { return lower.indexOf(k) !== -1; })) {
        return {
          success: true,
          text: formatSlackText(msg.text),
          postedAt: new Date(parseFloat(msg.ts) * 1000).toLocaleString('en-US', {
            year: 'numeric', month: 'short', day: 'numeric',
            hour: '2-digit', minute: '2-digit'
          })
        };
      }
    }

    // No keyword match — return most recent human message
    for (var j = 0; j < messages.length; j++) {
      var m = messages[j];
      if (m.type === 'message' && !m.subtype && m.text) {
        return {
          success: true,
          text: formatSlackText(m.text),
          postedAt: new Date(parseFloat(m.ts) * 1000).toLocaleString('en-US', {
            year: 'numeric', month: 'short', day: 'numeric',
            hour: '2-digit', minute: '2-digit'
          }),
          note: 'No exact Entertainment 5.0 keyword match — showing most recent message'
        };
      }
    }

    return { success: false, message: 'No messages found in #prod-green-commerce' };

  } catch (error) {
    return { success: false, message: 'Error fetching from Slack: ' + error.message };
  }
}

/**
 * Convert raw Slack API message text to clean readable text.
 * - Decodes HTML entities (&amp; &lt; &gt;)
 * - Converts <mailto:email|Name> and <url|Label> links to just the label
 * - Removes bare URL links and user/channel mentions
 * - Maps common Slack :emoji_name: codes to Unicode emoji characters
 * - Collapses extra whitespace left by removed elements
 * Bold markers (*text*) are left in place — the client converts them to HTML <strong>.
 */
function formatSlackText(text) {
  if (!text) return '';

  // 1. Decode HTML entities Slack encodes in API responses
  text = text.replace(/&amp;/g, '&')
             .replace(/&lt;/g, '<')
             .replace(/&gt;/g, '>');

  // 2. Email/URL links with display label: <mailto:...|Name> or <https://...|Label> → Label
  text = text.replace(/<(?:mailto|https?|http):[^|>]+\|([^>]+)>/g, '$1');

  // 3. Bare URL links: <https://...> → remove
  text = text.replace(/<https?:[^>]+>/g, '');

  // 4. User mentions: <@UXXXXXXX> → remove
  text = text.replace(/<@[A-Z0-9]+>/g, '');

  // 5. Channel links: <#CXXXXXXX|channel-name> → #channel-name
  text = text.replace(/<#[A-Z0-9]+\|([^>]+)>/g, '#$1');

  // 6. Map Slack emoji codes to Unicode
  var EMOJI = {
    'hourglass_flowing_sand': '⏳', 'hourglass': '⌛',
    'dart': '🎯', 'rocket': '🚀',
    'large_yellow_circle': '{CIRCLE_YELLOW}', 'large_green_circle': '{CIRCLE_GREEN}',
    'large_orange_circle': '{CIRCLE_ORANGE}', 'large_blue_circle': '{CIRCLE_BLUE}',
    'red_circle': '{CIRCLE_RED}', 'large_red_circle': '{CIRCLE_RED}',
    'white_check_mark': '✅', 'heavy_check_mark': '✔️', 'check': '✔️',
    'ballot_box_with_check': '☑️',
    'x': '❌', 'warning': '⚠️', 'fire': '🔥',
    'chart_with_upwards_trend': '📈', 'chart_with_downwards_trend': '📉',
    'bar_chart': '📊', 'clipboard': '📋', 'memo': '📝', 'pushpin': '📌',
    'calendar': '📅', 'checkered_flag': '🏁', 'rotating_light': '🚨',
    'eyes': '👀', 'thumbsup': '👍', 'thumbsdown': '👎',
    'raised_hands': '🙌', 'tada': '🎉', 'trophy': '🏆',
    'star': '⭐', 'bell': '🔔', 'bulb': '💡',
    'construction': '🚧', 'hammer': '🔨', 'wrench': '🔧', 'mag': '🔍',
    'exclamation': '❗', 'question': '❓', 'information_source': 'ℹ️',
    'arrow_right': '➡️', 'fast_forward': '⏩', 'stopwatch': '⏱️',
    'zap': '⚡', 'no_entry': '⛔', 'no_entry_sign': '🚫',
    'heavy_minus_sign': '➖', 'heavy_plus_sign': '➕', 'ok': '🆗',
    'link': '🔗', 'lock': '🔒', 'unlock': '🔓', 'key': '🔑'
  };

  text = text.replace(/:([a-z0-9_+-]+):/g, function(match, name) {
    return Object.prototype.hasOwnProperty.call(EMOJI, name) ? EMOJI[name] : '';
  });

  // 7. Collapse multiple spaces/tabs left by removals (preserve newlines)
  text = text.replace(/[ \t]{2,}/g, ' ');

  return text.trim();
}
