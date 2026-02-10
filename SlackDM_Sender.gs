/**
 * Slack Direct Message Sender
 * 
 * This script reads recipient names and messages from a Google Sheet
 * and sends Slack direct messages to each recipient.
 * 
 * SUPPORTS BOTH BOT AND USER TOKENS:
 * - Bot Token: Messages appear FROM THE BOT (one token for all users)
 * - User Token: Messages appear FROM YOUR PERSONAL SLACK ACCOUNT
 * 
 * SETUP INSTRUCTIONS:
 * 1. Open your Google Sheet
 * 2. Go to Extensions > Apps Script
 * 3. Paste this code
 * 4. Replace 'YOUR_SLACK_BOT_TOKEN' below with your token:
 *    - Bot token (starts with xoxb-) OR
 *    - User token (starts with xoxp-)
 * 5. Save the script
 * 6. Refresh your Google Sheet - you'll see a new "Slack Tools" menu
 * 7. Set TOKEN TYPE in Row 4, Column A: "Bot" or "User"
 * 
 * SHEET FORMAT (auto-created):
 * Row 1: "TAB NAME:" label
 * Row 2: Tab name input field
 * Row 3: "TOKEN TYPE:" label
 * Row 4: Token type input ("Bot" or "User")
 * Row 5: Headers (Recipient Name, Message, Status, Response)
 * Row 6+: Data rows
 * 
 * TO GET YOUR TOKEN:
 * Bot Token:
 * 1. Go to https://api.slack.com/apps
 * 2. Select your app (e.g., "BulkDM")
 * 3. Go to "OAuth & Permissions"
 * 4. Copy the "Bot User OAuth Token" (starts with xoxb-)
 * 
 * User Token:
 * 1. Use browser extension (e.g., "Slack User Token Fetcher") OR
 * 2. OAuth flow if app supports user scopes OR
 * 3. Legacy tokens (may be deprecated)
 * 4. Required scopes: chat:write, im:write, users:read, im:history
 * 5. Copy the token (starts with xoxp-)
 * Note: User tokens are harder to obtain - see SETUP_GUIDE.md for details
 * 
 * Required scopes for both: chat:write, im:write, users:read, im:history
 */

// ============================================
// CONFIGURATION - UPDATE THIS WITH YOUR TOKEN
// ============================================
// Replace with your Bot token (xoxb-...) OR User token (xoxp-...)
// Set TOKEN TYPE in Row 4, Column A of your sheet to match
const SLACK_BOT_TOKEN = 'YOUR_SLACK_BOT_TOKEN'; // Can be Bot or User token

// Optional: Customize these if your sheet has different column positions
const COLUMN_RECIPIENT = 1; // Column A
const COLUMN_MESSAGE = 2;   // Column B
const COLUMN_STATUS = 3;    // Column C
const COLUMN_RESPONSE = 4;  // Column D - Recipient responses

// Configuration row positions
const ROW_TAB_NAME_LABEL = 1;      // Row 1: "TAB NAME:"
const ROW_TAB_NAME_VALUE = 2;      // Row 2: [Input field for tab name]
const ROW_TOKEN_TYPE_LABEL = 3;     // Row 3: "TOKEN TYPE:"
const ROW_TOKEN_TYPE_VALUE = 4;    // Row 4: [Bot or User]
const ROW_HEADERS = 5;              // Row 5: Headers start here
const ROW_DATA_START = 6;           // Row 6: Data starts here

// ============================================
// MAIN FUNCTIONS
// ============================================

/**
 * Creates a custom menu in Google Sheets and sets up headers
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Slack Tools')
    .addItem('Send Messages', 'sendAllMessages')
    .addItem('Read Responses', 'readAllResponses')
    .addItem('Test Connection', 'testSlackConnection')
    .addSeparator()
    .addItem('Setup Headers', 'setupHeaders')
    .addToUi();
  
  // Automatically set up headers if sheet is empty or missing headers
  setupHeaders();
}

/**
 * Sets up configuration section and column headers in the sheet
 */
function setupHeaders() {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Set up TAB NAME configuration section
  const tabNameLabel = sheet.getRange(ROW_TAB_NAME_LABEL, COLUMN_RECIPIENT).getValue();
  if (!tabNameLabel || tabNameLabel.toString().trim() !== 'TAB NAME:') {
    sheet.getRange(ROW_TAB_NAME_LABEL, COLUMN_RECIPIENT).setValue('TAB NAME:');
    sheet.getRange(ROW_TAB_NAME_LABEL, COLUMN_RECIPIENT).setFontWeight('bold');
    
    // Set default tab name to current sheet if empty
    const currentTabName = sheet.getRange(ROW_TAB_NAME_VALUE, COLUMN_RECIPIENT).getValue();
    if (!currentTabName || currentTabName.toString().trim() === '') {
      sheet.getRange(ROW_TAB_NAME_VALUE, COLUMN_RECIPIENT).setValue(sheet.getName());
    }
    
    // Format the input cell
    const inputCell = sheet.getRange(ROW_TAB_NAME_VALUE, COLUMN_RECIPIENT);
    inputCell.setBackground('#fff9c4'); // Light yellow background
    inputCell.setNote('Enter the name of the tab where your data lives. This tab will be used for all operations.');
  }
  
  // Set up TOKEN TYPE configuration section
  const tokenTypeLabel = sheet.getRange(ROW_TOKEN_TYPE_LABEL, COLUMN_RECIPIENT).getValue();
  if (!tokenTypeLabel || tokenTypeLabel.toString().trim() !== 'TOKEN TYPE:') {
    sheet.getRange(ROW_TOKEN_TYPE_LABEL, COLUMN_RECIPIENT).setValue('TOKEN TYPE:');
    sheet.getRange(ROW_TOKEN_TYPE_LABEL, COLUMN_RECIPIENT).setFontWeight('bold');
    
    // Set default to Bot if empty
    const currentTokenType = sheet.getRange(ROW_TOKEN_TYPE_VALUE, COLUMN_RECIPIENT).getValue();
    if (!currentTokenType || currentTokenType.toString().trim() === '') {
      sheet.getRange(ROW_TOKEN_TYPE_VALUE, COLUMN_RECIPIENT).setValue('Bot');
    }
    
    // Format the input cell with data validation
    const tokenTypeCell = sheet.getRange(ROW_TOKEN_TYPE_VALUE, COLUMN_RECIPIENT);
    tokenTypeCell.setBackground('#fff9c4'); // Light yellow background
    tokenTypeCell.setNote('Enter "Bot" to send from bot account, or "User" to send from your personal Slack account. See setup guide for User token instructions.');
    
    // Create data validation for Bot/User
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Bot', 'User'], true)
      .setAllowInvalid(false)
      .build();
    tokenTypeCell.setDataValidation(rule);
  }
  
  // Check if headers already exist
  const headerRow = sheet.getRange(ROW_HEADERS, 1, 1, 4).getValues()[0];
  const hasHeaders = headerRow[0] && headerRow[0].toString().trim().length > 0;
  
  if (!hasHeaders) {
    // Set up headers
    sheet.getRange(ROW_HEADERS, COLUMN_RECIPIENT).setValue('Recipient Name');
    sheet.getRange(ROW_HEADERS, COLUMN_MESSAGE).setValue('Message');
    sheet.getRange(ROW_HEADERS, COLUMN_STATUS).setValue('Status');
    sheet.getRange(ROW_HEADERS, COLUMN_RESPONSE).setValue('Response');
    
    // Format header row (bold, background color)
    const headerRange = sheet.getRange(ROW_HEADERS, 1, 1, 4);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    
    // Set column widths for better readability
    sheet.setColumnWidth(COLUMN_RECIPIENT, 200); // Recipient Name
    sheet.setColumnWidth(COLUMN_MESSAGE, 400);   // Message
    sheet.setColumnWidth(COLUMN_STATUS, 150);    // Status
    sheet.setColumnWidth(COLUMN_RESPONSE, 500);  // Response
    
    // Freeze header row
    sheet.setFrozenRows(ROW_HEADERS);
  }
}

/**
 * Gets the target sheet based on TAB NAME configuration
 * Returns the sheet object or null if not found
 */
function getTargetSheet() {
  const currentSheet = SpreadsheetApp.getActiveSheet();
  const tabName = currentSheet.getRange(ROW_TAB_NAME_VALUE, COLUMN_RECIPIENT).getValue();
  
  if (!tabName || tabName.toString().trim() === '') {
    SpreadsheetApp.getUi().alert(
      'Configuration Required',
      `Please enter a TAB NAME in row ${ROW_TAB_NAME_VALUE}, column A.\n\nThis should be the name of the tab where your data lives.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return null;
  }
  
  const tabNameStr = tabName.toString().trim();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const targetSheet = spreadsheet.getSheetByName(tabNameStr);
    if (!targetSheet) {
      SpreadsheetApp.getUi().alert(
        'Tab Not Found',
        `The tab "${tabNameStr}" does not exist.\n\nPlease check the TAB NAME in row ${ROW_TAB_NAME_VALUE}, column A and make sure the tab exists in this spreadsheet.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return null;
    }
    return targetSheet;
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Error',
      `Error accessing tab "${tabNameStr}": ${error.toString()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return null;
  }
}

/**
 * Gets the token type (Bot or User) from configuration
 * Returns 'Bot' or 'User'
 */
function getTokenType() {
  const currentSheet = SpreadsheetApp.getActiveSheet();
  const tokenType = currentSheet.getRange(ROW_TOKEN_TYPE_VALUE, COLUMN_RECIPIENT).getValue();
  
  if (!tokenType) {
    return 'Bot'; // Default to Bot
  }
  
  const tokenTypeStr = tokenType.toString().trim();
  if (tokenTypeStr.toLowerCase() === 'user') {
    return 'User';
  }
  return 'Bot'; // Default to Bot
}

/**
 * Gets the appropriate token based on token type configuration
 * Returns the token string or null if not set
 */
function getSlackToken() {
  const tokenType = getTokenType();
  
  if (tokenType === 'User') {
    // For User tokens, check if there's a USER_TOKEN constant or use the same token
    // Note: User tokens start with xoxp-, Bot tokens start with xoxb-
    // For now, we'll use the same SLACK_BOT_TOKEN variable but expect it to be a User token
    // In a real implementation, you might want separate variables
    if (SLACK_BOT_TOKEN === 'YOUR_SLACK_BOT_TOKEN') {
      return null;
    }
    return SLACK_BOT_TOKEN;
  } else {
    // Bot token
    if (SLACK_BOT_TOKEN === 'YOUR_SLACK_BOT_TOKEN') {
      return null;
    }
    return SLACK_BOT_TOKEN;
  }
}

/**
 * Main function to send all messages from the sheet
 */
function sendAllMessages() {
  const ui = SpreadsheetApp.getUi();
  
  // Get the target sheet based on TAB NAME configuration
  const sheet = getTargetSheet();
  if (!sheet) {
    return; // Error already shown in getTargetSheet()
  }
  
  // Check if token is set
  if (SLACK_BOT_TOKEN === 'YOUR_SLACK_BOT_TOKEN') {
    ui.alert('Setup Required', 
      'Please update SLACK_BOT_TOKEN in the script with your Bot token.\n\n' +
      'Go to Extensions > Apps Script and replace YOUR_SLACK_BOT_TOKEN', 
      ui.ButtonSet.OK);
    return;
  }
  
  // Get data from sheet (starting from ROW_DATA_START)
  const lastRow = sheet.getLastRow();
  if (lastRow < ROW_DATA_START) {
    ui.alert('No Data', `Please add recipient names and messages starting from row ${ROW_DATA_START}.`, ui.ButtonSet.OK);
    return;
  }
  
  const dataRange = sheet.getRange(ROW_DATA_START, 1, lastRow - ROW_HEADERS, 4);
  const values = dataRange.getValues();
  
  if (values.length === 0) {
    ui.alert('No Data', `Please add recipient names and messages starting from row ${ROW_DATA_START}.`, ui.ButtonSet.OK);
    return;
  }
  
  // Confirm before sending
  const tokenType = getTokenType();
  const senderName = tokenType === 'User' ? 'YOUR SLACK ACCOUNT' : 'THE BOT';
  const response = ui.alert(
    'Send Messages?',
    `This will send ${values.length} message(s) FROM ${senderName}.\n\nToken Type: ${tokenType}\n\nContinue?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  // Get all Slack users once (for efficiency)
  const slackUsers = getAllSlackUsers();
  if (!slackUsers) {
    ui.alert('Error', `Failed to connect to Slack. Please check your ${tokenType} token.`, ui.ButtonSet.OK);
    return;
  }
  
  // Process each row
  let successCount = 0;
  let errorCount = 0;
  
  for (let i = 0; i < values.length; i++) {
    const row = ROW_DATA_START + i; // Sheet row number (1-indexed)
    const recipient = values[i][COLUMN_RECIPIENT - 1];
    const message = values[i][COLUMN_MESSAGE - 1];
    
    // Skip empty rows
    if (!recipient || !message) {
      sheet.getRange(row, COLUMN_STATUS).setValue('⏭️ Skipped (empty)');
      continue;
    }
    
    // Update status to processing
    sheet.getRange(row, COLUMN_STATUS).setValue('⏳ Sending...');
    SpreadsheetApp.flush(); // Force update
    
    // Find user
    const userId = findSlackUser(recipient, slackUsers);
    
    if (!userId) {
      sheet.getRange(row, COLUMN_STATUS).setValue('❌ User not found');
      errorCount++;
      continue;
    }
    
    // Send message
    const result = sendSlackDM(userId, message);
    
    if (result.success) {
      sheet.getRange(row, COLUMN_STATUS).setValue('✅ Sent');
      successCount++;
    } else {
      sheet.getRange(row, COLUMN_STATUS).setValue(`❌ Error: ${result.error}`);
      errorCount++;
    }
    
    // Rate limiting: wait 1.1 seconds between messages (Slack allows ~1/sec)
    if (i < values.length - 1) {
      Utilities.sleep(1100);
    }
  }
  
  // Show completion message
  ui.alert(
    'Complete!',
    `Messages sent:\n✅ Success: ${successCount}\n❌ Errors: ${errorCount}`,
    ui.ButtonSet.OK
  );
}

/**
 * Read responses from all recipients
 * Fetches all DM conversations and records recipient responses in the sheet
 */
function readAllResponses() {
  const ui = SpreadsheetApp.getUi();
  
  // Get the target sheet based on TAB NAME configuration
  const sheet = getTargetSheet();
  if (!sheet) {
    return; // Error already shown in getTargetSheet()
  }
  
  // Check if token is set
  const token = getSlackToken();
  if (!token) {
    const tokenType = getTokenType();
    ui.alert('Setup Required', 
      `Please update SLACK_BOT_TOKEN in the script with your ${tokenType} token.\n\n` +
      'Go to Extensions > Apps Script and replace YOUR_SLACK_BOT_TOKEN', 
      ui.ButtonSet.OK);
    return;
  }
  
  // Get user/bot info to identify sender ID
  const tokenType = getTokenType();
  const senderInfo = getBotInfo();
  if (!senderInfo) {
    ui.alert('Error', `Failed to get ${tokenType === 'User' ? 'user' : 'bot'} information. Please check your token.`, ui.ButtonSet.OK);
    return;
  }
  const senderUserId = senderInfo.user_id;
  
  // Get all Slack users for matching
  const slackUsers = getAllSlackUsers();
  if (!slackUsers) {
    const tokenType = getTokenType();
    ui.alert('Error', `Failed to connect to Slack. Please check your ${tokenType} token.`, ui.ButtonSet.OK);
    return;
  }
  
  // Get data from sheet (starting from ROW_DATA_START)
  const lastRow = sheet.getLastRow();
  if (lastRow < ROW_DATA_START) {
    ui.alert('No Data', `Please add recipient names starting from row ${ROW_DATA_START}.`, ui.ButtonSet.OK);
    return;
  }
  
  const dataRange = sheet.getRange(ROW_DATA_START, 1, lastRow - ROW_HEADERS, 4);
  const values = dataRange.getValues();
  
  if (values.length === 0) {
    ui.alert('No Data', `Please add recipient names starting from row ${ROW_DATA_START}.`, ui.ButtonSet.OK);
    return;
  }
  
  // Confirm before reading
  const response = ui.alert(
    'Read Responses?',
    `This will check for responses from ${values.length} recipient(s). Continue?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  // Get all DM conversations
  const dmConversations = getAllDMConversations();
  if (!dmConversations || dmConversations.length === 0) {
    ui.alert('No Conversations', 'No DM conversations found.', ui.ButtonSet.OK);
    return;
  }
  
  // Create a map of user ID to user info for quick lookup
  const userMap = {};
  for (const user of slackUsers) {
    userMap[user.id] = user;
  }
  
  // Create a map of recipient name/email to row number
  const recipientMap = {};
  for (let i = 0; i < values.length; i++) {
    const recipient = values[i][COLUMN_RECIPIENT - 1];
    if (recipient) {
      recipientMap[recipient.toString().trim().toLowerCase()] = ROW_DATA_START + i; // row number
    }
  }
  
  let foundCount = 0;
  let responseCount = 0;
  
  // Process each DM conversation
  for (const conversation of dmConversations) {
    // Get the other user in the DM (not the sender)
    const otherUserId = conversation.user;
    if (!otherUserId || otherUserId === senderUserId) {
      continue;
    }
    
    const otherUser = userMap[otherUserId];
    if (!otherUser) {
      continue;
    }
    
    // Find matching row in sheet by user
    let matchedRow = null;
    const userEmail = otherUser.profile && otherUser.profile.email ? otherUser.profile.email.toLowerCase() : '';
    const userName = otherUser.name ? otherUser.name.toLowerCase() : '';
    const userDisplayName = otherUser.profile && otherUser.profile.display_name ? 
      otherUser.profile.display_name.toLowerCase() : '';
    const userRealName = otherUser.profile && otherUser.profile.real_name ? 
      otherUser.profile.real_name.toLowerCase() : '';
    
    // Try to match by various identifiers
    for (const [recipientKey, rowNum] of Object.entries(recipientMap)) {
      const recipientLower = recipientKey;
      if (recipientLower === userEmail ||
          recipientLower === userName ||
          recipientLower === userDisplayName ||
          recipientLower === userRealName ||
          (userRealName && (userRealName.includes(recipientLower) || recipientLower.includes(userRealName)))) {
        matchedRow = rowNum;
        break;
      }
    }
    
    if (!matchedRow) {
      continue; // No match found in sheet
    }
    
    foundCount++;
    
    // Update status to processing
    sheet.getRange(matchedRow, COLUMN_RESPONSE).setValue('⏳ Reading...');
    SpreadsheetApp.flush();
    
    // Get conversation history
    const messages = getConversationHistory(conversation.id);
    if (!messages) {
      sheet.getRange(matchedRow, COLUMN_RESPONSE).setValue('❌ Error reading');
      continue;
    }
    
    // Filter to only messages from the recipient (not from bot)
    const recipientMessages = messages.filter(msg => 
      msg.user === otherUserId && msg.text && !msg.subtype
    );
    
    if (recipientMessages.length === 0) {
      sheet.getRange(matchedRow, COLUMN_RESPONSE).setValue('No response yet');
    } else {
      // Combine all messages into one response (newest first, then reverse)
      recipientMessages.sort((a, b) => parseFloat(b.ts) - parseFloat(a.ts)); // Sort by timestamp
      recipientMessages.reverse(); // Oldest first
      const allResponses = recipientMessages.map(msg => msg.text).join('\n\n---\n\n');
      
      sheet.getRange(matchedRow, COLUMN_RESPONSE).setValue(allResponses);
      responseCount += recipientMessages.length;
    }
    
    // Rate limiting
    Utilities.sleep(200); // Small delay between conversations
  }
  
  // Show completion message
  ui.alert(
    'Complete!',
    `Responses read:\n✅ Conversations checked: ${foundCount}\n💬 Total messages found: ${responseCount}`,
    ui.ButtonSet.OK
  );
}

/**
 * Test Slack connection
 */
function testSlackConnection() {
  const ui = SpreadsheetApp.getUi();
  
  if (SLACK_BOT_TOKEN === 'YOUR_SLACK_BOT_TOKEN') {
    ui.alert('Setup Required', 'Please set your Bot token first.', ui.ButtonSet.OK);
    return;
  }
  
  ui.alert('Testing...', 'Checking Slack connection...', ui.ButtonSet.OK);
  
  const users = getAllSlackUsers();
  if (users) {
    ui.alert(
      '✅ Connection Successful!',
      `Connected to Slack!\n\nFound ${users.length} users in your workspace.\n\n` +
      `Note: Messages will be sent FROM THE BOT, not from individual users.`,
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      '❌ Connection Failed',
      'Could not connect to Slack. Please check:\n\n' +
      '1. Your Bot token is correct (should start with xoxb-)\n' +
      '2. Bot has required scopes (chat:write, im:write, users:read)\n' +
      '3. Bot is installed in your workspace\n' +
      '4. Token is not expired',
      ui.ButtonSet.OK
    );
  }
}

// ============================================
// SLACK API FUNCTIONS
// ============================================

/**
 * Get all users from Slack workspace
 */
function getAllSlackUsers() {
  try {
    const token = getSlackToken();
    if (!token) {
      return null;
    }
    
    const url = 'https://slack.com/api/users.list';
    const options = {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (data.ok) {
      return data.members.filter(user => !user.deleted && !user.is_bot);
    } else {
      Logger.log('Slack API Error: ' + data.error);
      return null;
    }
  } catch (error) {
    Logger.log('Error fetching users: ' + error);
    return null;
  }
}

/**
 * Find Slack user ID by name or email
 * Tries multiple matching strategies for best results
 */
function findSlackUser(recipient, slackUsers) {
  if (!recipient || !slackUsers) return null;
  
  const searchTerm = recipient.toString().trim().toLowerCase();
  
  for (const user of slackUsers) {
    // Match by email (most reliable)
    if (user.profile && user.profile.email) {
      if (user.profile.email.toLowerCase() === searchTerm) {
        return user.id;
      }
    }
    
    // Match by display name
    if (user.profile && user.profile.display_name) {
      if (user.profile.display_name.toLowerCase() === searchTerm) {
        return user.id;
      }
    }
    
    // Match by real name
    if (user.profile && user.profile.real_name) {
      if (user.profile.real_name.toLowerCase() === searchTerm) {
        return user.id;
      }
    }
    
    // Match by username (without @)
    if (user.name) {
      if (user.name.toLowerCase() === searchTerm.replace('@', '')) {
        return user.id;
      }
    }
    
    // Partial match on real name (e.g., "John" matches "John Smith")
    if (user.profile && user.profile.real_name) {
      const realName = user.profile.real_name.toLowerCase();
      if (realName.includes(searchTerm) || searchTerm.includes(realName)) {
        return user.id;
      }
    }
  }
  
  return null;
}

/**
 * Get bot/user information to identify sender user ID
 */
function getBotInfo() {
  try {
    const token = getSlackToken();
    if (!token) {
      return null;
    }
    
    const url = 'https://slack.com/api/auth.test';
    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (data.ok) {
      return {
        user_id: data.user_id,
        bot_id: data.bot_id
      };
    } else {
      Logger.log('Auth test error: ' + data.error);
      return null;
    }
  } catch (error) {
    Logger.log('Error getting sender info: ' + error);
    return null;
  }
}

/**
 * Get all DM conversations for the bot/user
 */
function getAllDMConversations() {
  try {
    const token = getSlackToken();
    if (!token) {
      return null;
    }
    
    const url = 'https://slack.com/api/conversations.list';
    const options = {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    };
    
    // Add query parameters
    const params = {
      types: 'im',
      exclude_archived: true,
      limit: 1000
    };
    
    const queryString = Object.keys(params).map(key => 
      encodeURIComponent(key) + '=' + encodeURIComponent(params[key])
    ).join('&');
    
    const response = UrlFetchApp.fetch(url + '?' + queryString, options);
    const data = JSON.parse(response.getContentText());
    
    if (data.ok) {
      return data.channels || [];
    } else {
      Logger.log('Conversations list error: ' + data.error);
      return null;
    }
  } catch (error) {
    Logger.log('Error fetching conversations: ' + error);
    return null;
  }
}

/**
 * Get conversation history (messages) for a channel
 */
function getConversationHistory(channelId) {
  try {
    const token = getSlackToken();
    if (!token) {
      return null;
    }
    
    const url = 'https://slack.com/api/conversations.history';
    const options = {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    };
    
    // Get up to 100 most recent messages
    const params = {
      channel: channelId,
      limit: 100
    };
    
    const queryString = Object.keys(params).map(key => 
      encodeURIComponent(key) + '=' + encodeURIComponent(params[key])
    ).join('&');
    
    const response = UrlFetchApp.fetch(url + '?' + queryString, options);
    const data = JSON.parse(response.getContentText());
    
    if (data.ok) {
      return data.messages || [];
    } else {
      Logger.log('Conversation history error: ' + data.error);
      return null;
    }
  } catch (error) {
    Logger.log('Error fetching conversation history: ' + error);
    return null;
  }
}

/**
 * Send a direct message to a Slack user
 */
function sendSlackDM(userId, message) {
  try {
    // First, open or get the DM channel
    const imOpenUrl = 'https://slack.com/api/conversations.open';
    const imOpenPayload = {
      users: userId
    };
    
    const token = getSlackToken();
    if (!token) {
      return {
        success: false,
        error: 'Token not configured'
      };
    }
    
    const imOpenOptions = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(imOpenPayload)
    };
    
    const imOpenResponse = UrlFetchApp.fetch(imOpenUrl, imOpenOptions);
    const imOpenData = JSON.parse(imOpenResponse.getContentText());
    
    if (!imOpenData.ok) {
      return {
        success: false,
        error: imOpenData.error || 'Failed to open DM'
      };
    }
    
    const channelId = imOpenData.channel.id;
    
    // Send the message
    const token = getSlackToken();
    if (!token) {
      return {
        success: false,
        error: 'Token not configured'
      };
    }
    
    const tokenType = getTokenType();
    const chatUrl = 'https://slack.com/api/chat.postMessage';
    const chatPayload = {
      channel: channelId,
      text: message
    };
    
    // as_user: true only works with User tokens (deprecated but may still work)
    // For User tokens, messages automatically appear from the user
    // For Bot tokens, messages appear from the bot
    if (tokenType === 'User') {
      chatPayload.as_user = true;
    }
    
    const chatOptions = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(chatPayload)
    };
    
    const chatResponse = UrlFetchApp.fetch(chatUrl, chatOptions);
    const chatData = JSON.parse(chatResponse.getContentText());
    
    if (chatData.ok) {
      return { success: true };
    } else {
      return {
        success: false,
        error: chatData.error || 'Failed to send message'
      };
    }
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}
