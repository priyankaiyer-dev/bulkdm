/**
 * BULKDM – Send Slack DMs from a Google Sheet and read replies
 *
 * ========== HOW TO USE (most users – you have a copy of the sheet) ==========
 *
 * 1. Open your copy of the BulkDM Google Sheet.
 *    When the sheet loads, it automatically creates the layout: headers, columns, and configuration rows.
 *    You don’t need to type anything in the config area unless you use multiple tabs (see step 2).
 *
 * 2. Add your data (the only required input):
 *    From row 6 down, add one row per recipient.
 *    • Column A (Recipient Email): who to message – email is best but name also works. Name will take about ~1-2 minutes the first time you run the script.
 *    • Column B (Message): the message to send.
 *    Columns C–F (Status, Response, Slack ID, Last Sent Timestamp) are filled in automatically; leave them blank.
 *
 * 3. Connect to Slack:
 *    Click Slack Tools → Connect to Slack (or Get User Token if you want to send as yourself).
 *    Follow the link, approve the app, then paste the token if prompted (or it may be saved automatically).
 *    Token type (Bot vs User) is set automatically when you connect – you don’t choose it manually.
 *
 * 4. Send messages:
 *    Click Slack Tools → Send Messages. Confirm when asked. The tool sends each row’s message and updates Status and Slack ID.
 *
 * 5. Read replies:
 *    Click Slack Tools → Read Responses. Replies appear in the Response column (only messages the recipient sent after your BulkDM message).
 *
 * 6. Other:
 *    Slack Tools → Test Connection: check your Slack connection. Slack Tools → Setup Headers: restore the layout if needed.
 *
 * ========== IF YOU ARE SETTING UP THE SHEET FROM SCRATCH OR OPENING APPS SCRIPT ==========
 *
 * • Open the sheet → Extensions → Apps Script. Paste this entire file (replace any existing code). Save.
 * • In the script, find the CONFIG section (near the top). Set SLACK_BOT_TOKEN to your Slack Bot User OAuth Token (xoxb-...) if you use a Bot; for User tokens you can leave it and use Connect to Slack from the sheet instead.
 * • Slack app scopes (api.slack.com → your app → OAuth & Permissions):
 *   Bot: chat:write, im:write, users:read, im:read, im:history. User: chat:write, im:write, users:read, users:read.email, im:read, im:history. Then use Connect to Slack in the sheet for User.
 * • Refresh the sheet. You should see Slack Tools. If not, in Apps Script choose "onOpen" and click Run once.

// ============================================
// CONFIG
// ============================================

// DO NOT CHANGE:
const SLACK_BOT_TOKEN = 'YOUR_SLACK_BOT_TOKEN';
const SLACK_CLIENT_ID = '3895842157.10360289984709';
const SLACK_CLIENT_SECRET = '19ea914a1237c9b07f677b53efe83045';
const SLACK_REDIRECT_URI = 'https://script.google.com/a/macros/samsara.com/s/AKfycbzWeDVnX7JPvOOto8bdkDILeOgs1TmJAH01iXrfo8JHewLeWXw2HfQq53_b7nRPKuv6/exec';

// DO NOT CHANGE: User scopes requested when someone clicks "Add to Slack" (for sending DMs as themselves)
// users:read.email enables lookup-by-email so we only look up recipients instead of pulling the full list
const SLACK_USER_SCOPES = 'chat:write,im:write,users:read,users:read.email,im:read,im:history';

// DO NOT CHANGE: Key for storing token in Document Properties (per-sheet); no copy-paste needed after Connect to Slack
const SLACK_TOKEN_PROPERTY = 'SLACK_TOKEN';

// OPTIONAL: Set to 200 or 400 to only fetch that many users (faster first run). Set to 0 to fetch all (slower).
const MAX_USERS_TO_FETCH = 0;

// OPTIONAL: Customize these if your sheet has different column positions
const COLUMN_RECIPIENT = 1; // Column A
const COLUMN_MESSAGE = 2;   // Column B
const COLUMN_STATUS = 3;   // Column C
const COLUMN_RESPONSE = 4;  // Column D - Recipient responses
const COLUMN_SLACK_ID = 5;  // Column E - Slack user ID (filled on first successful send)
const COLUMN_LAST_SENT_TS = 6; // Column F - Timestamp of our last sent message (Read Responses only shows recipient messages after this, excluding the BulkDM message itself)

// OPTIONAL: Configuration row positions
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
    .addItem('Connect to Slack', 'showConnectToSlackSidebar')
    .addItem('Send Messages', 'sendAllMessages')
    .addItem('Read Responses', 'readAllResponses')
    .addItem('Test Connection', 'testSlackConnection')
    .addSeparator()
    .addItem('Get User Token (OAuth)', 'showGetUserTokenDialog')
    .addSeparator()
    .addItem('Setup Headers', 'setupHeaders')
    .addToUi();
  
  // Automatically set up headers if sheet is empty or missing headers
  setupHeaders();
}

/**
 * Opens a dialog with a link to the "Get User Token" OAuth page.
 * Users click the link to authorize with Slack and copy their User Token into the sheet.
 */
function showGetUserTokenDialog() {
  const ui = SpreadsheetApp.getUi();
  const redirectUri = (typeof SLACK_REDIRECT_URI !== 'undefined' && SLACK_REDIRECT_URI) ? SLACK_REDIRECT_URI.trim() : '';
  if (!redirectUri) {
    ui.alert(
      'Get User Token',
      'OAuth is not configured yet.\n\n' +
      'The person who set up the BulkDM app needs to:\n' +
      '1. Deploy this script as a Web app (Deploy > New deployment > Web app)\n' +
      '2. Set SLACK_CLIENT_ID, SLACK_CLIENT_SECRET, and SLACK_REDIRECT_URI in the script\n' +
      '3. Add the same SLACK_REDIRECT_URI to the Slack app under OAuth & Permissions → Redirect URLs\n\n' +
      'Then use Slack Tools → Get User Token (OAuth) again.',
      ui.ButtonSet.OK
    );
    return;
  }
  const html = HtmlService.createHtmlOutput(
    '<p style="font-family:sans-serif;">To get your <strong>User Token</strong> so messages are sent from your Slack account:</p>' +
    '<ol style="font-family:sans-serif; text-align:left;">' +
    '<li>Click the link below (opens in your browser)</li>' +
    '<li>Click <strong>Allow</strong> to install BulkDM</li>' +
    '<li>Copy the <strong>User Token</strong> shown on the next page</li>' +
    '<li>Open <strong>Slack Tools → Connect to Slack</strong>, paste into the box, and click <strong>Save token</strong></li>' +
    '<li>Token type is set automatically when you save.</li>' +
    '</ol>' +
    '<p><a href="' + redirectUri + '" target="_blank" style="font-size:14px;">Open Get User Token page</a></p>' +
    '<p style="font-size:11px; color:#666;">If the link does not open, copy this URL: ' + redirectUri + '</p>'
  )
    .setWidth(420)
    .setHeight(260);
  ui.showModalDialog(html, 'Get User Token');
}

function getSlackOAuthUrl() {
  var redirectUri = (typeof SLACK_REDIRECT_URI !== 'undefined' && SLACK_REDIRECT_URI) ? String(SLACK_REDIRECT_URI).trim() : '';
  var clientId = (typeof SLACK_CLIENT_ID !== 'undefined' && SLACK_CLIENT_ID) ? String(SLACK_CLIENT_ID).trim() : '';
  if (!redirectUri || !clientId) return null;
  var scopes = (typeof SLACK_USER_SCOPES !== 'undefined' && SLACK_USER_SCOPES) ? SLACK_USER_SCOPES : 'chat:write,im:write,users:read,im:read,im:history';
  return 'https://slack.com/oauth/v2/authorize?client_id=' + encodeURIComponent(clientId) +
    '&user_scope=' + encodeURIComponent(scopes) + '&redirect_uri=' + encodeURIComponent(redirectUri);
}

function showConnectToSlackSidebar() {
  var html = '<div id="content" style="font-family:sans-serif; padding:12px;">Loading...</div>' +
    '<script>' +
    'function onStatus(status) {' +
    '  var el = document.getElementById("content");' +
    '  if (!el) return;' +
    '  if (status && status.connected) {' +
    '    el.innerHTML = \'<p><strong>You are connected.</strong></p><p>Messages will send from your Slack account.</p>\' +' +
    '      \'<p style="font-size:11px; margin-top:10px;">To get new permissions (e.g. email lookup): In Slack go to <strong>Settings &gt; Apps</strong>, find BulkDM, <strong>Remove</strong>. Then click Reconnect below.</p>\' +' +
    '      \'<p><button id="reconnect">Reconnect</button></p>\';' +
    '    var btn = document.getElementById("reconnect");' +
    '    if (btn) btn.onclick = startConnect;' +
    '    return;' +
    '  }' +
    '  el.innerHTML = \'<p>Connect once so BulkDM can send DMs from your Slack account.</p>\' +' +
    '    \'<p><button id="connect">Connect to Slack</button></p>\' +' +
    '    \'<p style="margin-top:12px; font-size:11px;">Or paste a token below:</p>\' +' +
    '    \'<textarea id="token" rows="3" style="width:100%; margin:4px 0;"></textarea>\' +' +
    '    \'<p><button id="save">Save token</button></p>\' +' +
    '    \'<p id="msg" style="font-size:11px;"></p>\';' +
    '  var c = document.getElementById("connect");' +
    '  var s = document.getElementById("save");' +
    '  if (c) c.onclick = startConnect;' +
    '  if (s) s.onclick = savePasted;' +
    '}' +
    'function refreshStatus() {' +
    '  google.script.run.withSuccessHandler(onStatus).getConnectionStatus();' +
    '}' +
    'function startConnect() {' +
    '  google.script.run.withSuccessHandler(function(url) {' +
    '    if (url) {' +
    '      window.open(url, "slack_oauth", "width=600,height=700");' +
    '      var msgEl = document.getElementById("msg");' +
    '      if (msgEl) msgEl.textContent = "Complete the sign-in in the popup, then this panel will update.";' +
    '      var poll = setInterval(function() {' +
    '        google.script.run.withSuccessHandler(function(st) {' +
    '          if (st && st.connected) { clearInterval(poll); onStatus(st); }' +
    '        }).getConnectionStatus();' +
    '      }, 2000);' +
    '      setTimeout(function() { clearInterval(poll); }, 60000);' +
    '    } else document.getElementById("content").innerHTML = "<p>OAuth not configured. Use Get User Token (OAuth) and paste the token here.</p>";' +
    '  }).getSlackOAuthUrl();' +
    '}' +
    'function savePasted() {' +
    '  var t = document.getElementById("token").value.trim();' +
    '  if (!t) { document.getElementById("msg").textContent = "Enter a token."; return; }' +
    '  document.getElementById("msg").textContent = "Saving...";' +
    '  google.script.run.withSuccessHandler(function(r) {' +
    '    if (r && r.ok) { refreshStatus(); document.getElementById("msg").textContent = "Saved."; }' +
    '    else document.getElementById("msg").textContent = (r && r.error) ? r.error : "Failed to save.";' +
    '  }).saveSlackToken(t);' +
    '}' +
    'window.addEventListener("message", function(e) {' +
    '  if (e.data && e.data.type === "slack_token" && e.data.token) {' +
    '    google.script.run.withSuccessHandler(function(r) { if (r && r.ok) refreshStatus(); }).saveSlackToken(e.data.token);' +
    '  }' +
    '});' +
    'document.addEventListener("visibilitychange", function() { if (document.visibilityState === "visible") refreshStatus(); });' +
    'google.script.run.withSuccessHandler(onStatus).getConnectionStatus();' +
    '</script>';
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle('Connect to Slack').setWidth(320));
}

/**
 * Web app entry point for OAuth "Get User Token" flow.
 * Deploy as Web app; set that URL as SLACK_REDIRECT_URI and in Slack app Redirect URLs.
 */
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const code = params.code;
  const redirectUri = (typeof SLACK_REDIRECT_URI !== 'undefined' && SLACK_REDIRECT_URI) ? String(SLACK_REDIRECT_URI).trim() : '';
  const clientId = (typeof SLACK_CLIENT_ID !== 'undefined' && SLACK_CLIENT_ID) ? String(SLACK_CLIENT_ID).trim() : '';
  const clientSecret = (typeof SLACK_CLIENT_SECRET !== 'undefined' && SLACK_CLIENT_SECRET) ? String(SLACK_CLIENT_SECRET).trim() : '';

  if (!redirectUri || !clientId || !clientSecret) {
    return HtmlService.createHtmlOutput(
      '<body style="font-family:sans-serif; padding:20px;"><h2>BulkDM – Get User Token</h2>' +
      '<p>OAuth is not configured. Set SLACK_CLIENT_ID, SLACK_CLIENT_SECRET, and SLACK_REDIRECT_URI in the script.</p></body>'
    ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (code) {
    return exchangeSlackOAuthCode(code, redirectUri, clientId, clientSecret);
  }

  const scopes = (typeof SLACK_USER_SCOPES !== 'undefined' && SLACK_USER_SCOPES) ? SLACK_USER_SCOPES : 'chat:write,im:write,users:read,im:read,im:history';
  const addUrl = 'https://slack.com/oauth/v2/authorize?client_id=' + encodeURIComponent(clientId) +
    '&user_scope=' + encodeURIComponent(scopes) + '&redirect_uri=' + encodeURIComponent(redirectUri);
  return HtmlService.createHtmlOutput(
    '<body style="font-family:sans-serif; padding:24px; max-width:480px;">' +
    '<h2>BulkDM – Get User Token</h2>' +
    '<p>Click below to authorize. You will get a User Token to paste into the sheet.</p>' +
    '<p style="margin:24px 0;"><a href="' + addUrl + '" style="display:inline-block; background:#4A154B; color:#fff; padding:12px 24px; text-decoration:none; border-radius:4px;">Add to Slack</a></p></body>'
  ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function exchangeSlackOAuthCode(code, redirectUri, clientId, clientSecret) {
  try {
    const options = {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: 'client_id=' + encodeURIComponent(clientId) + '&client_secret=' + encodeURIComponent(clientSecret) +
        '&code=' + encodeURIComponent(code) + '&redirect_uri=' + encodeURIComponent(redirectUri),
      muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch('https://slack.com/api/oauth.v2.access', options);
    const data = JSON.parse(response.getContentText() || '{}');
    if (!data.ok) {
      return HtmlService.createHtmlOutput(
        '<body style="font-family:sans-serif; padding:20px;"><h2>Error</h2><p>' + (data.error || 'unknown') + '</p><p><a href="' + redirectUri + '">Try again</a></p></body>'
      ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    const userToken = (data.authed_user && data.authed_user.access_token) ? data.authed_user.access_token : '';
    const escapedToken = userToken.replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/</g, '\\u003c');
    let body = '<body style="font-family:sans-serif; padding:24px;"><h2>BulkDM – Connected</h2>';
    if (userToken) {
      body += '<p>If the sheet showed "You are connected," you can close this window.</p>' +
        '<p><strong>If the sheet did not update:</strong> Copy the token below. In the sheet open <strong>Slack Tools → Connect to Slack</strong>, paste into the box, and click <strong>Save token</strong>.</p>' +
        '<p style="background:#f5f5f5; padding:12px; word-break:break-all;" id="tok">' + userToken + '</p>' +
        '<p><button onclick="navigator.clipboard.writeText(document.getElementById(\'tok\').textContent); this.textContent=\'Copied!\';">Copy token</button></p>';
      body += '<script>' +
        '(function(){ var t=\'' + escapedToken + '\'; var p={type:\'slack_token\',token:t}; var n=0; function send(){ if(window.opener){ try{ window.opener.postMessage(p,\'*\'); }catch(e){} } n++; if(n<7) setTimeout(send,400); else setTimeout(function(){ window.close(); },600); } send(); })();' +
        '</script>';
    } else {
      body += '<p>No user token in response. Ensure your Slack app has user scopes: chat:write, im:write, users:read, im:read, im:history.</p>';
    }
    body += '</body>';
    return HtmlService.createHtmlOutput(body).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput('<body style="font-family:sans-serif; padding:20px;"><h2>Error</h2><p>' + String(err) + '</p></body>').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * Sets up configuration section and column headers in the sheet.
 * Called automatically when the sheet loads (onOpen). Layout and config rows are auto-created.
 */
function setupHeaders() {
  const sheet = SpreadsheetApp.getActiveSheet();

  // Data tab: auto-filled with current sheet name; user only changes if data lives in another tab
  const tabNameLabel = sheet.getRange(ROW_TAB_NAME_LABEL, COLUMN_RECIPIENT).getValue();
  if (!tabNameLabel || tabNameLabel.toString().trim() !== 'TAB NAME:') {
    sheet.getRange(ROW_TAB_NAME_LABEL, COLUMN_RECIPIENT).setValue('TAB NAME:');
    sheet.getRange(ROW_TAB_NAME_LABEL, COLUMN_RECIPIENT).setFontWeight('bold');

    const currentTabName = sheet.getRange(ROW_TAB_NAME_VALUE, COLUMN_RECIPIENT).getValue();
    if (!currentTabName || currentTabName.toString().trim() === '') {
      sheet.getRange(ROW_TAB_NAME_VALUE, COLUMN_RECIPIENT).setValue(sheet.getName());
    }

    const inputCell = sheet.getRange(ROW_TAB_NAME_VALUE, COLUMN_RECIPIENT);
    inputCell.setBackground('#e8eaed');
    inputCell.setNote('Auto-filled with this tab. Change only if your recipient data is in a different tab.');
  }

  // Token type: set automatically from connected token (Bot vs User); display-only
  const tokenTypeLabel = sheet.getRange(ROW_TOKEN_TYPE_LABEL, COLUMN_RECIPIENT).getValue();
  if (!tokenTypeLabel || tokenTypeLabel.toString().trim() !== 'TOKEN TYPE:') {
    sheet.getRange(ROW_TOKEN_TYPE_LABEL, COLUMN_RECIPIENT).setValue('TOKEN TYPE:');
    sheet.getRange(ROW_TOKEN_TYPE_LABEL, COLUMN_RECIPIENT).setFontWeight('bold');

    const tokenTypeCell = sheet.getRange(ROW_TOKEN_TYPE_VALUE, COLUMN_RECIPIENT);
    tokenTypeCell.setBackground('#e8eaed');
    tokenTypeCell.setFontColor('#5f6368');
    tokenTypeCell.setNote('Set automatically when you connect to Slack. Bot = messages from the app; User = messages from your Slack account.');
    // Sync value from current token (xoxp = User, xoxb = Bot)
    const token = getSlackToken();
    if (token && token.length > 0) {
      const tokenType = token.indexOf('xoxp-') === 0 ? 'User' : 'Bot';
      tokenTypeCell.setValue(tokenType);
    } else {
      if (!tokenTypeCell.getValue() || tokenTypeCell.getValue().toString().trim() === '') {
        tokenTypeCell.setValue('Bot');
      }
    }
  } else {
    // Already have label; still sync token type from token when we have one and keep display-only formatting
    const token = getSlackToken();
    if (token && token.length > 0) {
      const tokenType = token.indexOf('xoxp-') === 0 ? 'User' : 'Bot';
      sheet.getRange(ROW_TOKEN_TYPE_VALUE, COLUMN_RECIPIENT).setValue(tokenType);
    }
    const tokenTypeCell = sheet.getRange(ROW_TOKEN_TYPE_VALUE, COLUMN_RECIPIENT);
    tokenTypeCell.setBackground('#e8eaed');
    tokenTypeCell.setFontColor('#5f6368');
    tokenTypeCell.setNote('Set automatically when you connect to Slack. Bot = messages from the app; User = messages from your Slack account.');
    try { tokenTypeCell.clearDataValidations(); } catch (e) {} // Remove dropdown so it stays display-only
  }
  const headerRow = sheet.getRange(ROW_HEADERS, 1, 1, 6).getValues()[0];
  const hasHeaders = headerRow[0] && headerRow[0].toString().trim().length > 0;

  if (!hasHeaders) {
    // Set up headers (including Slack ID and Last Sent Timestamp)
    sheet.getRange(ROW_HEADERS, COLUMN_RECIPIENT).setValue('Recipient Email');
    sheet.getRange(ROW_HEADERS, COLUMN_MESSAGE).setValue('Message');
    sheet.getRange(ROW_HEADERS, COLUMN_STATUS).setValue('Status');
    sheet.getRange(ROW_HEADERS, COLUMN_RESPONSE).setValue('Response');
    sheet.getRange(ROW_HEADERS, COLUMN_SLACK_ID).setValue('Slack ID');
    sheet.getRange(ROW_HEADERS, COLUMN_LAST_SENT_TS).setValue('Last Sent Timestamp');

    // Format header row (bold, background color) for user-input columns A–B
    const headerRange = sheet.getRange(ROW_HEADERS, 1, 1, 2);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');

    // Status and Response headers (gray, auto-filled)
    const statusResponseRange = sheet.getRange(ROW_HEADERS, COLUMN_STATUS, 1, 2);
    statusResponseRange.setFontWeight('bold');
    statusResponseRange.setBackground('#e8eaed');
    statusResponseRange.setFontColor('#5f6368');
    sheet.getRange(ROW_HEADERS, COLUMN_STATUS).setNote('Filled automatically when messages are sent.');
    sheet.getRange(ROW_HEADERS, COLUMN_RESPONSE).setNote('Filled when you call Slack Tools → Read Responses. Only shows messages the recipient sent after your BulkDM message (excluding the BulkDM message itself).');

    // Recipient column note
    sheet.getRange(ROW_HEADERS, COLUMN_RECIPIENT).setNote('Email is fastest but you can also add their full name as displayed in Slack. Name will take about ~1–2 minutes the first time you run the script.');

    // Slack ID column: distinct style to show it's auto-filled (no user input needed)
    const slackIdHeader = sheet.getRange(ROW_HEADERS, COLUMN_SLACK_ID);
    slackIdHeader.setFontWeight('bold');
    slackIdHeader.setBackground('#e8eaed');
    slackIdHeader.setFontColor('#5f6368');
    slackIdHeader.setNote('Filled automatically when a message is sent. Leave blank—you don\'t need to enter anything here.\n\nTo speed up, you can find a user\'s Slack ID from their profile (click the three dots).');

    // Last Sent Timestamp column
    const lastSentTsHeader = sheet.getRange(ROW_HEADERS, COLUMN_LAST_SENT_TS);
    lastSentTsHeader.setFontWeight('bold');
    lastSentTsHeader.setBackground('#e8eaed');
    lastSentTsHeader.setFontColor('#5f6368');
    lastSentTsHeader.setNote('Filled when you send. Read Responses only shows messages the recipient sent after this time (excluding the BulkDM message itself).');
    lastSentTsHeader.setNumberFormat('@'); // Keep as text so precision is preserved for filtering

    // Light gray background on Slack ID and Last Sent Timestamp data cells
    sheet.getRange(ROW_DATA_START, COLUMN_SLACK_ID, ROW_DATA_START + 199, COLUMN_SLACK_ID).setBackground('#f1f3f4');
    sheet.getRange(ROW_DATA_START, COLUMN_LAST_SENT_TS, ROW_DATA_START + 199, COLUMN_LAST_SENT_TS).setBackground('#f1f3f4').setNumberFormat('@');

    // Set column widths for better readability
    sheet.setColumnWidth(COLUMN_RECIPIENT, 200); // Recipient Email
    sheet.setColumnWidth(COLUMN_MESSAGE, 400);   // Message
    sheet.setColumnWidth(COLUMN_STATUS, 150);   // Status
    sheet.setColumnWidth(COLUMN_RESPONSE, 500);  // Response
    sheet.setColumnWidth(COLUMN_SLACK_ID, 120);  // Slack ID
    sheet.setColumnWidth(COLUMN_LAST_SENT_TS, 130); // Last Sent Timestamp

    // Freeze header row
    sheet.setFrozenRows(ROW_HEADERS);
  } else {
    if (!headerRow[4] || headerRow[4].toString().trim() !== 'Slack ID') {
      // Existing sheet: add Slack ID column if missing
      const slackIdHeader = sheet.getRange(ROW_HEADERS, COLUMN_SLACK_ID);
      slackIdHeader.setValue('Slack ID');
      slackIdHeader.setFontWeight('bold');
      slackIdHeader.setBackground('#e8eaed');
      slackIdHeader.setFontColor('#5f6368');
      slackIdHeader.setNote('Filled automatically when a message is sent. Leave blank—you don\'t need to enter anything here.');
      sheet.setColumnWidth(COLUMN_SLACK_ID, 120);
      sheet.getRange(ROW_DATA_START, COLUMN_SLACK_ID, ROW_DATA_START + 199, COLUMN_SLACK_ID).setBackground('#f1f3f4');
    }
    if (!headerRow[5] || headerRow[5].toString().trim() !== 'Last Sent Timestamp') {
      const lastSentTsHeader = sheet.getRange(ROW_HEADERS, COLUMN_LAST_SENT_TS);
      lastSentTsHeader.setValue('Last Sent Timestamp');
      lastSentTsHeader.setFontWeight('bold');
      lastSentTsHeader.setBackground('#e8eaed');
      lastSentTsHeader.setFontColor('#5f6368');
      lastSentTsHeader.setNote('Filled when you send. Read Responses only shows messages the recipient sent after this time (excluding the BulkDM message itself).');
      lastSentTsHeader.setNumberFormat('@');
      sheet.setColumnWidth(COLUMN_LAST_SENT_TS, 130);
      sheet.getRange(ROW_DATA_START, COLUMN_LAST_SENT_TS, ROW_DATA_START + 199, COLUMN_LAST_SENT_TS).setNumberFormat('@');
      sheet.getRange(ROW_DATA_START, COLUMN_LAST_SENT_TS, ROW_DATA_START + 199, COLUMN_LAST_SENT_TS).setBackground('#f1f3f4');
    }
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
      `The data tab name is usually auto-filled when the sheet loads.\n\nEnter the name of the tab where your recipient data lives in row ${ROW_TAB_NAME_VALUE}, column A (often the same as the current tab).`,
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
        `The tab "${tabNameStr}" does not exist.\n\nCheck the tab name in row ${ROW_TAB_NAME_VALUE}, column A. It is usually auto-filled with the current tab name.`,
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
 * Prefers token stored via "Connect to Slack" (Document Properties); falls back to SLACK_BOT_TOKEN.
 */
function getSlackToken() {
  const docProps = PropertiesService.getDocumentProperties();
  const storedToken = docProps.getProperty(SLACK_TOKEN_PROPERTY);
  if (storedToken && storedToken.trim().length > 0) {
    return storedToken.trim();
  }
  if (SLACK_BOT_TOKEN === 'YOUR_SLACK_BOT_TOKEN') {
    return null;
  }
  return SLACK_BOT_TOKEN;
}

/**
 * Saves the Slack token to Document Properties (used after Connect to Slack or "I've connected Slack").
 * Sets TOKEN TYPE in the sheet to User or Bot based on the token (xoxp = User, xoxb = Bot).
 */
function saveSlackToken(token) {
  if (!token || String(token).trim().length === 0) {
    return { ok: false, error: 'Token is empty.' };
  }
  const t = String(token).trim();
  PropertiesService.getDocumentProperties().setProperty(SLACK_TOKEN_PROPERTY, t);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet) {
    try {
      const tokenType = t.indexOf('xoxp-') === 0 ? 'User' : 'Bot';
      sheet.getRange(ROW_TOKEN_TYPE_VALUE, COLUMN_RECIPIENT).setValue(tokenType);
    } catch (e) {}
  }
  return { ok: true };
}

/**
 * Returns connection status for the Connect to Slack sidebar.
 */
function getConnectionStatus() {
  const token = getSlackToken();
  if (!token) {
    return { connected: false, message: 'Not connected' };
  }
  return { connected: true, message: 'Connected' };
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
  const token = getSlackToken();
  if (!token) {
    ui.alert('Connect to Slack',
      'No Slack token found. Use Slack Tools → Connect to Slack, or set SLACK_BOT_TOKEN in Extensions → Apps Script.',
      ui.ButtonSet.OK);
    return;
  }
  
  // Get data from sheet (starting from ROW_DATA_START)
  const lastRow = sheet.getLastRow();
  if (lastRow < ROW_DATA_START) {
    ui.alert('No Data', `Please add recipient emails and messages starting from row ${ROW_DATA_START}.`, ui.ButtonSet.OK);
    return;
  }
  
  const dataRange = sheet.getRange(ROW_DATA_START, 1, lastRow - ROW_HEADERS, 6);
  const values = dataRange.getValues();

  if (values.length === 0) {
    ui.alert('No Data', `Please add recipient emails and messages starting from row ${ROW_DATA_START}.`, ui.ButtonSet.OK);
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

  // Rows with recipient + message; which have a saved Slack ID vs need lookup
  var recipientsWithMessage = [];
  var recipientsNeedingLookup = [];
  for (var r = 0; r < values.length; r++) {
    var rec = values[r][COLUMN_RECIPIENT - 1];
    var msg = values[r][COLUMN_MESSAGE - 1];
    var savedSlackId = values[r][COLUMN_SLACK_ID - 1];
    if (rec && msg) {
      recipientsWithMessage.push(rec);
      if (!looksLikeSlackUserId(savedSlackId)) {
        recipientsNeedingLookup.push(rec);
      }
    }
    sheet.getRange(ROW_DATA_START + r, COLUMN_STATUS).setValue('⏳ Finding recipients...');
    SpreadsheetApp.flush();
  }

  var slackUsers = null;
  var lookupMap = null;
  var emailToIdMap = {};

  if (recipientsNeedingLookup.length > 0) {
    var allLookLikeEmail = recipientsNeedingLookup.length > 0 && recipientsNeedingLookup.every(function(r) { return looksLikeEmail(r); });
    if (allLookLikeEmail) {
      try {
        SpreadsheetApp.getActiveSpreadsheet().toast('Looking up recipients by email...', 'BulkDM', 45);
      } catch (e) {}
      var seen = {};
      for (var e = 0; e < recipientsNeedingLookup.length; e++) {
        var em = String(recipientsNeedingLookup[e]).trim().toLowerCase();
        if (seen[em]) continue;
        seen[em] = true;
        var uid = lookupSlackUserByEmail(em);
        if (uid) emailToIdMap[em] = uid;
        Utilities.sleep(1100);
      }
    } else {
      try {
        SpreadsheetApp.getActiveSpreadsheet().toast('Finding recipients...', 'BulkDM', 45);
      } catch (e) {}
      var result = getAllSlackUsersUntilMatched(recipientsNeedingLookup);
      if (!result) {
        ui.alert('Error', 'Failed to connect to Slack. Please check your ' + tokenType + ' token.', ui.ButtonSet.OK);
        return;
      }
      slackUsers = result.users;
      lookupMap = result.lookupMap;
    }

    if (allLookLikeEmail && Object.keys(emailToIdMap).length === 0 && recipientsNeedingLookup.length > 0) {
      ui.alert('Error', 'Could not find any recipients by email. Add users:read.email to your Slack app scopes and re-connect (Connect to Slack), or use full names instead of emails.', ui.ButtonSet.OK);
      return;
    }
  } else {
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast('Using saved Slack IDs...', 'BulkDM', 3);
    } catch (e) {}
  }
  
  // Process each row
  let successCount = 0;
  let errorCount = 0;

  for (let i = 0; i < values.length; i++) {
    const row = ROW_DATA_START + i; // Sheet row number (1-indexed)
    const recipient = values[i][COLUMN_RECIPIENT - 1];
    const message = values[i][COLUMN_MESSAGE - 1];
    const savedSlackId = values[i][COLUMN_SLACK_ID - 1];

    // Skip empty rows
    if (!recipient || !message) {
      sheet.getRange(row, COLUMN_STATUS).setValue('⏭️ Skipped (empty)');
      continue;
    }

    // Update status to processing
    sheet.getRange(row, COLUMN_STATUS).setValue('⏳ Sending...');
    SpreadsheetApp.flush(); // Force update

    // Use saved Slack ID if present and valid; otherwise resolve from sheet
    var userId = null;
    if (looksLikeSlackUserId(savedSlackId)) {
      userId = String(savedSlackId).trim();
    } else if (recipientsNeedingLookup.length > 0) {
      if (slackUsers && lookupMap) {
        userId = findSlackUserFast(recipient, slackUsers, lookupMap);
      } else {
        userId = emailToIdMap[recipient.toString().trim().toLowerCase()] || null;
      }
    }

    if (!userId) {
      sheet.getRange(row, COLUMN_STATUS).setValue('❌ User not found');
      errorCount++;
      continue;
    }

    // Send message
    const result = sendSlackDM(userId, message);

    if (result.success) {
      sheet.getRange(row, COLUMN_STATUS).setValue('✅ Sent');
      sheet.getRange(row, COLUMN_SLACK_ID).setValue(userId); // Save for next time
      if (result.ts) {
        const tsCell = sheet.getRange(row, COLUMN_LAST_SENT_TS);
        tsCell.setNumberFormat('@');
        tsCell.setValue(result.ts); // Read Responses only shows recipient messages after this (excluding this message)
      }
      successCount++;
    } else {
      const friendlyError = getSlackErrorMessage(result.error);
      sheet.getRange(row, COLUMN_STATUS).setValue(`❌ Error: ${friendlyError}`);
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
    ui.alert('Connect to Slack',
      'No Slack token found. Use Slack Tools → Connect to Slack, or set SLACK_BOT_TOKEN in Extensions → Apps Script.',
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
  
  // Get data from sheet (starting from ROW_DATA_START)
  const lastRow = sheet.getLastRow();
  if (lastRow < ROW_DATA_START) {
    ui.alert('No Data', `Please add recipient emails starting from row ${ROW_DATA_START}.`, ui.ButtonSet.OK);
    return;
  }
  
  const dataRange = sheet.getRange(ROW_DATA_START, 1, lastRow - ROW_HEADERS, 6);
  const values = dataRange.getValues();

  if (values.length === 0) {
    ui.alert('No Data', `Please add recipient emails starting from row ${ROW_DATA_START}.`, ui.ButtonSet.OK);
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

  // Rows with Slack ID use quick path (conversations.open + history). Others need full DM list + user list.
  const rowsWithSlackId = [];
  const rowsWithoutSlackIdSet = {};
  for (let i = 0; i < values.length; i++) {
    const recipient = values[i][COLUMN_RECIPIENT - 1];
    const slackId = values[i][COLUMN_SLACK_ID - 1];
    const lastSentTs = values[i][COLUMN_LAST_SENT_TS - 1];
    const rowNum = ROW_DATA_START + i;
    if (!recipient) continue;
    if (looksLikeSlackUserId(slackId)) {
      rowsWithSlackId.push({ rowNum, slackId: String(slackId).trim(), lastSentTs });
    } else {
      rowsWithoutSlackIdSet[rowNum] = true;
    }
  }

  let foundCount = 0;
  let responseCount = 0;

  // Quick path: use saved Slack ID to open DM and read history (no need to list all DMs or users)
  if (rowsWithSlackId.length > 0) {
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast('Reading responses (by Slack ID)...', 'BulkDM', 5);
    } catch (e) {}
    const slackIdToChannelId = {};
    for (const { rowNum, slackId, lastSentTs } of rowsWithSlackId) {
      let channelId = slackIdToChannelId[slackId];
      if (!channelId) {
        channelId = getChannelIdForUser(slackId);
        if (channelId) slackIdToChannelId[slackId] = channelId;
      }
      sheet.getRange(rowNum, COLUMN_RESPONSE).setValue('⏳ Reading...');
      SpreadsheetApp.flush();
      if (!channelId) {
        sheet.getRange(rowNum, COLUMN_RESPONSE).setValue('❌ Error reading');
        continue;
      }
      const messages = getConversationHistory(channelId);
      if (!messages) {
        sheet.getRange(rowNum, COLUMN_RESPONSE).setValue('❌ Error reading');
        continue;
      }
      const sentTsStr = lastSentTs !== undefined && lastSentTs !== null && lastSentTs !== '' ? (typeof lastSentTs === 'number' ? lastSentTs.toFixed(6) : String(lastSentTs).trim()) : '';
      // Only messages after our sent time: use string comparison so we exclude the BulkDM message (exact ts match) without float rounding issues
      const recipientMessages = messages.filter(msg =>
        msg.user === slackId && msg.text && !msg.subtype &&
        (!sentTsStr || (msg.ts > sentTsStr))
      );
      foundCount++;
      if (recipientMessages.length === 0) {
        sheet.getRange(rowNum, COLUMN_RESPONSE).setValue('No response yet');
      } else {
        recipientMessages.sort((a, b) => parseFloat(b.ts) - parseFloat(a.ts));
        recipientMessages.reverse();
        const allResponses = recipientMessages.map(msg => msg.text).join('\n\n---\n\n');
        sheet.getRange(rowNum, COLUMN_RESPONSE).setValue(allResponses);
        responseCount += recipientMessages.length;
      }
      Utilities.sleep(200);
    }
  }

  // Fallback: rows without Slack ID — list all DMs and match by name/email
  if (Object.keys(rowsWithoutSlackIdSet).length > 0) {
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast('Loading workspace members...', 'BulkDM', 15);
    } catch (e) {}
    const slackUsers = getAllSlackUsers();
    if (!slackUsers) {
      ui.alert('Error', `Failed to connect to Slack. Please check your ${tokenType} token.`, ui.ButtonSet.OK);
      return;
    }
    const userMap = {};
    for (const user of slackUsers) {
      userMap[user.id] = user;
    }
    const recipientMapNoSlackId = {};
    for (let i = 0; i < values.length; i++) {
      const recipient = values[i][COLUMN_RECIPIENT - 1];
      const slackId = values[i][COLUMN_SLACK_ID - 1];
      const rowNum = ROW_DATA_START + i;
      if (recipient && !looksLikeSlackUserId(slackId)) {
        recipientMapNoSlackId[recipient.toString().trim().toLowerCase()] = rowNum;
      }
    }
    const dmResult = getAllDMConversations();
    if (dmResult === null) {
      if (foundCount === 0) {
        ui.alert('No Conversations', 'No DM conversations found.', ui.ButtonSet.OK);
      }
    } else if (typeof dmResult === 'object' && dmResult.error) {
      const hint = dmResult.error === 'missing_scope'
        ? '\n\nAdd im:read to your Slack app (api.slack.com → OAuth & Permissions): under Bot Token Scopes if using a Bot token, or User Token Scopes if using a User token. Then re-connect (Slack Tools → Connect to Slack) or re-paste the token.'
        : '';
      ui.alert('Cannot list DMs', 'Slack returned: ' + dmResult.error + hint, ui.ButtonSet.OK);
      return;
    }
    const dmConversations = Array.isArray(dmResult) ? dmResult : [];
    for (const conversation of dmConversations) {
      const otherUserId = conversation.user;
      if (!otherUserId || otherUserId === senderUserId) continue;
      const otherUser = userMap[otherUserId];
      if (!otherUser) continue;
      let matchedRow = null;
      const userEmail = (otherUser.profile && otherUser.profile.email) ? otherUser.profile.email.toLowerCase() : '';
      const userName = otherUser.name ? otherUser.name.toLowerCase() : '';
      const userDisplayName = (otherUser.profile && otherUser.profile.display_name) ? otherUser.profile.display_name.toLowerCase() : '';
      const userRealName = (otherUser.profile && otherUser.profile.real_name) ? otherUser.profile.real_name.toLowerCase() : '';
      for (const [key, rowNum] of Object.entries(recipientMapNoSlackId)) {
        if (key === userEmail || key === userName || key === userDisplayName || key === userRealName ||
            (userRealName && (userRealName.includes(key) || key.includes(userRealName)))) {
          matchedRow = rowNum;
          break;
        }
      }
      if (!matchedRow || !rowsWithoutSlackIdSet[matchedRow]) continue;
      foundCount++;
      sheet.getRange(matchedRow, COLUMN_RESPONSE).setValue('⏳ Reading...');
      SpreadsheetApp.flush();
      const messages = getConversationHistory(conversation.id);
      if (!messages) {
        sheet.getRange(matchedRow, COLUMN_RESPONSE).setValue('❌ Error reading');
        continue;
      }
      const lastSentTs = values[matchedRow - ROW_DATA_START][COLUMN_LAST_SENT_TS - 1];
      const sentTsStr = lastSentTs !== undefined && lastSentTs !== null && lastSentTs !== '' ? (typeof lastSentTs === 'number' ? lastSentTs.toFixed(6) : String(lastSentTs).trim()) : '';
      // Only messages after our sent time: use string comparison so we exclude the BulkDM message (exact ts match) without float rounding issues
      const recipientMessages = messages.filter(msg =>
        msg.user === otherUserId && msg.text && !msg.subtype &&
        (!sentTsStr || (msg.ts > sentTsStr))
      );
      if (recipientMessages.length === 0) {
        sheet.getRange(matchedRow, COLUMN_RESPONSE).setValue('No response yet');
      } else {
        recipientMessages.sort((a, b) => parseFloat(b.ts) - parseFloat(a.ts));
        recipientMessages.reverse();
        const allResponses = recipientMessages.map(msg => msg.text).join('\n\n---\n\n');
        sheet.getRange(matchedRow, COLUMN_RESPONSE).setValue(allResponses);
        responseCount += recipientMessages.length;
      }
      Utilities.sleep(200);
    }
  }

  // Show completion message
  ui.alert(
    'Complete!',
    `Responses read:\n✅ Conversations checked: ${foundCount}\n💬 Total messages found: ${responseCount}`,
    ui.ButtonSet.OK
  );
}

/**
 * Test Slack connection (quick check via auth.test; does not load full user list)
 */
function testSlackConnection() {
  const ui = SpreadsheetApp.getUi();
  
  if (!getSlackToken()) {
    ui.alert('Connect to Slack',
      'No Slack token found. Use Slack Tools → Connect to Slack, or set SLACK_BOT_TOKEN in Extensions → Apps Script.',
      ui.ButtonSet.OK);
    return;
  }
  
  const info = getBotInfo();
  if (info) {
    const tokenType = getTokenType();
    ui.alert(
      'Connection successful',
      'Your Slack connection is valid.\n\n' +
      'Messages will be sent from your ' + (tokenType === 'User' ? 'Slack account' : 'bot') + '.\n\n' +
      'Use Slack Tools → Send Messages to send DMs, and Read Responses to pull replies into the sheet.',
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      'Connection failed',
      'Could not connect to Slack. Please check:\n\n' +
      '1. Your token is correct (User: xoxp-... or Bot: xoxb-...)\n' +
      '2. Required scopes: chat:write, im:write, users:read, im:read, im:history\n' +
      '3. App is installed in your workspace\n' +
      '4. Token is not revoked or expired',
      ui.ButtonSet.OK
    );
  }
}

// ============================================
// SLACK API FUNCTIONS
// ============================================

/**
 * Returns a user-friendly message for known Slack API errors (e.g. when sending DMs).
 * Falls back to the raw error if unknown.
 */
function getSlackErrorMessage(slackError) {
  if (!slackError) return slackError;
  const code = slackError.toString().trim();
  const known = {
    'messages_tab_disabled': 'Recipient has DMs disabled or restricted. Ask them to allow messages from apps (Settings → Privacy) or skip this recipient.',
    'channel_not_found': 'Channel or user not found.',
    'user_not_found': 'User not found.',
    'not_in_channel': 'Bot is not in the channel.',
    'invalid_auth': 'Invalid token. Check your Slack token.',
    'account_inactive': 'Recipient\'s account is deactivated.',
    'user_is_restricted': 'Recipient is a restricted user (guest); some actions are not allowed.',
    'cant_dm_bot': 'Cannot open a DM with a bot user.',
    'ratelimited': 'Rate limited by Slack. Wait a minute and try again.'
  };
  return known[code] || code;
}

/**
 * Get all users from Slack workspace (with pagination)
 * Slack's users.list returns max 200 per request; we follow cursor to get all users.
 * Handles 429 rate limit by waiting per Retry-After and retrying.
 */
function getAllSlackUsers() {
  try {
    const token = getSlackToken();
    if (!token) {
      return null;
    }

    const allMembers = [];
    let cursor = '';

    do {
      const url = 'https://slack.com/api/users.list';
      const params = {
        limit: 200,
        cursor: cursor || undefined
      };
      const queryString = Object.keys(params)
        .filter(k => params[k] !== undefined && params[k] !== '')
        .map(k => encodeURIComponent(k) + '=' + encodeURIComponent(params[k]))
        .join('&');

      const options = {
        method: 'get',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true,
        timeout: 15
      };

      const fullUrl = queryString ? url + '?' + queryString : url;
      let response = UrlFetchApp.fetch(fullUrl, options);
      const responseCode = response.getResponseCode();

      // Handle rate limit: wait and retry same request
      if (responseCode === 429) {
        const retryAfter = response.getHeaders()['Retry-After'] || response.getHeaders()['retry-after'];
        const waitSeconds = retryAfter ? Math.max(2, parseInt(retryAfter, 10)) : 3;
        Logger.log('Slack rate limited (429), waiting ' + waitSeconds + 's before retry');
        try {
          SpreadsheetApp.getActiveSpreadsheet().toast('Slack rate limit – waiting ' + waitSeconds + 's...', 'BulkDM', 5);
        } catch (e) {}
        Utilities.sleep(waitSeconds * 1000);
        response = UrlFetchApp.fetch(fullUrl, options);
        if (response.getResponseCode() === 429) {
          Logger.log('Slack API Error: still rate limited after retry');
          return null;
        }
      }

      const data = JSON.parse(response.getContentText());

      if (!data.ok) {
        Logger.log('Slack API Error: ' + data.error);
        return null;
      }

      if (data.members && data.members.length > 0) {
        allMembers.push.apply(allMembers, data.members);
      }

      if (MAX_USERS_TO_FETCH > 0 && allMembers.length >= MAX_USERS_TO_FETCH) {
        break;
      }

      cursor = (data.response_metadata && data.response_metadata.next_cursor) ? data.response_metadata.next_cursor : '';

      if (cursor) {
        Utilities.sleep(1200);
      }
    } while (cursor);

    return allMembers.filter(user => !user.deleted && !user.is_bot);
  } catch (error) {
    Logger.log('Error fetching users: ' + error);
    return null;
  }
}

/**
 * Fetches users page-by-page (200 per page) and stops as soon as all recipients are matched.
 * Returns { users: array, lookupMap: object } or null.
 */
function getAllSlackUsersUntilMatched(recipientList) {
  var token = getSlackToken();
  if (!token) return null;
  var normalizedRecipients = [];
  for (var i = 0; i < recipientList.length; i++) {
    var r = recipientList[i];
    if (r && String(r).trim()) normalizedRecipients.push(String(r).trim());
  }
  if (normalizedRecipients.length === 0) return { users: [], lookupMap: {} };

  function allMatched(users, lookupMap) {
    for (var i = 0; i < normalizedRecipients.length; i++) {
      if (!findSlackUserFast(normalizedRecipients[i], users, lookupMap)) return false;
    }
    return true;
  }

  try {
    var filtered = [];
    var lookupMap = {};
    var cursor = '';
    var pageCount = 0;

    do {
      pageCount++;
      var url = 'https://slack.com/api/users.list';
      var params = { limit: 200 };
      if (cursor) params.cursor = cursor;
      var queryString = Object.keys(params)
        .filter(function(k) { return params[k] !== undefined && params[k] !== ''; })
        .map(function(k) { return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]); })
        .join('&');
      var options = {
        method: 'get',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        muteHttpExceptions: true,
        timeout: 15
      };
      var fullUrl = queryString ? url + '?' + queryString : url;
      var response = UrlFetchApp.fetch(fullUrl, options);
      var responseCode = response.getResponseCode();

      if (responseCode === 429) {
        var retryAfter = response.getHeaders()['Retry-After'] || response.getHeaders()['retry-after'];
        var waitSeconds = retryAfter ? Math.max(2, parseInt(retryAfter, 10)) : 3;
        try { SpreadsheetApp.getActiveSpreadsheet().toast('Slack rate limit – waiting ' + waitSeconds + 's...', 'BulkDM', 5); } catch (e) {}
        Utilities.sleep(waitSeconds * 1000);
        response = UrlFetchApp.fetch(fullUrl, options);
        if (response.getResponseCode() === 429) return null;
      }

      var data = JSON.parse(response.getContentText());
      if (!data.ok) {
        Logger.log('Slack API Error: ' + data.error);
        return null;
      }

      if (data.members && data.members.length > 0) {
        var pageFiltered = data.members.filter(function(u) { return !u.deleted && !u.is_bot; });
        for (var p = 0; p < pageFiltered.length; p++) {
          filtered.push(pageFiltered[p]);
          addUserToLookupMap(lookupMap, pageFiltered[p]);
        }
        if (allMatched(filtered, lookupMap)) {
          try { SpreadsheetApp.getActiveSpreadsheet().toast('All recipients found after ' + pageCount + ' page(s).', 'BulkDM', 3); } catch (e) {}
          return { users: filtered, lookupMap: lookupMap };
        }
      }

      cursor = (data.response_metadata && data.response_metadata.next_cursor) ? data.response_metadata.next_cursor : '';
      if (cursor) Utilities.sleep(1200);
    } while (cursor);

    return { users: filtered, lookupMap: lookupMap };
  } catch (e) {
    Logger.log('getAllSlackUsersUntilMatched: ' + e);
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
    if (user.profile && user.profile.email) {
      if (user.profile.email.toLowerCase() === searchTerm) {
        return user.id;
      }
    }
    if (user.profile && user.profile.display_name) {
      if (user.profile.display_name.toLowerCase() === searchTerm) {
        return user.id;
      }
    }
    if (user.profile && user.profile.real_name) {
      if (user.profile.real_name.toLowerCase() === searchTerm) {
        return user.id;
      }
    }
    if (user.name) {
      if (user.name.toLowerCase() === searchTerm.replace('@', '')) {
        return user.id;
      }
    }
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
 * Build a single lookup map from all users for O(1) recipient resolution.
 * Keys: normalized email, display_name, real_name, username, and each word in real_name.
 */
function buildUserLookupMap(slackUsers) {
  var map = {};
  if (!slackUsers) return map;
  for (var u = 0; u < slackUsers.length; u++) {
    addUserToLookupMap(map, slackUsers[u]);
  }
  return map;
}

/**
 * Add one user's keys to an existing lookup map (only sets if key not already present).
 * Used to build the map incrementally page-by-page so we don't rebuild from scratch each time.
 */
function addUserToLookupMap(map, user) {
  if (!user || !map) return;
  var id = user.id;
  var add = function(key) {
    if (key && key.length > 0 && map[key] === undefined) {
      map[key] = id;
    }
  };
  var email = user.profile && user.profile.email ? user.profile.email.toLowerCase().trim() : '';
  var displayName = user.profile && user.profile.display_name ? user.profile.display_name.toLowerCase().trim() : '';
  var realName = user.profile && user.profile.real_name ? user.profile.real_name.toLowerCase().trim() : '';
  var name = user.name ? user.name.toLowerCase().replace('@', '').trim() : '';
  add(email);
  add(displayName);
  add(realName);
  add(name);
  if (realName) {
    var words = realName.split(/\s+/);
    for (var w = 0; w < words.length; w++) {
      add(words[w]);
    }
  }
}

/**
 * Fast lookup: use prebuilt map for exact match, then fall back to partial match only if needed.
 */
function findSlackUserFast(recipient, slackUsers, lookupMap) {
  if (!recipient || !slackUsers) return null;
  var searchTerm = recipient.toString().trim().toLowerCase();
  if (!searchTerm) return null;
  var withoutAt = searchTerm.replace('@', '');
  if (lookupMap[searchTerm]) return lookupMap[searchTerm];
  if (lookupMap[withoutAt]) return lookupMap[withoutAt];
  for (var i = 0; i < slackUsers.length; i++) {
    var user = slackUsers[i];
    if (user.profile && user.profile.real_name) {
      var realName = user.profile.real_name.toLowerCase();
      if (realName.indexOf(searchTerm) !== -1 || searchTerm.indexOf(realName) !== -1) {
        return user.id;
      }
    }
  }
  return null;
}

/**
 * Look up a single Slack user by email (no full user list needed).
 * Requires scope users:read.email. Returns user id or null.
 */
function lookupSlackUserByEmail(email) {
  if (!email || String(email).indexOf('@') === -1) return null;
  try {
    var token = getSlackToken();
    if (!token) return null;
    var url = 'https://slack.com/api/users.lookupByEmail?email=' + encodeURIComponent(String(email).trim());
    var options = {
      method: 'get',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true,
      timeout: 10
    };
    var response = UrlFetchApp.fetch(url, options);
    var data = JSON.parse(response.getContentText() || '{}');
    if (data.ok && data.user && data.user.id) {
      return data.user.deleted ? null : data.user.id;
    }
    return null;
  } catch (e) {
    return null;
  }
}

function looksLikeEmail(str) {
  if (!str || typeof str !== 'string') return false;
  var s = str.trim();
  return s.indexOf('@') !== -1 && /^\S+@\S+\.\S+$/.test(s);
}

/** Returns true if the value looks like a Slack user ID (e.g. U01234ABCDE or W01234ABCDE). */
function looksLikeSlackUserId(val) {
  if (val === null || val === undefined) return false;
  var s = String(val).trim();
  return s.length >= 8 && (s.indexOf('U') === 0 || s.indexOf('W') === 0);
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
 * Get all DM conversations for the bot/user.
 * Uses pagination to fetch every IM channel.
 * @return {Array|null|{error: string}} Array of channel objects, or null if no token/exception, or { error: "slack_error" } if API returned ok: false
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

    const allChannels = [];
    let cursor = '';

    do {
      const params = {
        types: 'im',
        exclude_archived: true,
        limit: 1000
      };
      if (cursor) {
        params.cursor = cursor;
      }

      const queryString = Object.keys(params).map(key =>
        encodeURIComponent(key) + '=' + encodeURIComponent(params[key])
      ).join('&');

      const response = UrlFetchApp.fetch(url + '?' + queryString, options);
      const data = JSON.parse(response.getContentText());

      if (!data.ok) {
        Logger.log('Conversations list error: ' + data.error);
        return { error: data.error || 'unknown' };
      }

      const page = data.channels || [];
      allChannels.push.apply(allChannels, page);
      cursor = (data.response_metadata && data.response_metadata.next_cursor) ? data.response_metadata.next_cursor : '';
    } while (cursor);

    return allChannels;
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
 * Get the DM channel ID for a Slack user (opens DM if needed).
 * Returns channel ID string or null.
 */
function getChannelIdForUser(slackUserId) {
  try {
    const token = getSlackToken();
    if (!token || !slackUserId) return null;
    const response = UrlFetchApp.fetch('https://slack.com/api/conversations.open', {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ users: slackUserId })
    });
    const data = JSON.parse(response.getContentText());
    if (data.ok && data.channel && data.channel.id) {
      return data.channel.id;
    }
    Logger.log('conversations.open error: ' + (data.error || 'unknown'));
    return null;
  } catch (e) {
    Logger.log('getChannelIdForUser: ' + e);
    return null;
  }
}
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
    
    // Send the message (reuse token from above)
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
      return { success: true, ts: chatData.ts || (chatData.message && chatData.message.ts) || '' };
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
