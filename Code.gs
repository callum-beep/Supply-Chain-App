/**
 * Unleashed → Google Sheets: Open Purchase Orders + Calendar Sync
 * ENHANCED: Persistent milestone completion, calendar invite removal, combined pull/create/cleanup, comment-only updates
 * NEW: Change tracking, fixed dates, status change logging, uncomplete function, unified feed, color updates, permit numbers
 * FIXED: API signature encoding, pagination, and debugging
 * UPDATED: Date parsing for /Date(timestamp)/ format
 * UPDATED: Calendar days with weekend adjustment, emoji event titles with quantity and product name, +1 day calendar fix, fixed date format
 * WEB APP: Added mobile web app compatibility
 */

////////////////////////////// CONFIG //////////////////////////////
const BASE_URL   = 'https://api.unleashedsoftware.com';
const SHEET_NAME = 'OpenPOs';
const ASSUMPTIONS_SHEET  = 'OpenPOs Assumptions';
const REQUIREMENTS_SHEET = 'OpenPOs Requirements';
const COMPLETED_EVENTS_SHEET = 'Completed Events';
const FEED_SHEET = 'Activity Feed';
const FIXED_DATES_SHEET = 'Fixed Dates';
const PERMIT_NUMBERS_SHEET = 'Permit Numbers';
const PAGE_SIZE  = 200;
const OPEN_STATUSES = ['Parked', 'Unapproved', 'Placed', 'Costed'];
const CLIENT_TYPE = 'google-apps-script/opos';
const EXCLUDED_WAREHOUSE = 'W1';

// Calendar
const CALENDAR_NAME = 'Aus Supply Chain Reminders';
const DEFAULT_TIMEZONE = 'Australia/Brisbane';

// Tagging
const SYNC_TAG = '[OpenPOs]';

// Calendar event colors
const PHASE_EVENT_COLOR = {
  'Depart Canada':         CalendarApp.EventColor.ORANGE,
  'Arrive in Aus':         CalendarApp.EventColor.BLUE, 
  'Clears Customs':        CalendarApp.EventColor.PINK,
  'Irradiation':           CalendarApp.EventColor.MAUVE,
  'Testing':               CalendarApp.EventColor.CYAN,
  'Packaging and Release': CalendarApp.EventColor.GREEN,
  'BLS Delivery':          CalendarApp.EventColor.GRAY
};

// Milestone emojis
const MILESTONE_EMOJIS = {
  'Depart Canada':         '🚢',
  'Arrive in Aus':         '🇦🇺',
  'Clears Customs':        '🛃',
  'Irradiation':           '☢️',
  'Testing':               '🧪',
  'Packaging and Release': '📦',
  'BLS Delivery':          '🚚'
};

// Sheet colors
const PHASE_SHEET_HEX = {
  'Depart Canada':         '#F6B26B',
  'Arrive in Aus':         '#9FC5E8',
  'Clears Customs':        '#F4CCCC',
  'Irradiation':           '#D9D2E9',
  'Testing':               '#CFE2F3',
  'Packaging and Release': '#B6D7A8',
  'BLS Delivery':          '#FFF2CC'
};

// Highlight colors
const COLOR_IRRADIATION_REQUIRED = '#FBE4D5';
const COLOR_PARKED_STATUS        = '#F4CCCC';
const LABEL_THIS_WEEK_HEX        = '#FFD966';
const LABEL_NEXT_WEEK_HEX        = '#FFF2CC';
const LABEL_LAST_WEEK_HEX        = '#F4CCCC';
const COLOR_COMPLETED            = '#000000';
const COLOR_COMPLETED_TEXT       = '#FFFFFF';
const COLOR_FIXED_DATE           = '#FFD700';

let __opsCountThisRun = 0;

////////////////////////////// MENU ////////////////////////////////
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('🥦 OpenPOs')
      .addItem('🔑 Set Unleashed Credentials', 'setupCredentials')
      .addItem('🔧 Diagnose Pull Issues', 'diagnosePullIssues')
      .addItem('📅 Debug Date Extraction', 'debugPODates')
      .addItem('🚀 Pull POs + Create/Update Calendar + Cleanup', 'pullAndSyncAll')
      .addSeparator()
      .addItem('📥 Pull Open Purchase Orders', 'pullOpenPurchaseOrders')
      .addItem('📅 Create/Update Calendar Events', 'syncCalendarEventsFromSheet')
      .addItem('🎨 Update Event Colors', 'updateEventColors')
      .addItem('🧹 Cleanup Old Calendar Events', 'cleanupCalendarEvents')
      .addItem('🚫 Remove Calendar Duplicates', 'removeCalendarDuplicates')
      .addItem('🗑️ Delete All Calendar Events', 'deleteAllCalendarEvents')
      .addSeparator()
      .addItem('✅ Complete Event', 'completeEventMenu')
      .addItem('↩️ Uncomplete Event', 'uncompleteEventMenu')
      .addItem('💬 Add Comment to PO', 'addCommentToPOMenu')
      .addItem('📌 Set Fixed Date', 'setFixedDateMenu')
      .addItem('❌ Clear Fixed Date', 'clearFixedDateMenu')
      .addItem('🔢 Assign Permit Number', 'assignPermitNumberMenu')
      .addItem('🔄 Refresh Requirements', 'refreshRequirementsFromOpenPOs_')
      .addItem('🗑️ Clear EventId Columns', 'clearAllEventIdColumns')
      .addToUi();
  } catch (error) {
    // Silently fail if there's no UI (like in web app context)
    console.log('onOpen: No UI available in this context');
  }
}

////////////////////////////// WEB APP FUNCTIONS //////////////////////////////

/**
 * ===== Web App bootstrap =====
 * Deploy: Deploy → New deployment → Web app
 * Execute as: Me     •     Who has access: Your org (or Anyone with link)
 */
function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('OpenPOs Mobile')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Utilities **/
function _requireSheetRows_(name) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) throw new Error(`No data found in sheet: ${name}`);
  return sh;
}

/** Mobile API: list POs (lightweight) with simple search - FIXED FOR WEB APP */
function apiListPOs(query) {
  try {
    console.log('apiListPOs called with query:', query);
    
    // CRITICAL FIX: Use getActiveSpreadsheet() for web app compatibility
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);
    
    if (!sh) {
      console.error('Sheet not found:', SHEET_NAME);
      return { success: false, error: 'Sheet not found', data: [] };
    }
    
    const lastRow = sh.getLastRow();
    console.log('Sheet last row:', lastRow);
    
    if (lastRow < 2) {
      console.log('No data in sheet');
      return { success: true, data: [] };
    }
    
    const values = sh.getDataRange().getValues();
    const h = values[0];
    
    console.log('Headers found:', h);
    console.log('Total data rows:', values.length - 1);
    
    // Validate required columns exist
    const idx = {
      order:  h.indexOf('OrderNumber'),
      status: h.indexOf('OrderStatus'),
      supplier: h.indexOf('SupplierName'),
      product: h.indexOf('ProductDescription'),
      qty:    h.indexOf('OrderedQuantity'),
      depart: h.indexOf('Depart Canada'),
      arrive: h.indexOf('Arrive in Aus'),
      customs:h.indexOf('Clears Customs'),
      irrad:  h.indexOf('Irradiation'),
      test:   h.indexOf('Testing'),
      release:h.indexOf('Packaging and Release')
    };
    
    console.log('Column indices:', idx);
    
    // Check if essential columns exist
    if (idx.order === -1) {
      console.error('OrderNumber column not found. Available columns:', h);
      return { success: false, error: 'OrderNumber column missing', data: [] };
    }
    
    const q = (query || '').toLowerCase().trim();
    const rows = [];
    
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const po = String(row[idx.order] || '');
      
      // Skip empty PO numbers
      if (!po.trim()) continue;
      
      const rec = {
        po: po,
        status:   row[idx.status]   || '',
        supplier: row[idx.supplier] || '',
        product:  row[idx.product]  || '',
        qty: Number(row[idx.qty] || 0),
        dates: {
          'Depart Canada':         toISODateCell_(row[idx.depart])  || '',
          'Arrive in Aus':         toISODateCell_(row[idx.arrive])  || '',
          'Clears Customs':        toISODateCell_(row[idx.customs]) || '',
          'Irradiation':           toISODateCell_(row[idx.irrad])   || '',
          'Testing':               toISODateCell_(row[idx.test])    || '',
          'Packaging and Release': toISODateCell_(row[idx.release]) || ''
        }
      };
      
      const hay = `${po} ${rec.status} ${rec.supplier} ${rec.product}`.toLowerCase();
      if (!q || hay.includes(q)) rows.push(rec);
    }

    console.log(`Filtered to ${rows.length} rows for query: "${q}"`);
    
    // Sort by release date, then by PO number
    rows.sort((a, b) => {
      const da = a.dates['Packaging and Release'] || a.dates['Testing'] || a.dates['Arrive in Aus'] || '';
      const db = b.dates['Packaging and Release'] || b.dates['Testing'] || b.dates['Arrive in Aus'] || '';
      return (da || '9999-12-31').localeCompare(db || '9999-12-31') || a.po.localeCompare(b.po);
    });

    const result = rows.slice(0, 300);
    console.log('Final result length:', result.length);
    
    return { success: true, data: result };

  } catch (error) {
    console.error('apiListPOs error:', error);
    return { success: false, error: error.message, data: [] };
  }
}

// WEB APP COMPATIBLE FUNCTIONS
function pullAndSyncAll() {
  try {
    console.log('Starting full sync from web app...');
    
    // Pull purchase orders
    pullOpenPurchaseOrders();
    
    // Sync calendar events
    syncCalendarEventsFromSheet();
    
    // Cleanup old calendar events
    cleanupCalendarEvents();
    
    console.log('Full sync completed successfully');
    return { success: true, message: 'Full sync completed successfully' };
  } catch (error) {
    console.error('Sync failed:', error);
    return { success: false, error: error.message };
  }
}

function completeEvent(poNumber, milestone, comments) {
  try {
    const user = Session.getEffectiveUser().getEmail();
    const result = markEventComplete_(poNumber, milestone, user, comments);
    return { success: true, data: result };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function uncompleteEvent(poNumber, milestone, reason) {
  try {
    const result = uncompleteEvent_(poNumber, milestone, reason);
    return { success: true, data: result };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function addCommentToPO(poNumber, comment) {
  try {
    const result = addCommentToPO_(poNumber, comment);
    return { success: true, data: result };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function setFixedDate(poNumber, milestone, dateStr) {
  try {
    const result = setFixedDate_(poNumber, milestone, dateStr);
    return { success: true, data: result };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function clearFixedDate(poNumber, milestone) {
  try {
    const result = clearFixedDate_(poNumber, milestone);
    return { success: true, data: result };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function assignPermitNumber(poNumber, permitNumber) {
  try {
    const result = assignPermitNumber_(poNumber, permitNumber);
    return { success: true, data: result };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function refreshRequirements() {
  try {
    refreshRequirementsFromOpenPOs_();
    return { success: true, message: 'Requirements refreshed successfully' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function updateEventColors() {
  try {
    updateEventColors();
    return { success: true, message: 'Event colors updated successfully' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function cleanupCalendar() {
  try {
    cleanupCalendarEvents();
    return { success: true, message: 'Calendar cleanup completed' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function removeCalendarDuplicates() {
  try {
    removeCalendarDuplicates();
    return { success: true, message: 'Calendar duplicates removed' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function pullOpenPurchaseOrdersWeb() {
  try {
    pullOpenPurchaseOrders();
    return { success: true, message: 'Purchase orders pulled successfully' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function syncCalendarEvents() {
  try {
    syncCalendarEventsFromSheet();
    return { success: true, message: 'Calendar events synced successfully' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function diagnoseIssues() {
  try {
    diagnosePullIssues();
    return { success: true, message: 'Diagnosis completed - check console for details' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function debugDates() {
  try {
    debugPODates();
    return { success: true, message: 'Date debugging completed - check console for details' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function clearEventIds() {
  try {
    clearAllEventIdColumns();
    return { success: true, message: 'Event IDs cleared successfully' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function apiGetPermits() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('Permit Numbers');
    const permits = [];
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, data: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      permits.push({
        poNumber: row[0] || '',
        permitNumber: row[1] || '',
        assignedBy: row[2] || '',
        timestamp: row[3] || ''
      });
    }
    
    return { success: true, data: permits };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function apiGetFeed() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('Activity Feed');
    const feed = [];
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, data: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    // Get last 50 entries, most recent first
    for (let r = Math.max(1, data.length - 50); r < data.length; r++) {
      const row = data[r];
      feed.unshift({
        timestamp: row[0] || '',
        poNumber: row[1] || '',
        action: row[2] || '',
        description: row[3] || '',
        oldValue: row[4] || '',
        newValue: row[5] || '',
        user: row[6] || ''
      });
    }
    
    return { success: true, data: feed };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function apiGetAssumptions() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('OpenPOs Assumptions');
    const assumptions = [];
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, data: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      assumptions.push({
        name: row[0] || '',
        value: Number(row[1]) || 0,
        description: row[2] || ''
      });
    }
    
    return { success: true, data: assumptions };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function apiUpdateAssumption(name, value) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('OpenPOs Assumptions');
    
    if (!sheet || sheet.getLastRow() < 2) {
      throw new Error('Assumptions sheet not found');
    }
    
    const data = sheet.getDataRange().getValues();
    let updated = false;
    
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][0]).trim() === name) {
        sheet.getRange(r + 1, 2).setValue(Number(value));
        updated = true;
        break;
      }
    }
    
    if (!updated) {
      throw new Error(`Assumption "${name}" not found`);
    }
    
    // Log the change
    logToFeed_('SYSTEM', 'Assumption Updated', `${name} changed to ${value} days`, '', '');
    
    return { success: true, name: name, value: value };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

////////////////////////////// DIAGNOSTICS //////////////////////////////
function diagnosePullIssues() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Test 1: Credentials
    ui.alert('Step 1: Testing credentials...');
    const creds = getCreds_();
    ui.alert('✅ Credentials OK', 'API credentials are set correctly.', ui.ButtonSet.OK);
    
    // Test 2: API Connection
    ui.alert('Step 2: Testing API connection...');
    const testData = debugAPICall();
    
    if (testData.length === 0) {
      ui.alert('⚠️ No Data', 'API connected but returned no data. Possible issues:\n\n• No POs with statuses: ' + OPEN_STATUSES.join(', ') + '\n• API permissions\n• Account has no open POs', ui.ButtonSet.OK);
    } else {
      ui.alert('✅ API Working', `Successfully retrieved ${testData.length} POs from API.`, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    ui.alert('❌ Diagnosis Failed', `Error: ${error.message}\n\nCheck:\n1. API credentials\n2. Internet connection\n3. Unleashed API access`, ui.ButtonSet.OK);
  }
}

function debugAPICall() {
  try {
    const { id, key } = getCreds_();
    console.log('API ID exists:', !!id);
    
    // Test with a simple query
    const testParams = { 
      orderStatus: 'Placed', 
      pageSize: 10  // Small test
    };
    
    console.log('Testing API call with params:', testParams);
    const results = fetchPurchaseOrders_(testParams);
    console.log('API returned:', results.length, 'items');
    console.log('First item sample:', results.length > 0 ? results[0] : 'No results');
    
    return results;
  } catch (error) {
    console.error('Debug API call failed:', error);
    throw error;
  }
}

////////////////////////////// NEW: DELETE ALL CALENDAR EVENTS //////////////////////////////
function deleteAllCalendarEvents() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '🗑️ Delete All Calendar Events',
    '⚠️ WARNING: This will delete ALL events from "' + CALENDAR_NAME + '"\nThis cannot be undone!',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ui.alert('Deletion cancelled.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const cal = getTargetCalendar_();
    const now = new Date();
    const oneYearAgo = new Date(now.getFullYear() - 1, now.getMonth(), now.getDate());
    const oneYearFuture = new Date(now.getFullYear() + 1, now.getMonth(), now.getDate());
    
    const events = cal.getEvents(oneYearAgo, oneYearFuture);
    let deletedCount = 0;
    
    for (let i = 0; i < events.length; i++) {
      try {
        events[i].deleteEvent();
        deletedCount++;
        
        // Add small delay to avoid rate limits
        if (i % 10 === 0) {
          Utilities.sleep(100);
        }
      } catch (error) {
        console.error(`Failed to delete event ${i + 1}:`, error);
      }
    }
    
    // Clear event IDs from sheet
    clearAllEventIdColumns();
    
    ui.alert('✅ Deletion Complete', `Successfully deleted ${deletedCount} calendar events.`, ui.ButtonSet.OK);
    
    logToFeed_('SYSTEM', 'Calendar Events Deleted', `Deleted ${deletedCount} events from calendar`, '', '');
    
  } catch (error) {
    ui.alert('❌ Deletion Failed', `Error: ${error.message}`, ui.ButtonSet.OK);
  }
}

function debugPODates() {
  try {
    const { id, key } = getCreds_();
    console.log('Testing date extraction from POs...');
    
    const testParams = { 
      orderStatus: 'Placed', 
      pageSize: 5
    };
    
    const results = fetchPurchaseOrders_(testParams);
    console.log(`Found ${results.length} POs for date debugging`);
    
    for (const po of results) {
      console.log(`\n=== PO ${po.OrderNumber} ===`);
      console.log('OrderDate:', po.OrderDate);
      console.log('DeliveryDate:', po.DeliveryDate);
      console.log('LastModifiedOn:', po.LastModifiedOn);
      console.log('Comments:', po.Comments);
      console.log('Supplier:', po.Supplier?.SupplierName);
      
      // Check lines
      const lines = po.PurchaseOrderLines || [];
      console.log(`Number of lines: ${lines.length}`);
      
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        console.log(`Line ${i + 1} DeliveryDate:`, line.DeliveryDate);
        console.log(`Line ${i + 1} Notes:`, line.Notes);
        console.log(`Line ${i + 1} Description:`, line.ProductDescription);
      }
      
      const extractedDate = getDepartDate_(po, lines[0]);
      console.log('Extracted departure date:', extractedDate);
    }
    
    return results;
  } catch (error) {
    console.error('Date debugging failed:', error);
    throw error;
  }
}

function testCredentials() {
  try {
    const creds = getCreds_();
    console.log('Credentials are set correctly');
    return true;
  } catch (error) {
    console.error('Credential error:', error.message);
    SpreadsheetApp.getUi().alert('❌ Credentials Error', 'Please set Unleashed API credentials first using the menu.', SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  }
}

////////////////////////////// UNLEASHED API - FIXED VERSION //////////////////////////////
function getCreds_() {
  const props = PropertiesService.getDocumentProperties();
  const id = props.getProperty('unleashed.apiId');
  const key = props.getProperty('unleashed.apiKey');
  
  console.log('Creds check - ID exists:', !!id, 'Key exists:', !!key);
  
  if (!id || !key) {
    throw new Error('Unleashed API credentials not set. Please use "Set Unleashed Credentials" from the menu.');
  }
  
  return { id, key };
}

function setupCredentials() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  
  const currentId = props.getProperty('unleashed.apiId') || '';
  const currentKey = props.getProperty('unleashed.apiKey') || '';
  
  const idResponse = ui.prompt('Unleashed API ID', 'Enter your Unleashed API ID:', ui.ButtonSet.OK_CANCEL);
  if (idResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const keyResponse = ui.prompt('Unleashed API Key', 'Enter your Unleashed API Key:', ui.ButtonSet.OK_CANCEL);
  if (keyResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const newId = idResponse.getResponseText().trim();
  const newKey = keyResponse.getResponseText().trim();
  
  if (!newId || !newKey) {
    ui.alert('Error', 'Both API ID and Key are required.', ui.ButtonSet.OK);
    return;
  }
  
  props.setProperty('unleashed.apiId', newId);
  props.setProperty('unleashed.apiKey', newKey);
  
  ui.alert('✅ Success', 'Unleashed API credentials saved successfully!', ui.ButtonSet.OK);
}

// FIXED: Use Base64 encoding for signature (not hex)
function signQueryString_(queryString, apiKey) {
  // Remove leading ? if present for signature calculation
  const stringToSign = queryString.startsWith('?') ? queryString.substring(1) : queryString;
  console.log('Signing query string:', stringToSign);
  
  const hmac = Utilities.computeHmacSha256Signature(stringToSign, apiKey);
  const signature = Utilities.base64Encode(hmac);
  console.log('Signature generated (first 10 chars):', signature.substring(0, 10));
  
  return signature;
}

function buildSortedQueryString_(params) {
  const sortedKeys = Object.keys(params).sort();
  const pairs = sortedKeys.map(key => `${encodeURIComponent(key)}=${encodeURIComponent(params[key])}`);
  return pairs.length > 0 ? `?${pairs.join('&')}` : '';
}

function unleashedHeaders_(apiId, signature) {
  return {
    'Accept': 'application/json',
    'api-auth-id': apiId,
    'api-auth-signature': signature,
    'Content-Type': 'application/json',
    'User-Agent': CLIENT_TYPE
  };
}

// FIXED: Proper pagination with page number in URL path + enhanced debugging
function fetchPurchaseOrders_(params) {
  const { id, key } = getCreds_();
  const queryParams = { ...params };
  delete queryParams.page; // Remove page from query params since it goes in URL path
  
  console.log('API Parms:', queryParams);
  
  let allResults = [];
  let page = 1;
  let hasMorePages = true;
  
  while (hasMorePages) {
    // Build query string WITHOUT page parameter
    const queryString = buildSortedQueryString_(queryParams);
    const signature = signQueryString_(queryString, key);
    
    // CRITICAL FIX: Page number goes in URL path, not query string
    const url = `${BASE_URL}/PurchaseOrders/${page}${queryString}`;
    console.log(`Fetching from URL: ${url}`);
    
    const options = {
      method: 'GET',
      headers: unleashedHeaders_(id, signature),
      muteHttpExceptions: true
    };
    
    try {
      const response = UrlFetchApp.fetch(url, options);
      const statusCode = response.getResponseCode();
      const content = response.getContentText();
      
      console.log(`Response Status: ${statusCode}`);
      
      if (statusCode !== 200) {
        console.error(`API Error ${statusCode}: ${content}`);
        
        // Handle rate limiting
        if (statusCode === 429) {
          console.log('Rate limit hit, waiting 2 seconds...');
          Utilities.sleep(2000);
          continue; // Retry the same page
        }
        
        throw new Error(`Unleashed API returned ${statusCode}: ${content.substring(0, 200)}`);
      }
      
      const result = JSON.parse(content);
      console.log(`Page ${page}: ${result.Items ? result.Items.length : 0} items`);
      
      if (result.Items && result.Items.length > 0) {
        allResults = allResults.concat(result.Items);
        
        // Check if we have more pages
        if (result.Items.length < (queryParams.pageSize || PAGE_SIZE)) {
          hasMorePages = false;
        } else {
          page++;
          // Small delay between pages to avoid rate limiting
          Utilities.sleep(500);
        }
      } else {
        hasMorePages = false;
      }
      
      // Safety limit
      if (page > 50) {
        console.warn('Reached safety limit of 50 pages');
        break;
      }
    } catch (error) {
      console.error(`Error fetching page ${page}:`, error);
      throw error;
    }
  }
  
  console.log(`Total items fetched: ${allResults.length}`);
  return allResults;
}

//////////////////////// NEW: PERMIT NUMBER FUNCTIONALITY ////////////////////////
function assignPermitNumberMenu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  
  const poSheet = ss.getSheetByName(SHEET_NAME);
  if (!poSheet || poSheet.getLastRow() < 2) {
    ui.alert('Error', 'No purchase orders found. Please pull data first.', ui.ButtonSet.OK);
    return;
  }
  
  const poData = poSheet.getRange(2, 1, poSheet.getLastRow() - 1, 1).getValues();
  const poNumbers = [...new Set(poData.map(row => row[0]).filter(po => po))];
  
  if (poNumbers.length === 0) {
    ui.alert('Error', 'No purchase order numbers found.', ui.ButtonSet.OK);
    return;
  }
  
  const poList = poNumbers.map(po => `• ${po}`).join('\n');
  const poResponse = ui.prompt('Assign Permit Number to Purchase Order', 
    `Available POs:\n${poList}\n\nEnter PO Number:`, ui.ButtonSet.OK_CANCEL);
  
  if (poResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedPO = poResponse.getResponseText().trim();
  if (!poNumbers.includes(selectedPO)) {
    ui.alert('Error', 'Invalid PO number selected.', ui.ButtonSet.OK);
    return;
  }
  
  const permitResponse = ui.prompt('Permit Number', 'Enter permit number:', ui.ButtonSet.OK_CANCEL);
  if (permitResponse.getSelectedButton() !== ui.Button.OK) return;
  const permitNumber = permitResponse.getResponseText().trim();
  
  if (!permitNumber) {
    ui.alert('Error', 'Permit number cannot be empty.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const result = assignPermitNumber_(selectedPO, permitNumber);
    ui.alert('✅ Permit Number Assigned', 
      `Successfully assigned permit number "${permitNumber}" to PO ${selectedPO}.`, 
      ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error', `Failed to assign permit number: ${error.message}`, ui.ButtonSet.OK);
  }
}

function assignPermitNumber_(poNumber, permitNumber) {
  const ss = SpreadsheetApp.getActive();
  let sheet = getOrCreateSheet_(ss, PERMIT_NUMBERS_SHEET);
  
  if (sheet.getLastRow() === 0) {
    const headers = ['PO Number', 'Permit Number', 'Assigned By', 'Timestamp'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e0f2f1');
    sheet.autoResizeColumns(1, headers.length);
  }
  
  const data = sheet.getDataRange().getValues();
  let existingRow = -1;
  
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][0]) === poNumber) {
      existingRow = r;
      break;
    }
  }
  
  const timestamp = new Date();
  const user = Session.getEffectiveUser().getEmail();
  const newRow = [
    poNumber,
    permitNumber,
    user,
    timestamp
  ];
  
  if (existingRow !== -1) {
    sheet.getRange(existingRow + 1, 1, 1, newRow.length).setValues([newRow]);
  } else {
    sheet.insertRowAfter(1);
    sheet.getRange(2, 1, 1, newRow.length).setValues([newRow]);
  }
  
  sheet.getRange(2, 4).setNumberFormat('yyyy-MM-dd HH:mm:ss');
  
  // Update the main sheet with permit number column
  updatePermitNumberColumn_(ss);
  
  logToFeed_(poNumber, 'Permit Number Assigned', `Permit number assigned: ${permitNumber}`, '', permitNumber);
  
  return { poNumber, permitNumber, user, timestamp };
}

function loadPermitNumbers_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(PERMIT_NUMBERS_SHEET);
  const permitNumbers = {};
  
  if (!sheet || sheet.getLastRow() < 2) return permitNumbers;
  
  const data = sheet.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const po = String(row[0] || '').trim();
    const permit = String(row[1] || '').trim();
    
    if (po && permit) {
      permitNumbers[po] = permit;
    }
  }
  
  return permitNumbers;
}

function updatePermitNumberColumn_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return;
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Check if Permit Number column exists, if not add it
  let permitCol = headers.indexOf('Permit Number');
  if (permitCol === -1) {
    // Add new column at the end
    permitCol = headers.length;
    sheet.getRange(1, permitCol + 1).setValue('Permit Number').setFontWeight('bold').setBackground('#e0f2f1');
    
    // Update all existing rows with empty permit numbers
    const permitNumbers = loadPermitNumbers_();
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      const poNumber = String(row[0] || '').trim();
      
      if (poNumber && permitNumbers[poNumber]) {
        sheet.getRange(r + 1, permitCol + 1).setValue(permitNumbers[poNumber]);
      } else {
        sheet.getRange(r + 1, permitCol + 1).setValue('');
      }
    }
  }
  
  sheet.autoResizeColumn(permitCol + 1);
}

//////////////////////// FIXED: PROPER WEEKEND ADJUSTMENT ////////////////////////
function addCalendarDaysWithWeekendAdjust_(iso, days) {
  if (!iso) return '';
  
  let d = new Date(iso + 'T00:00:00');
  if (isNaN(d)) return '';
  
  // Add raw calendar days
  let current = new Date(d);
  current.setDate(current.getDate() + Number(days));
  
  // Apply weekend adjustment to the FINAL date
  const finalDay = current.getDay();
  
  if (finalDay === 6) { // Saturday
    current.setDate(current.getDate() + 2);
  } else if (finalDay === 0) { // Sunday  
    current.setDate(current.getDate() + 1);
  }
  
  return Utilities.formatDate(current, DEFAULT_TIMEZONE, 'yyyy-MM-dd');
}

//////////////////////// ENHANCED FORMATTING - LABELS BLACK WHEN COMPLETED ////////////////////////
function applyCompleteFormatting_(sheet, dataRows) {
  if (dataRows < 1) return;
  
  sheet.getRange(2, 1, dataRows, sheet.getLastColumn()).setHorizontalAlignment('center');
  
  // ENSURE ALL MAIN DATE COLUMNS (L-Q) USE ISO FORMAT
  const dateColumns = [12, 13, 14, 15, 16, 17]; // Columns L-Q
  for (let col of dateColumns) {
    sheet.getRange(2, col, dataRows, 1).setNumberFormat('yyyy-MM-dd');
  }
  
  sheet.getRange(2, 10, dataRows, 8).setNumberFormat('yyyy-MM-dd');
  sheet.getRange(2, 7, dataRows, 1).setNumberFormat('0.00');
  sheet.getRange(2, 8, dataRows, 2).setNumberFormat('$#,##0.00');
  
  const phaseColors = [
    PHASE_SHEET_HEX['Depart Canada'],
    PHASE_SHEET_HEX['Arrive in Aus'],
    PHASE_SHEET_HEX['Clears Customs'],
    PHASE_SHEET_HEX['Irradiation'],
    PHASE_SHEET_HEX['Testing'],
    PHASE_SHEET_HEX['Packaging and Release']
  ];
  sheet.getRange(2, 12, dataRows, 6).setBackgrounds(Array(dataRows).fill(phaseColors));
  
  // NEW: Enhanced label formatting with completion priority
  const labelRange = sheet.getRange(2, 25, dataRows, 6); // Columns Y-AD
  const labelBackgrounds = labelRange.getBackgrounds();
  const labelTextColors = labelRange.getFontColors();
  const dateValues = sheet.getRange(2, 12, dataRows, 6).getValues(); // Main date columns L-Q
  
  const completedEvents = loadCompletedEvents_();
  
  const now = new Date();
  const nowBrisbane = Utilities.formatDate(now, DEFAULT_TIMEZONE, 'yyyy-MM-dd');
  const today = new Date(nowBrisbane + 'T00:00:00');
  
  const todayDay = today.getDay();
  const thisWeekStart = new Date(today);
  thisWeekStart.setDate(today.getDate() - (todayDay === 0 ? 6 : todayDay - 1));
  const thisWeekEnd = new Date(thisWeekStart);
  thisWeekEnd.setDate(thisWeekStart.getDate() + 6);
  const nextWeekStart = new Date(thisWeekEnd);
  nextWeekStart.setDate(thisWeekEnd.getDate() + 1);
  const nextWeekEnd = new Date(thisWeekStart);
  nextWeekEnd.setDate(thisWeekStart.getDate() + 6);

  for (let r = 0; r < dataRows; r++) {
    const rowNumber = r + 2; // Actual row number in sheet
    const poNumber = String(sheet.getRange(rowNumber, 1).getValue() || '').trim();
    
    for (let c = 0; c < 6; c++) {
      const milestoneCols = ['Depart Canada', 'Arrive in Aus', 'Clears Customs', 'Irradiation', 'Testing', 'Packaging and Release'];
      const milestone = milestoneCols[c];
      
      // Check if this milestone is completed
      const isCompleted = poNumber && completedEvents[poNumber] && completedEvents[poNumber].has(milestone);
      
      if (isCompleted) {
        // NEW: Completed milestones get black background and empty value in labels
        labelBackgrounds[r][c] = COLOR_COMPLETED;
        labelTextColors[r][c] = COLOR_COMPLETED_TEXT;
        sheet.getRange(rowNumber, 25 + c).setValue(''); // Clear the label value
      } else {
        const dateValue = dateValues[r][c];
        if (dateValue instanceof Date && !isNaN(dateValue)) {
          const date = new Date(dateValue);
          date.setHours(0, 0, 0, 0);
          
          if (date >= thisWeekStart && date <= thisWeekEnd) {
            labelBackgrounds[r][c] = LABEL_THIS_WEEK_HEX;
          } else if (date >= nextWeekStart && date <= nextWeekEnd) {
            labelBackgrounds[r][c] = LABEL_NEXT_WEEK_HEX;
          } else if (date < thisWeekStart) {
            labelBackgrounds[r][c] = LABEL_LAST_WEEK_HEX;
          } else {
            labelBackgrounds[r][c] = phaseColors[c];
          }
        } else {
          labelBackgrounds[r][c] = phaseColors[c];
        }
        labelTextColors[r][c] = '#000000'; // Default black text
      }
    }
  }
  
  labelRange.setBackgrounds(labelBackgrounds);
  labelRange.setFontColors(labelTextColors);
  
  const orderNumRange = sheet.getRange(2, 1, dataRows, 1);
  const irradRequiredValues = sheet.getRange(2, 32, dataRows, 1).getValues().flat();
  const orderBgColors = orderNumRange.getBackgrounds();
  
  for (let i = 0; i < dataRows; i++) {
    if (irradRequiredValues[i] === true || String(irradRequiredValues[i]).toUpperCase() === 'TRUE') {
      orderBgColors[i][0] = COLOR_IRRADIATION_REQUIRED;
    }
  }
  orderNumRange.setBackgrounds(orderBgColors);
  
  const statusRange = sheet.getRange(2, 2, dataRows, 1);
  const statusValues = statusRange.getValues().flat();
  const statusBgColors = statusRange.getBackgrounds();
  
  for (let i = 0; i < dataRows; i++) {
    if (String(statusValues[i]).trim().toUpperCase() === 'PARKED') {
      statusBgColors[i][0] = COLOR_PARKED_STATUS;
    }
  }
  statusRange.setBackgrounds(statusBgColors);
  
  // Update permit number column
  updatePermitNumberColumn_(SpreadsheetApp.getActive());
  
  sheet.autoResizeColumns(1, sheet.getLastColumn());
}

//////////////////////// ENHANCED CHANGE TRACKING - ONLY RELEASE DATE CHANGES ////////////////////////
function trackDateChanges_(oldData, newData, headers) {
  const orderCol = headers.indexOf('OrderNumber');
  const releaseDateCol = headers.indexOf('Packaging and Release'); // Only track release date changes
  
  const statusCol = headers.indexOf('OrderStatus');
  
  for (let r = 1; r < Math.min(oldData.length, newData.length); r++) {
    const oldRow = oldData[r];
    const newRow = newData[r];
    const poNumber = String(newRow[orderCol] || '').trim();
    
    if (!poNumber) continue;
    
    // Track release date changes only
    if (releaseDateCol !== -1) {
      const oldDate = toISODateCell_(oldRow[releaseDateCol]);
      const newDate = toISODateCell_(newRow[releaseDateCol]);
      
      if (oldDate && newDate && oldDate !== newDate) {
        const oldDateObj = parseISO_(oldDate);
        const newDateObj = parseISO_(newDate);
        
        if (oldDateObj && newDateObj) {
          const dayDiff = Math.round((newDateObj - oldDateObj) / (1000 * 60 * 60 * 24));
          
          // Log all release date changes regardless of magnitude
          const direction = dayDiff > 0 ? 'delayed' : 'accelerated';
          logToFeed_(
            poNumber,
            'Release Date Change',
            `Packaging and Release date ${direction} by ${Math.abs(dayDiff)} days`,
            oldDate,
            newDate
          );
        }
      }
    }
    
    // Track status changes (Parked → Placed only)
    const oldStatus = String(oldRow[statusCol] || '').trim();
    const newStatus = String(newRow[statusCol] || '').trim();
    
    if (oldStatus && newStatus && oldStatus !== newStatus) {
      if (oldStatus === 'Parked' && newStatus === 'Placed') {
        logToFeed_(
          poNumber,
          'Status Change',
          `Purchase Order status changed from Parked to Placed`,
          oldStatus,
          newStatus
        );
      }
    }
  }
}

//////////////////////// ENHANCED COMPLETE EVENT - UPDATE LABELS TO BLACK ////////////////////////
function markEventComplete_(poNumber, milestone, completedBy, comments) {
  const ss = SpreadsheetApp.getActive();
  const poSheet = ss.getSheetByName(SHEET_NAME);
  const data = poSheet.getDataRange().getValues();
  const headers = data[0];
  
  const col = {
    order: headers.indexOf('OrderNumber'),
    depart: headers.indexOf('Depart Canada'),
    arrive: headers.indexOf('Arrive in Aus'),
    customs: headers.indexOf('Clears Customs'),
    irrad: headers.indexOf('Irradiation'),
    testing: headers.indexOf('Testing'),
    release: headers.indexOf('Packaging and Release'),
    departId: headers.indexOf('DepartEventId'),
    arriveId: headers.indexOf('ArriveEventId'),
    customsId: headers.indexOf('CustomsEventId'),
    irradId: headers.indexOf('IrradiationEventId'),
    testId: headers.indexOf('TestingEventId'),
    releaseId: headers.indexOf('PackRelEventId'),
    comments: headers.indexOf('Comments')
  };
  
  const milestoneCols = {
    'Depart Canada': { date: col.depart, eventId: col.departId, label: 25 }, // Column Y
    'Arrive in Aus': { date: col.arrive, eventId: col.arriveId, label: 26 }, // Column Z
    'Clears Customs': { date: col.customs, eventId: col.customsId, label: 27 }, // Column AA
    'Irradiation': { date: col.irrad, eventId: col.irradId, label: 28 }, // Column AB
    'Testing': { date: col.testing, eventId: col.testId, label: 29 }, // Column AC
    'Packaging and Release': { date: col.release, eventId: col.releaseId, label: 30 }, // Column AD
    'BLS Delivery': { date: col.depart, eventId: col.departId, label: 25 } // BLS uses depart column and label Y
  };
  
  const milestoneInfo = milestoneCols[milestone];
  if (!milestoneInfo) {
    throw new Error(`Unknown milestone: ${milestone}`);
  }
  
  let updated = false;
  let eventId = '';
  const timestamp = new Date();
  const formattedTimestamp = Utilities.formatDate(timestamp, DEFAULT_TIMEZONE, 'yyyy-MM-dd HH:mm');
  
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (String(row[col.order]) === poNumber) {
      // Mark date as completed (black background, white text)
      const dateCell = poSheet.getRange(r + 1, milestoneInfo.date + 1);
      dateCell.setBackground(COLOR_COMPLETED)
              .setFontColor(COLOR_COMPLETED_TEXT);
      
      // NEW: Also mark corresponding label as black and empty
      const labelCell = poSheet.getRange(r + 1, milestoneInfo.label);
      labelCell.setBackground(COLOR_COMPLETED)
               .setFontColor(COLOR_COMPLETED_TEXT)
               .setValue('');
      
      if (col.comments !== -1) {
        const currentComment = String(row[col.comments] || '').trim();
        const completionNote = `${formattedTimestamp}: ${milestone} completed by ${completedBy}${comments ? ` - ${comments}` : ''}`;
        let newComment = '';
        
        if (currentComment) {
          newComment = `${currentComment}\n${completionNote}`;
        } else {
          newComment = completionNote;
        }
        
        poSheet.getRange(r + 1, col.comments + 1).setValue(newComment);
      }
      
      if (!eventId) {
        eventId = String(row[milestoneInfo.eventId] || '');
      }
      
      updated = true;
    }
  }
  
  if (!updated) {
    throw new Error(`No rows found for PO ${poNumber}`);
  }
  
  let eventDeleted = false;
  if (eventId) {
    try {
      const cal = getTargetCalendar_();
      const event = cal.getEventById(eventId);
      if (event) {
        event.deleteEvent();
        eventDeleted = true;
      }
    } catch (e) {
      console.log(`Event ${eventId} not found: ${e.message}`);
    }
  }
  
  storeCompletedEvent_(poNumber, milestone, completedBy, comments);
  logToFeed_(poNumber, 'Milestone Completed', 
    `${milestone} completed by ${completedBy}`, 
    '', 
    comments || 'No comments'
  );
  
  return { poNumber, milestone, completedBy, comments, eventDeleted };
}

//////////////////////// ENHANCED UNCOMPLETE EVENT - RESTORE LABELS ////////////////////////
function uncompleteEvent_(poNumber, milestone, reason) {
  const ss = SpreadsheetApp.getActive();
  const poSheet = ss.getSheetByName(SHEET_NAME);
  const data = poSheet.getDataRange().getValues();
  const headers = data[0];
  
  const col = {
    order: headers.indexOf('OrderNumber'),
    depart: headers.indexOf('Depart Canada'),
    arrive: headers.indexOf('Arrive in Aus'),
    customs: headers.indexOf('Clears Customs'),
    irrad: headers.indexOf('Irradiation'),
    testing: headers.indexOf('Testing'),
    release: headers.indexOf('Packaging and Release'),
    comments: headers.indexOf('Comments')
  };
  
  const milestoneCols = {
    'Depart Canada': { date: col.depart, label: 25 },
    'Arrive in Aus': { date: col.arrive, label: 26 },
    'Clears Customs': { date: col.customs, label: 27 },
    'Irradiation': { date: col.irrad, label: 28 },
    'Testing': { date: col.testing, label: 29 },
    'Packaging and Release': { date: col.release, label: 30 },
    'BLS Delivery': { date: col.depart, label: 25 }
  };
  
  const milestoneInfo = milestoneCols[milestone];
  if (!milestoneInfo) {
    throw new Error(`Unknown milestone: ${milestone}`);
  }
  
  let updated = false;
  const timestamp = new Date();
  const formattedTimestamp = Utilities.formatDate(timestamp, DEFAULT_TIMEZONE, 'yyyy-MM-dd HH:mm');
  
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (String(row[col.order]) === poNumber) {
      const dateCell = poSheet.getRange(r + 1, milestoneInfo.date + 1);
      const phaseColors = {
        'Depart Canada': PHASE_SHEET_HEX['Depart Canada'],
        'Arrive in Aus': PHASE_SHEET_HEX['Arrive in Aus'],
        'Clears Customs': PHASE_SHEET_HEX['Clears Customs'],
        'Irradiation': PHASE_SHEET_HEX['Irradiation'],
        'Testing': PHASE_SHEET_HEX['Testing'],
        'Packaging and Release': PHASE_SHEET_HEX['Packaging and Release'],
        'BLS Delivery': PHASE_SHEET_HEX['BLS Delivery']
      };
      
      dateCell.setBackground(phaseColors[milestone] || '#FFFFFF')
              .setFontColor('#000000')
              .setNote('');
      
      // NEW: Also restore the label formatting and value
      const labelCell = poSheet.getRange(r + 1, milestoneInfo.label);
      const dateValue = dateCell.getValue();
      if (dateValue) {
        const friendlyLabel = friendlyDateLabel_(toISODateCell_(dateValue));
        labelCell.setValue(friendlyLabel);
      }
      // Background and text color will be updated in the next refresh
      
      if (col.comments !== -1) {
        const currentComment = String(row[col.comments] || '').trim();
        const uncompletionNote = `${formattedTimestamp}: ${milestone} uncompleted${reason ? ` - ${reason}` : ''}`;
        let newComment = '';
        
        if (currentComment) {
          newComment = `${currentComment}\n${uncompletionNote}`;
        } else {
          newComment = uncompletionNote;
        }
        
        poSheet.getRange(r + 1, col.comments + 1).setValue(newComment);
      }
      
      updated = true;
    }
  }
  
  if (!updated) {
    throw new Error(`No rows found for PO ${poNumber}`);
  }
  
  removeCompletedEvent_(poNumber, milestone);
  logToFeed_(poNumber, 'Event Uncompleted', `${milestone} marked as incomplete`, '', reason || 'No reason provided');
  
  return { poNumber, milestone, reason, timestamp: formattedTimestamp };
}

//////////////////////// FIXED DATE EXTRACTION - USING DELIVERY DATE ////////////////////////
function getDepartDate_(po, line) {
  const isoHeader = toISODateCell_(po.DeliveryDate);
  if (isoHeader) return isoHeader;
  const ln = line || {};
  const isoLine = toISODateCell_(ln.DeliveryDate);
  return isoLine || '';
}

//////////////////////// FIXED DATE PARSING FUNCTION ////////////////////////
function toISODateCell_(cellValue) {
  if (!cellValue) return '';
  
  if (typeof cellValue === 'string') {
    // Handle Unleashed /Date(timestamp)/ format
    const dateMatch = cellValue.match(/\/Date\((\-?\d+)\)\//);
    if (dateMatch) {
      const timestamp = Number(dateMatch[1]);
      const d = new Date(timestamp);
      if (!isNaN(d)) {
        // Convert to Brisbane timezone and return YYYY-MM-DD
        return Utilities.formatDate(d, DEFAULT_TIMEZONE, 'yyyy-MM-dd');
      }
    }
    
    // Handle simple YYYY-MM-DD strings
    if (cellValue.match(/^\d{4}-\d{2}-\d{2}$/)) return cellValue;
    
    // Handle ISO strings with time components
    const isoMatch = cellValue.match(/^(\d{4}-\d{2}-\d{2})T/);
    if (isoMatch) return isoMatch[1];
    
    // Fallback: try direct parsing
    const d = new Date(cellValue);
    return isNaN(d) ? '' : Utilities.formatDate(d, DEFAULT_TIMEZONE, 'yyyy-MM-dd');
  }
  
  if (cellValue instanceof Date && !isNaN(cellValue)) {
    return Utilities.formatDate(cellValue, DEFAULT_TIMEZONE, 'yyyy-MM-dd');
  }
  
  return '';
}

//////////////////////// ENHANCED MAIN FUNCTION WITH DEBUGGING ////////////////////////
function pullOpenPurchaseOrders() {
  console.log('=== STARTING PULL OPEN PURCHASE ORDERS ===');
  
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = getOrCreateSheet_(ss, SHEET_NAME);
    console.log('Sheet ready:', SHEET_NAME);

    // Test credentials first
    const creds = getCreds_();
    console.log('Credentials verified');

    let oldData = [];
    if (sheet.getLastRow() > 1) {
      oldData = sheet.getDataRange().getValues();
      console.log('Existing data rows:', oldData.length - 1);
    }

    sheet.clear();
    
    const headers = [
      'OrderNumber','OrderStatus','SupplierName','WarehouseCode',
      'ProductCode','ProductDescription','OrderedQuantity','UnitPrice','Total',
      'LastModifiedOn','Comments',
      'Depart Canada','Arrive in Aus','Clears Customs','Irradiation','Testing','Packaging and Release',
      'DepartEventId','ArriveEventId','CustomsEventId','IrradiationEventId','TestingEventId','PackRelEventId','SyncStatus',
      'Canada Export','Arrives Aus','Customs Release','Irradiation Complete','Testing Complete','Packaged and Released',
      'Key','Irradiation Required', 'Permit Number'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e0f2f1');

    const assumptions  = readAssumptions_(ss);
    ensureRequirementsSheet_(ss);

    let rows = [];
    
    console.log('Starting to fetch POs for statuses:', OPEN_STATUSES);
    
    for (const status of OPEN_STATUSES) {
      console.log(`=== Fetching POs with status: ${status} ===`);
      const pos = fetchPurchaseOrders_({ orderStatus: status, pageSize: PAGE_SIZE });
      console.log(`Found ${pos.length} POs with status ${status}`);
      
      for (const po of pos) {
        const lines = po.PurchaseOrderLines || [];
        if (lines.length === 0) {
          console.log(`PO ${po.OrderNumber}: No lines, creating empty row`);
          rows.push(createRow_(po, {}));
        } else {
          console.log(`PO ${po.OrderNumber}: ${lines.length} lines`);
          for (const line of lines) {
            rows.push(createRow_(po, line));
          }
        }
      }
    }

    console.log(`Total rows before filtering: ${rows.length}`);
    
    // Apply warehouse filtering
    rows = rows.filter(r => {
      const warehouse = String(r[3] || '').toUpperCase();
      const shouldExclude = warehouse === EXCLUDED_WAREHOUSE.toUpperCase();
      if (shouldExclude) {
        console.log(`Excluding row with warehouse: ${warehouse}`);
      }
      return !shouldExclude;
    });
    
    console.log(`Total rows after filtering: ${rows.length}`);

    if (rows.length === 0) {
      const ui = SpreadsheetApp.getUi();
      ui.alert('❌ No Data Found', 
        'No purchase orders found. Check:\n\n• API credentials are set\n• Order statuses exist in Unleashed\n• There are actual open POs\n• Check console for detailed logs', 
        ui.ButtonSet.OK);
      return;
    }

    const requirements = readRequirements_(ss);
    const fixedDates = loadFixedDates_();
    
    const blsRows = [];
    const otherRows = [];
    
    rows.forEach(r => {
      if (String(r[2]).trim() === 'BLS Wholesalers Pty Ltd') {
        blsRows.push(r);
      } else {
        otherRows.push(r);
      }
    });
    
    console.log(`Processing: ${otherRows.length} regular rows, ${blsRows.length} BLS rows`);
    
    const processedOtherRows = otherRows.map(r => {
      const orderNumber = r[0];
      const departISO = toISODateCell_(r[11]);
      
      console.log(`Processing PO ${orderNumber}: departISO = ${departISO}`);
      
      const fixedDateData = fixedDates[orderNumber] || {};
      let m;
      
      // FIXED: Proper fixed date dependency - use fixed date as starting point
      if (fixedDateData['Depart Canada']) {
        m = computeMilestones_(orderNumber, fixedDateData['Depart Canada'], assumptions, requirements);
      } else {
        m = computeMilestones_(orderNumber, departISO, assumptions, requirements);
      }
      
      // Apply any other fixed dates
      if (fixedDateData['Arrive in Aus']) m.arrive = fixedDateData['Arrive in Aus'];
      if (fixedDateData['Clears Customs']) m.customs = fixedDateData['Clears Customs'];
      if (fixedDateData['Irradiation']) m.irradiation = fixedDateData['Irradiation'];
      if (fixedDateData['Testing']) m.testing = fixedDateData['Testing'];
      if (fixedDateData['Packaging and Release']) m.packRel = fixedDateData['Packaging and Release'];
      
      r[11] = m.depart || '';
      r[12] = m.arrive || '';
      r[13] = m.customs || '';
      r[14] = m.irradiation || '';
      r[15] = m.testing || '';
      r[16] = m.packRel || '';
      
      r[24] = friendlyDateLabel_(r[11]) || '';
      r[25] = friendlyDateLabel_(r[12]) || '';
      r[26] = friendlyDateLabel_(r[13]) || '';
      r[27] = friendlyDateLabel_(r[14]) || '';
      r[28] = friendlyDateLabel_(r[15]) || '';
      r[29] = friendlyDateLabel_(r[16]) || '';
      
      const req = requirements[(orderNumber || '').toString().trim().toLowerCase()] || { irradiation:false };
      r[31] = req.irradiation ? 'TRUE' : 'FALSE';
      
      r[30] = makeKey_(r[0], r[4], r[3]);
      
      // Add empty permit number column to match headers
      r[32] = '';
      
      return r;
    });

    const processedBLSRows = blsRows.map(r => {
      const orderNumber = r[0];
      
      const fixedDateData = fixedDates[orderNumber] || {};
      
      let departISO = toISODateCell_(r[11]);
      if (fixedDateData['BLS Delivery'] || fixedDateData['Depart Canada']) {
        departISO = fixedDateData['BLS Delivery'] || fixedDateData['Depart Canada'];
      }
      
      r[11] = departISO || '';
      r[12] = '';
      r[13] = '';
      r[14] = '';
      r[15] = '';
      r[16] = '';
      
      r[24] = friendlyDateLabel_(r[11]) || '';
      r[25] = '';
      r[26] = '';
      r[27] = '';
      r[28] = '';
      r[29] = '';
      
      const req = requirements[(orderNumber || '').toString().trim().toLowerCase()] || { irradiation:false };
      r[31] = req.irradiation ? 'TRUE' : 'FALSE';
      
      r[30] = makeKey_(r[0], r[4], r[3]);
      
      // Add empty permit number column to match headers
      r[32] = '';
      
      return r;
    });

    processedOtherRows.sort((a, b) => {
      const dA = parseISO_(toISODateCell_(a[16]));
      const dB = parseISO_(toISODateCell_(b[16]));
      if (!dA && !dB) return 0;
      if (!dA) return 1;
      if (!dB) return -1;
      return dA - dB;
    });

    processedBLSRows.sort((a, b) => {
      const dA = parseISO_(toISODateCell_(a[11]));
      const dB = parseISO_(toISODateCell_(b[11]));
      if (!dA && !dB) return 0;
      if (!dA) return 1;
      if (!dB) return -1;
      return dA - dB;
    });

    const allProcessedRows = processedOtherRows.concat(processedBLSRows);

    if (allProcessedRows.length > 0) {
      console.log(`Writing ${allProcessedRows.length} rows with ${headers.length} columns`);
      console.log('First row sample:', allProcessedRows[0]);
      
      sheet.getRange(2, 1, allProcessedRows.length, headers.length).setValues(allProcessedRows);
      
      applyCompleteFormatting_(sheet, allProcessedRows.length);
      applyPersistentCompletionFormatting_(sheet, allProcessedRows.length);
      applyFixedDateFormatting_(sheet, allProcessedRows.length);
      
      if (oldData.length > 1) {
        trackDateChanges_(oldData, sheet.getDataRange().getValues(), headers);
      }
      
      SpreadsheetApp.getUi().alert(`✅ Pulled ${allProcessedRows.length} open PO lines (${processedOtherRows.length} regular + ${processedBLSRows.length} BLS at bottom)`);
    } else {
      SpreadsheetApp.getUi().alert('ℹ️ No open purchase orders found after filtering.');
    }
  } catch (error) {
    console.error('Error in pullOpenPurchaseOrders:', error);
    SpreadsheetApp.getUi().alert(`❌ Error: ${error.message}`);
  }
}

////////////////////////////// UTILITY FUNCTIONS //////////////////////////////
function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function parseISO_(iso) {
  if (!iso) return null;
  const d = new Date(iso + 'T00:00:00');
  return isNaN(d) ? null : d;
}

function friendlyDateLabel_(iso) {
  if (!iso) return '';
  const d = parseISO_(iso);
  if (!d) return '';
  
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);
  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  
  if (d.getTime() === today.getTime()) return 'Today';
  if (d.getTime() === tomorrow.getTime()) return 'Tomorrow';
  if (d.getTime() === yesterday.getTime()) return 'Yesterday';
  
  const diffDays = Math.round((d - today) / (1000 * 60 * 60 * 24));
  if (diffDays === 0) return 'Today';
  if (diffDays === 1) return 'Tomorrow';
  if (diffDays === -1) return 'Yesterday';
  if (diffDays > 0 && diffDays <= 7) return `In ${diffDays} days`;
  if (diffDays < 0 && diffDays >= -7) return `${Math.abs(diffDays)} days ago`;
  
  return Utilities.formatDate(d, DEFAULT_TIMEZONE, 'E, MMM d');
}

//////////////////////// FIXED: PROPER FIXED DATE DEPENDENCY ////////////////////////
function computeMilestones_(orderNumber, departISO, assumptions, requirements) {
  if (!departISO) {
    return {
      depart: '',
      arrive: '',
      customs: '',
      irradiation: '',
      testing: '',
      packRel: ''
    };
  }

  const req = requirements[(orderNumber || '').toString().trim().toLowerCase()] || { irradiation: false };
  
  // ALWAYS calculate from the given depart date (which may be fixed)
  let arrive = addCalendarDaysWithWeekendAdjust_(departISO, assumptions.shippingDays);
  let customs = addCalendarDaysWithWeekendAdjust_(arrive, assumptions.customsDays);
  
  let irradiation = '';
  if (req.irradiation) {
    irradiation = addCalendarDaysWithWeekendAdjust_(customs, assumptions.irradiationDays);
  }
  
  let testing = addCalendarDaysWithWeekendAdjust_(irradiation || customs, assumptions.testingDays);
  let packRel = addCalendarDaysWithWeekendAdjust_(testing, assumptions.packagingDays);
  
  return {
    depart: departISO,
    arrive: arrive,
    customs: customs,
    irradiation: irradiation,
    testing: testing,
    packRel: packRel
  };
}

function makeKey_(orderNum, productCode, warehouse) {
  return [orderNum, productCode, warehouse].filter(x => x).join('|');
}

function loadCompletedEvents_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(COMPLETED_EVENTS_SHEET);
  const completed = {};
  
  if (!sheet || sheet.getLastRow() < 2) return completed;
  
  const data = sheet.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const po = String(row[0] || '').trim();
    const milestone = String(row[1] || '').trim();
    
    if (po && milestone) {
      if (!completed[po]) completed[po] = new Set();
      completed[po].add(milestone);
    }
  }
  
  return completed;
}

function applyPersistentCompletionFormatting_(sheet, dataRows) {
  if (dataRows < 1) return;
  
  const completedEvents = loadCompletedEvents_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const col = {
    order: headers.indexOf('OrderNumber'),
    depart: headers.indexOf('Depart Canada'),
    arrive: headers.indexOf('Arrive in Aus'),
    customs: headers.indexOf('Clears Customs'),
    irrad: headers.indexOf('Irradiation'),
    testing: headers.indexOf('Testing'),
    release: headers.indexOf('Packaging and Release')
  };
  
  const milestoneMap = {
    'Depart Canada': col.depart,
    'Arrive in Aus': col.arrive,
    'Clears Customs': col.customs,
    'Irradiation': col.irrad,
    'Testing': col.testing,
    'Packaging and Release': col.release,
    'BLS Delivery': col.depart
  };
  
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const poNumber = String(row[col.order] || '').trim();
    
    if (poNumber && completedEvents[poNumber]) {
      for (const milestone of completedEvents[poNumber]) {
        const colIndex = milestoneMap[milestone];
        if (colIndex !== undefined) {
          const cell = sheet.getRange(r + 1, colIndex + 1);
          cell.setBackground(COLOR_COMPLETED)
              .setFontColor(COLOR_COMPLETED_TEXT);
        }
      }
    }
  }
}

function createRow_(po, line) {
  const orderDate = po.OrderDate ? new Date(po.OrderDate) : null;
  const lastModified = po.LastModifiedOn ? new Date(po.LastModifiedOn) : null;
  
  // Extract departure date using DeliveryDate
  const departDate = getDepartDate_(po, line);
  
  // Format dates properly
  const formattedOrderDate = orderDate && !isNaN(orderDate) ? 
    Utilities.formatDate(orderDate, DEFAULT_TIMEZONE, 'yyyy-MM-dd') : '';
  const formattedLastModified = lastModified && !isNaN(lastModified) ? 
    Utilities.formatDate(lastModified, DEFAULT_TIMEZONE, 'yyyy-MM-dd') : '';
  
  console.log(`Creating row for PO ${po.OrderNumber}:`, {
    orderDate: formattedOrderDate,
    lastModified: formattedLastModified,
    departDate: departDate,
    deliveryDate: po.DeliveryDate
  });
  
  return [
    po.OrderNumber || '',
    po.OrderStatus || '',
    po.Supplier?.SupplierName || '',
    po.Warehouse?.WarehouseCode || '',
    line.Product?.ProductCode || '',
    // UPDATED: Product description fix
    (line.Product?.ProductDescription || line.ProductDescription || ''),
    safeNum_(line.OrderQuantity),
    safeNum_(line.UnitPrice),
    safeNum_(line.LineTotal),
    formattedLastModified,
    po.Comments || '',
    departDate, // Use extracted delivery date as departure date
    '', // Arrive in Aus (will be calculated)
    '', // Clears Customs (will be calculated)
    '', // Irradiation (will be calculated)
    '', // Testing (will be calculated)
    '', // Packaging and Release (will be calculated)
    '', // DepartEventId
    '', // ArriveEventId
    '', // CustomsEventId
    '', // IrradiationEventId
    '', // TestingEventId
    '', // PackRelEventId
    '', // SyncStatus
    '', // Canada Export (label)
    '', // Arrives Aus (label)
    '', // Customs Release (label)
    '', // Irradiation Complete (label)
    '', // Testing Complete (label)
    '', // Packaged and Released (label)
    '', // Key
    '', // Irradiation Required
    ''  // Permit Number
  ];
}

function safeNum_(val) {
  if (val === null || val === undefined) return 0;
  const n = Number(val);
  return isNaN(n) ? 0 : n;
}

////////////////////////////// CALENDAR SYNC - FIXED DUPLICATES //////////////////////////////
function syncCalendarEventsFromSheet() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No data found. Please pull purchase orders first.');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const col = {
    order: headers.indexOf('OrderNumber'),
    status: headers.indexOf('OrderStatus'),
    supplier: headers.indexOf('SupplierName'),
    product: headers.indexOf('ProductDescription'),
    quantity: headers.indexOf('OrderedQuantity'),
    depart: headers.indexOf('Depart Canada'),
    arrive: headers.indexOf('Arrive in Aus'),
    customs: headers.indexOf('Clears Customs'),
    irrad: headers.indexOf('Irradiation'),
    testing: headers.indexOf('Testing'),
    release: headers.indexOf('Packaging and Release'),
    departId: headers.indexOf('DepartEventId'),
    arriveId: headers.indexOf('ArriveEventId'),
    customsId: headers.indexOf('CustomsEventId'),
    irradId: headers.indexOf('IrradiationEventId'),
    testId: headers.indexOf('TestingEventId'),
    releaseId: headers.indexOf('PackRelEventId'),
    syncStatus: headers.indexOf('SyncStatus')
  };

  const cal = getTargetCalendar_();
  const completedEvents = loadCompletedEvents_();
  
  let created = 0, updated = 0, errors = 0;
  const statusUpdates = [];

  // First, remove any existing events for these POs to prevent duplicates
  const existingEvents = getExistingEventsForPOs_(cal, data, col);
  
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const orderNumber = String(row[col.order] || '').trim();
    const supplier = String(row[col.supplier] || '').trim();
    const product = String(row[col.product] || '').trim();
    const quantity = safeNum_(row[col.quantity]);
    
    if (!orderNumber) continue;

    const isBLS = supplier === 'BLS Wholesalers Pty Ltd';
    const milestones = [];

    if (isBLS) {
      const departISO = toISODateCell_(row[col.depart]);
      if (departISO) {
        milestones.push({
          name: 'BLS Delivery',
          date: departISO,
          eventId: String(row[col.departId] || ''),
          colIndex: col.departId
        });
      }
    } else {
      const phaseDates = [
        { name: 'Depart Canada', date: row[col.depart], eventId: row[col.departId], colIndex: col.departId },
        { name: 'Arrive in Aus', date: row[col.arrive], eventId: row[col.arriveId], colIndex: col.arriveId },
        { name: 'Clears Customs', date: row[col.customs], eventId: row[col.customsId], colIndex: col.customsId },
        { name: 'Irradiation', date: row[col.irrad], eventId: row[col.irradId], colIndex: col.irradId },
        { name: 'Testing', date: row[col.testing], eventId: row[col.testId], colIndex: col.testId },
        { name: 'Packaging and Release', date: row[col.release], eventId: row[col.releaseId], colIndex: col.releaseId }
      ];

      for (const phase of phaseDates) {
        const iso = toISODateCell_(phase.date);
        if (iso) {
          milestones.push({
            name: phase.name,
            date: iso,
            eventId: String(phase.eventId || ''),
            colIndex: phase.colIndex
          });
        }
      }
    }

    for (const milestone of milestones) {
      const isCompleted = completedEvents[orderNumber] && completedEvents[orderNumber].has(milestone.name);
      if (isCompleted) continue;

      try {
        // Enhanced event title with quantity, product name, and emoji
        const eventTitle = `${quantity} ${product} - ${MILESTONE_EMOJIS[milestone.name]} ${milestone.name} - ${orderNumber}`;
        const eventDesc = `Purchase Order: ${orderNumber}\nSupplier: ${supplier}\nProduct: ${product}\nQuantity: ${quantity}\n${SYNC_TAG}`;
        
        // Check if event already exists to prevent duplicates
        const existingEvent = findExistingEvent_(existingEvents, orderNumber, milestone.name, milestone.date);
        
        const result = upsertAllDayEventColored_(
          cal,
          eventTitle,
          eventDesc,
          milestone.date,
          existingEvent ? existingEvent.getId() : milestone.eventId,
          PHASE_EVENT_COLOR[milestone.name] || CalendarApp.EventColor.YELLOW
        );

        statusUpdates.push({
          row: r + 1,
          col: milestone.colIndex + 1,
          value: result.eventId
        });

        if (result.created) created++;
        else if (result.updated) updated++;
      } catch (e) {
        console.error(`Failed to sync ${milestone.name} for ${orderNumber}:`, e);
        errors++;
      }
    }
  }

  // Batch update event IDs
  for (const update of statusUpdates) {
    sheet.getRange(update.row, update.col).setValue(update.value);
  }

  // Update sync status
  const now = new Date();
  const status = `Last sync: ${Utilities.formatDate(now, DEFAULT_TIMEZONE, 'yyyy-MM-dd HH:mm')} - Created: ${created}, Updated: ${updated}, Errors: ${errors}`;
  
  if (col.syncStatus !== -1) {
    sheet.getRange(2, col.syncStatus + 1, data.length - 1, 1).setValue(status);
  }

  SpreadsheetApp.getUi().alert(`✅ Calendar Sync Complete\nCreated: ${created}, Updated: ${updated}, Errors: ${errors}`);
}

//////////////////////// NEW: DUPLICATE PREVENTION FUNCTIONS ////////////////////////
function getExistingEventsForPOs_(cal, data, col) {
  const now = new Date();
  const oneYearAgo = new Date(now.getFullYear() - 1, now.getMonth(), now.getDate());
  const oneYearFuture = new Date(now.getFullYear() + 1, now.getMonth(), now.getDate());
  
  const events = cal.getEvents(oneYearAgo, oneYearFuture);
  const poEvents = [];
  
  for (const event of events) {
    const desc = event.getDescription() || '';
    if (desc.includes(SYNC_TAG)) {
      poEvents.push(event);
    }
  }
  
  return poEvents;
}

function findExistingEvent_(events, orderNumber, milestone, date) {
  const targetDate = parseISO_(date);
  if (!targetDate) return null;
  
  for (const event of events) {
    const eventDate = event.getStartTime();
    const title = event.getTitle();
    const desc = event.getDescription() || '';
    
    // Check if this event matches our PO and milestone
    if (desc.includes(`Purchase Order: ${orderNumber}`) && 
        title.includes(milestone) &&
        eventDate.getTime() === targetDate.getTime()) {
      return event;
    }
  }
  return null;
}

//////////////////////// NEW: REMOVE CALENDAR DUPLICATES FUNCTION ////////////////////////
function removeCalendarDuplicates() {
  const ui = SpreadsheetApp.getUi();
  const cal = getTargetCalendar_();
  const now = new Date();
  const oneYearAgo = new Date(now.getFullYear() - 1, now.getMonth(), now.getDate());
  const oneYearFuture = new Date(now.getFullYear() + 1, now.getMonth(), now.getDate());
  
  const events = cal.getEvents(oneYearAgo, oneYearFuture);
  const eventMap = new Map();
  let duplicatesRemoved = 0;
  
  for (const event of events) {
    const desc = event.getDescription() || '';
    if (desc.includes(SYNC_TAG)) {
      const title = event.getTitle();
      const date = event.getStartTime();
      const key = `${title}|${date}`;
      
      if (eventMap.has(key)) {
        // This is a duplicate - remove it
        try {
          event.deleteEvent();
          duplicatesRemoved++;
          console.log(`Removed duplicate event: ${title} on ${date}`);
        } catch (e) {
          console.error(`Failed to remove duplicate event: ${title}`, e);
        }
      } else {
        // First time seeing this event, add to map
        eventMap.set(key, event);
      }
    }
  }
  
  ui.alert('✅ Duplicate Cleanup', `Removed ${duplicatesRemoved} duplicate calendar events.`, ui.ButtonSet.OK);
}

function getTargetCalendar_() {
  const cals = CalendarApp.getCalendarsByName(CALENDAR_NAME);
  if (cals.length > 0) return cals[0];
  
  // Create calendar if it doesn't exist
  const newCal = CalendarApp.createCalendar(CALENDAR_NAME, {
    summary: CALENDAR_NAME,
    timeZone: DEFAULT_TIMEZONE
  });
  
  SpreadsheetApp.getUi().alert(`Created new calendar: ${CALENDAR_NAME}`);
  return newCal;
}

//////////////////////// FIXED: +1 DAY CALENDAR EVENT FIX ////////////////////////
function upsertAllDayEventColored_(cal, title, desc, isoDate, existingEventId, color) {
  const date = parseISO_(isoDate);
  if (!date) throw new Error(`Invalid date: ${isoDate}`);

  // FIX: Add +1 day to compensate for calendar display offset
  const correctedDate = new Date(date);
  correctedDate.setDate(correctedDate.getDate() + 1);

  let event = null;
  let created = false;
  let updated = false;

  if (existingEventId) {
    try {
      event = cal.getEventById(existingEventId);
    } catch (e) {
      console.log(`Event ${existingEventId} not found, creating new: ${e.message}`);
    }
  }

  if (event) {
    // Update existing event
    event.setTitle(title);
    event.setDescription(desc);
    event.setAllDayDate(correctedDate); // Use corrected date
    event.setColor(color);
    updated = true;
  } else {
    // Create new event
    event = cal.createAllDayEvent(title, correctedDate, { description: desc }); // Use corrected date
    event.setColor(color);
    created = true;
  }

  return {
    eventId: event.getId(),
    created: created,
    updated: updated
  };
}

function cleanupCalendarEvents() {
  const cal = getTargetCalendar_();
  const now = new Date();
  const oneYearAgo = new Date(now.getFullYear() - 1, now.getMonth(), now.getDate());
  const oneYearFuture = new Date(now.getFullYear() + 1, now.getMonth(), now.getDate());
  
  const events = cal.getEvents(oneYearAgo, oneYearFuture);
  let deleted = 0;
  
  for (const event of events) {
    const desc = event.getDescription() || '';
    if (desc.includes(SYNC_TAG)) {
      const title = event.getTitle();
      const eventDate = event.getStartTime();
      
      // Only delete if it's from the past
      if (eventDate < now) {
        try {
          event.deleteEvent();
          deleted++;
        } catch (e) {
          console.error(`Failed to delete event: ${title}`, e);
        }
      }
    }
  }
  
  SpreadsheetApp.getUi().alert(`✅ Cleaned up ${deleted} past calendar events`);
}

////////////////////////////// ASSUMPTIONS & REQUIREMENTS //////////////////////////////
function normalizeAssumptionsSheet_(ss) {
  const sheet = ss.getSheetByName(ASSUMPTIONS_SHEET);
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  const headers = ['Parameter', 'Value', 'Description'];
  
  if (data.length === 0 || data[0][0] !== 'Parameter') {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
}

function readAssumptions_(ss) {
  const sheet = ss.getSheetByName(ASSUMPTIONS_SHEET);
  const defaults = {
    shippingDays: 3,
    customsDays: 5,
    irradiationDays: 3,
    testingDays: 12,
    packagingDays: 3
  };
  
  if (!sheet || sheet.getLastRow() < 2) return defaults;
  
  const data = sheet.getDataRange().getValues();
  const assumptions = { ...defaults };
  
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const param = String(row[0] || '').trim().toLowerCase();
    const value = safeNum_(row[1]);
    
    switch (param) {
      case 'shipping days':
      case 'shippingdays':
        assumptions.shippingDays = value || defaults.shippingDays;
        break;
      case 'customs days':
      case 'customsdays':
        assumptions.customsDays = value || defaults.customsDays;
        break;
      case 'irradiation days':
      case 'irradiationdays':
        assumptions.irradiationDays = value || defaults.irradiationDays;
        break;
      case 'testing days':
      case 'testingdays':
        assumptions.testingDays = value || defaults.testingDays;
        break;
      case 'packaging days':
      case 'packagingdays':
        assumptions.packagingDays = value || defaults.packagingDays;
        break;
    }
  }
  
  return assumptions;
}

function ensureRequirementsSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, REQUIREMENTS_SHEET);
  if (sheet.getLastRow() === 0) {
    const headers = ['OrderNumber', 'Irradiation', 'Notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e0f2f1');
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

function readRequirements_(ss) {
  const sheet = ss.getSheetByName(REQUIREMENTS_SHEET);
  const requirements = {};
  
  if (!sheet || sheet.getLastRow() < 2) return requirements;
  
  const data = sheet.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const orderNum = String(row[0] || '').trim().toLowerCase();
    const irrad = String(row[1] || '').trim().toLowerCase();
    
    if (orderNum) {
      requirements[orderNum] = {
        irradiation: irrad === 'true' || irrad === 'yes' || irrad === '1'
      };
    }
  }
  
  return requirements;
}

function refreshRequirementsFromOpenPOs_() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  
  const poSheet = ss.getSheetByName(SHEET_NAME);
  if (!poSheet || poSheet.getLastRow() < 2) {
    ui.alert('Error', 'No purchase orders found. Please pull data first.', ui.ButtonSet.OK);
    return;
  }
  
  const reqSheet = ensureRequirementsSheet_(ss);
  const existingReqs = readRequirements_(ss);
  
  const poData = poSheet.getRange(2, 1, poSheet.getLastRow() - 1, 1).getValues();
  const poNumbers = [...new Set(poData.map(row => row[0]).filter(po => po))];
  
  let added = 0;
  for (const po of poNumbers) {
    const poKey = String(po).trim().toLowerCase();
    if (!existingReqs[poKey]) {
      reqSheet.appendRow([po, 'FALSE', '']);
      added++;
    }
  }
  
  if (added > 0) {
    ui.alert('✅ Requirements Updated', `Added ${added} new POs to requirements sheet.`, ui.ButtonSet.OK);
  } else {
    ui.alert('ℹ️ No Changes', 'All POs already exist in requirements sheet.', ui.ButtonSet.OK);
  }
}

////////////////////////////// COMPLETION SYSTEM //////////////////////////////
function completeEventMenu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  
  const poSheet = ss.getSheetByName(SHEET_NAME);
  if (!poSheet || poSheet.getLastRow() < 2) {
    ui.alert('Error', 'No purchase orders found. Please pull data first.', ui.ButtonSet.OK);
    return;
  }
  
  const poData = poSheet.getRange(2, 1, poSheet.getLastRow() - 1, 1).getValues();
  const poNumbers = [...new Set(poData.map(row => row[0]).filter(po => po))];
  
  if (poNumbers.length === 0) {
    ui.alert('Error', 'No purchase order numbers found.', ui.ButtonSet.OK);
    return;
  }
  
  const poList = poNumbers.map(po => `• ${po}`).join('\n');
  const poResponse = ui.prompt('Complete Event', 
    `Available POs:\n${poList}\n\nEnter PO Number:`, ui.ButtonSet.OK_CANCEL);
  
  if (poResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedPO = poResponse.getResponseText().trim();
  if (!poNumbers.includes(selectedPO)) {
    ui.alert('Error', 'Invalid PO number selected.', ui.ButtonSet.OK);
    return;
  }
  
  const milestones = ['Depart Canada', 'Arrive in Aus', 'Clears Customs', 'Irradiation', 'Testing', 'Packaging and Release', 'BLS Delivery'];
  const milestoneResponse = ui.prompt('Milestone', 
    `Available milestones:\n${milestones.map(m => `• ${m}`).join('\n')}\n\nEnter milestone:`, ui.ButtonSet.OK_CANCEL);
  
  if (milestoneResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedMilestone = milestoneResponse.getResponseText().trim();
  if (!milestones.includes(selectedMilestone)) {
    ui.alert('Error', 'Invalid milestone selected.', ui.ButtonSet.OK);
    return;
  }
  
  const commentResponse = ui.prompt('Comments (optional)', 'Enter any comments:', ui.ButtonSet.OK_CANCEL);
  const comments = commentResponse.getSelectedButton() === ui.Button.OK ? commentResponse.getResponseText().trim() : '';
  
  const user = Session.getEffectiveUser().getEmail();
  
  try {
    const result = markEventComplete_(selectedPO, selectedMilestone, user, comments);
    ui.alert('✅ Event Completed', 
      `Successfully marked ${selectedMilestone} as completed for PO ${selectedPO}.`, 
      ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error', `Failed to complete event: ${error.message}`, ui.ButtonSet.OK);
  }
}

function uncompleteEventMenu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  
  const poSheet = ss.getSheetByName(SHEET_NAME);
  if (!poSheet || poSheet.getLastRow() < 2) {
    ui.alert('Error', 'No purchase orders found. Please pull data first.', ui.ButtonSet.OK);
    return;
  }
  
  const completedEvents = loadCompletedEvents_();
  const completablePOs = Object.keys(completedEvents).filter(po => completedEvents[po].size > 0);
  
  if (completablePOs.length === 0) {
    ui.alert('Info', 'No completed events found to uncomplete.', ui.ButtonSet.OK);
    return;
  }
  
  const poList = completablePOs.map(po => `• ${po}`).join('\n');
  const poResponse = ui.prompt('Uncomplete Event', 
    `POs with completed events:\n${poList}\n\nEnter PO Number:`, ui.ButtonSet.OK_CANCEL);
  
  if (poResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedPO = poResponse.getResponseText().trim();
  if (!completablePOs.includes(selectedPO)) {
    ui.alert('Error', 'Invalid PO number or no completed events for this PO.', ui.ButtonSet.OK);
    return;
  }
  
  const availableMilestones = Array.from(completedEvents[selectedPO]);
  const milestoneResponse = ui.prompt('Milestone to Uncomplete', 
    `Completed milestones:\n${availableMilestones.map(m => `• ${m}`).join('\n')}\n\nEnter milestone:`, ui.ButtonSet.OK_CANCEL);
  
  if (milestoneResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedMilestone = milestoneResponse.getResponseText().trim();
  if (!availableMilestones.includes(selectedMilestone)) {
    ui.alert('Error', 'Invalid milestone selected.', ui.ButtonSet.OK);
    return;
  }
  
  const reasonResponse = ui.prompt('Reason (optional)', 'Enter reason for uncompleting:', ui.ButtonSet.OK_CANCEL);
  const reason = reasonResponse.getSelectedButton() === ui.Button.OK ? reasonResponse.getResponseText().trim() : '';
  
  try {
    const result = uncompleteEvent_(selectedPO, selectedMilestone, reason);
    ui.alert('✅ Event Uncompleted', 
      `Successfully marked ${selectedMilestone} as incomplete for PO ${selectedPO}.`, 
      ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error', `Failed to uncomplete event: ${error.message}`, ui.ButtonSet.OK);
  }
}

function addCommentToPOMenu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  
  const poSheet = ss.getSheetByName(SHEET_NAME);
  if (!poSheet || poSheet.getLastRow() < 2) {
    ui.alert('Error', 'No purchase orders found. Please pull data first.', ui.ButtonSet.OK);
    return;
  }
  
  const poData = poSheet.getRange(2, 1, poSheet.getLastRow() - 1, 1).getValues();
  const poNumbers = [...new Set(poData.map(row => row[0]).filter(po => po))];
  
  if (poNumbers.length === 0) {
    ui.alert('Error', 'No purchase order numbers found.', ui.ButtonSet.OK);
    return;
  }
  
  const poList = poNumbers.map(po => `• ${po}`).join('\n');
  const poResponse = ui.prompt('Add Comment to PO', 
    `Available POs:\n${poList}\n\nEnter PO Number:`, ui.ButtonSet.OK_CANCEL);
  
  if (poResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedPO = poResponse.getResponseText().trim();
  if (!poNumbers.includes(selectedPO)) {
    ui.alert('Error', 'Invalid PO number selected.', ui.ButtonSet.OK);
    return;
  }
  
  const commentResponse = ui.prompt('Comment', 'Enter your comment:', ui.ButtonSet.OK_CANCEL);
  if (commentResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const comment = commentResponse.getResponseText().trim();
  if (!comment) {
    ui.alert('Error', 'Comment cannot be empty.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const result = addCommentToPO_(selectedPO, comment);
    ui.alert('✅ Comment Added', `Successfully added comment to PO ${selectedPO}.`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error', `Failed to add comment: ${error.message}`, ui.ButtonSet.OK);
  }
}

function addCommentToPO_(poNumber, comment) {
  const ss = SpreadsheetApp.getActive();
  const poSheet = ss.getSheetByName(SHEET_NAME);
  const data = poSheet.getDataRange().getValues();
  const headers = data[0];
  
  const col = {
    order: headers.indexOf('OrderNumber'),
    comments: headers.indexOf('Comments')
  };
  
  if (col.comments === -1) {
    throw new Error('Comments column not found in sheet');
  }
  
  let updated = false;
  const timestamp = new Date();
  const formattedTimestamp = Utilities.formatDate(timestamp, DEFAULT_TIMEZONE, 'yyyy-MM-dd HH:mm');
  const user = Session.getEffectiveUser().getEmail();
  const fullComment = `${formattedTimestamp} (${user}): ${comment}`;
  
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (String(row[col.order]) === poNumber) {
      const currentComment = String(row[col.comments] || '').trim();
      let newComment = '';
      
      if (currentComment) {
        newComment = `${currentComment}\n${fullComment}`;
      } else {
        newComment = fullComment;
      }
      
      poSheet.getRange(r + 1, col.comments + 1).setValue(newComment);
      updated = true;
    }
  }
  
  if (!updated) {
    throw new Error(`No rows found for PO ${poNumber}`);
  }
  
  logToFeed_(poNumber, 'Comment Added', comment, '', user);
  
  return { poNumber, comment, user, timestamp: formattedTimestamp };
}

function clearAllEventIdColumns() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Clear Event IDs',
    'This will clear all calendar event IDs and force recreation of events. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('Error', 'No data found.', ui.ButtonSet.OK);
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const eventIdCols = [
    headers.indexOf('DepartEventId'),
    headers.indexOf('ArriveEventId'),
    headers.indexOf('CustomsEventId'),
    headers.indexOf('IrradiationEventId'),
    headers.indexOf('TestingEventId'),
    headers.indexOf('PackRelEventId')
  ].filter(idx => idx !== -1);
  
  for (const col of eventIdCols) {
    sheet.getRange(2, col + 1, data.length - 1, 1).clearContent();
  }
  
  ui.alert('✅ Cleared', 'All event IDs have been cleared.', ui.ButtonSet.OK);
}

function updateEventColors() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('Error', 'No data found. Please pull purchase orders first.', ui.ButtonSet.OK);
    return;
  }
  
  const dataRows = sheet.getLastRow() - 1;
  applyCompleteFormatting_(sheet, dataRows);
  applyPersistentCompletionFormatting_(sheet, dataRows);
  applyFixedDateFormatting_(sheet, dataRows);
  
  ui.alert('✅ Updated', 'Event colors and formatting have been updated.', ui.ButtonSet.OK);
}

////////////////////////////// ACTIVITY FEED //////////////////////////////
function logToFeed_(poNumber, action, description, oldValue, newValue) {
  const ss = SpreadsheetApp.getActive();
  let sheet = getOrCreateSheet_(ss, FEED_SHEET);
  
  if (sheet.getLastRow() === 0) {
    const headers = ['Timestamp', 'PO Number', 'Action', 'Description', 'Old Value', 'New Value', 'User'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e0f2f1');
    sheet.autoResizeColumns(1, headers.length);
  }
  
  const timestamp = new Date();
  const user = Session.getEffectiveUser().getEmail();
  
  const newRow = [
    timestamp,
    poNumber,
    action,
    description,
    oldValue,
    newValue,
    user
  ];
  
  sheet.insertRowAfter(1);
  sheet.getRange(2, 1, 1, newRow.length).setValues([newRow]);
  sheet.getRange(2, 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');
  
  // Keep only last 1000 entries
  if (sheet.getLastRow() > 1000) {
    sheet.deleteRow(1001);
  }
}

function storeCompletedEvent_(poNumber, milestone, completedBy, comments) {
  const ss = SpreadsheetApp.getActive();
  let sheet = getOrCreateSheet_(ss, COMPLETED_EVENTS_SHEET);
  
  if (sheet.getLastRow() === 0) {
    const headers = ['PO Number', 'Milestone', 'Completed By', 'Comments', 'Timestamp'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e0f2f1');
    sheet.autoResizeColumns(1, headers.length);
  }
  
  const timestamp = new Date();
  
  const newRow = [
    poNumber,
    milestone,
    completedBy,
    comments,
    timestamp
  ];
  
  sheet.appendRow(newRow);
  sheet.getRange(sheet.getLastRow(), 5).setNumberFormat('yyyy-MM-dd HH:mm:ss');
}

function removeCompletedEvent_(poNumber, milestone) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(COMPLETED_EVENTS_SHEET);
  
  if (!sheet || sheet.getLastRow() < 2) return;
  
  const data = sheet.getDataRange().getValues();
  let rowsToDelete = [];
  
  for (let r = data.length - 1; r >= 1; r--) {
    const row = data[r];
    if (String(row[0]) === poNumber && String(row[1]) === milestone) {
      rowsToDelete.push(r + 1);
    }
  }
  
  for (const rowNum of rowsToDelete) {
    sheet.deleteRow(rowNum);
  }
}

////////////////////////////// FIXED DATES //////////////////////////////
function setFixedDateMenu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  
  const poSheet = ss.getSheetByName(SHEET_NAME);
  if (!poSheet || poSheet.getLastRow() < 2) {
    ui.alert('Error', 'No purchase orders found. Please pull data first.', ui.ButtonSet.OK);
    return;
  }
  
  const poData = poSheet.getRange(2, 1, poSheet.getLastRow() - 1, 1).getValues();
  const poNumbers = [...new Set(poData.map(row => row[0]).filter(po => po))];
  
  if (poNumbers.length === 0) {
    ui.alert('Error', 'No purchase order numbers found.', ui.ButtonSet.OK);
    return;
  }
  
  const poList = poNumbers.map(po => `• ${po}`).join('\n');
  const poResponse = ui.prompt('Set Fixed Date', 
    `Available POs:\n${poList}\n\nEnter PO Number:`, ui.ButtonSet.OK_CANCEL);
  
  if (poResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedPO = poResponse.getResponseText().trim();
  if (!poNumbers.includes(selectedPO)) {
    ui.alert('Error', 'Invalid PO number selected.', ui.ButtonSet.OK);
    return;
  }
  
  const milestones = ['Depart Canada', 'Arrive in Aus', 'Clears Customs', 'Irradiation', 'Testing', 'Packaging and Release', 'BLS Delivery'];
  const milestoneResponse = ui.prompt('Milestone', 
    `Available milestones:\n${milestones.map(m => `• ${m}`).join('\n')}\n\nEnter milestone:`, ui.ButtonSet.OK_CANCEL);
  
  if (milestoneResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedMilestone = milestoneResponse.getResponseText().trim();
  if (!milestones.includes(selectedMilestone)) {
    ui.alert('Error', 'Invalid milestone selected.', ui.ButtonSet.OK);
    return;
  }
  
  const dateResponse = ui.prompt('Fixed Date', 'Enter date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (dateResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const dateStr = dateResponse.getResponseText().trim();
  if (!dateStr.match(/^\d{4}-\d{2}-\d{2}$/)) {
    ui.alert('Error', 'Invalid date format. Please use YYYY-MM-DD.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const result = setFixedDate_(selectedPO, selectedMilestone, dateStr);
    ui.alert('✅ Fixed Date Set', 
      `Successfully set ${selectedMilestone} to ${dateStr} for PO ${selectedPO}.`, 
      ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error', `Failed to set fixed date: ${error.message}`, ui.ButtonSet.OK);
  }
}

function clearFixedDateMenu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  
  const fixedDates = loadFixedDates_();
  const posWithFixedDates = Object.keys(fixedDates);
  
  if (posWithFixedDates.length === 0) {
    ui.alert('Info', 'No fixed dates found to clear.', ui.ButtonSet.OK);
    return;
  }
  
  const poList = posWithFixedDates.map(po => `• ${po}`).join('\n');
  const poResponse = ui.prompt('Clear Fixed Date', 
    `POs with fixed dates:\n${poList}\n\nEnter PO Number:`, ui.ButtonSet.OK_CANCEL);
  
  if (poResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedPO = poResponse.getResponseText().trim();
  if (!posWithFixedDates.includes(selectedPO)) {
    ui.alert('Error', 'Invalid PO number or no fixed dates for this PO.', ui.ButtonSet.OK);
    return;
  }
  
  const milestones = Object.keys(fixedDates[selectedPO]);
  const milestoneResponse = ui.prompt('Milestone', 
    `Fixed dates for ${selectedPO}:\n${milestones.map(m => `• ${m}`).join('\n')}\n\nEnter milestone:`, ui.ButtonSet.OK_CANCEL);
  
  if (milestoneResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedMilestone = milestoneResponse.getResponseText().trim();
  if (!milestones.includes(selectedMilestone)) {
    ui.alert('Error', 'Invalid milestone selected.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const result = clearFixedDate_(selectedPO, selectedMilestone);
    ui.alert('✅ Fixed Date Cleared', 
      `Successfully cleared ${selectedMilestone} fixed date for PO ${selectedPO}.`, 
      ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error', `Failed to clear fixed date: ${error.message}`, ui.ButtonSet.OK);
  }
}

function setFixedDate_(poNumber, milestone, dateStr) {
  const ss = SpreadsheetApp.getActive();
  let sheet = getOrCreateSheet_(ss, FIXED_DATES_SHEET);
  
  if (sheet.getLastRow() === 0) {
    const headers = ['PO Number', 'Milestone', 'Fixed Date', 'Set By', 'Timestamp'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e0f2f1');
    sheet.autoResizeColumns(1, headers.length);
  }
  
  const data = sheet.getDataRange().getValues();
  let existingRow = -1;
  
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][0]) === poNumber && String(data[r][1]) === milestone) {
      existingRow = r;
      break;
    }
  }
  
  const timestamp = new Date();
  const user = Session.getEffectiveUser().getEmail();
  const newRow = [
    poNumber,
    milestone,
    dateStr,
    user,
    timestamp
  ];
  
  if (existingRow !== -1) {
    sheet.getRange(existingRow + 1, 1, 1, newRow.length).setValues([newRow]);
  } else {
    sheet.appendRow(newRow);
  }
  
  sheet.getRange(sheet.getLastRow(), 5).setNumberFormat('yyyy-MM-dd HH:mm:ss');
  
  logToFeed_(poNumber, 'Fixed Date Set', `${milestone} set to ${dateStr}`, '', dateStr);
  
  return { poNumber, milestone, dateStr, user, timestamp };
}

function clearFixedDate_(poNumber, milestone) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(FIXED_DATES_SHEET);
  
  if (!sheet || sheet.getLastRow() < 2) return;
  
  const data = sheet.getDataRange().getValues();
  let rowsToDelete = [];
  
  for (let r = data.length - 1; r >= 1; r--) {
    const row = data[r];
    if (String(row[0]) === poNumber && String(row[1]) === milestone) {
      rowsToDelete.push(r + 1);
    }
  }
  
  for (const rowNum of rowsToDelete) {
    sheet.deleteRow(rowNum);
  }
  
  logToFeed_(poNumber, 'Fixed Date Cleared', `${milestone} fixed date removed`, '', '');
  
  return { poNumber, milestone };
}

function loadFixedDates_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(FIXED_DATES_SHEET);
  const fixedDates = {};
  
  if (!sheet || sheet.getLastRow() < 2) return fixedDates;
  
  const data = sheet.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const po = String(row[0] || '').trim();
    const milestone = String(row[1] || '').trim();
    const dateStr = String(row[2] || '').trim();
    
    if (po && milestone && dateStr) {
      if (!fixedDates[po]) fixedDates[po] = {};
      fixedDates[po][milestone] = dateStr;
    }
  }
  
  return fixedDates;
}

//////////////////////// FIXED: PROPER FIXED DATE FORMATTING - SIMPLE ISO DATES ////////////////////////
function applyFixedDateFormatting_(sheet, dataRows) {
  if (dataRows < 1) return;
  
  const fixedDates = loadFixedDates_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const col = {
    order: headers.indexOf('OrderNumber'),
    depart: headers.indexOf('Depart Canada'),
    arrive: headers.indexOf('Arrive in Aus'),
    customs: headers.indexOf('Clears Customs'),
    irrad: headers.indexOf('Irradiation'),
    testing: headers.indexOf('Testing'),
    release: headers.indexOf('Packaging and Release')
  };
  
  const milestoneMap = {
    'Depart Canada': col.depart,
    'Arrive in Aus': col.arrive,
    'Clears Customs': col.customs,
    'Irradiation': col.irrad,
    'Testing': col.testing,
    'Packaging and Release': col.release,
    'BLS Delivery': col.depart
  };
  
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const poNumber = String(row[col.order] || '').trim();
    
    if (poNumber && fixedDates[poNumber]) {
      for (const [milestone, colIndex] of Object.entries(milestoneMap)) {
        if (fixedDates[poNumber][milestone]) {
          const cell = sheet.getRange(r + 1, colIndex + 1);
          const fixedDate = fixedDates[poNumber][milestone];
          
          // ENSURE SIMPLE ISO DATE FORMAT (strip time component if present)
          let cleanDate = fixedDate;
          if (fixedDate.includes('T')) {
            cleanDate = fixedDate.split('T')[0];
          }
          if (fixedDate.includes(' ')) {
            cleanDate = fixedDate.split(' ')[0];
          }
          
          // Set the clean ISO date and formatting
          cell.setValue(cleanDate)
              .setBackground(COLOR_FIXED_DATE)
              .setFontColor('#000000')
              .setNote(`Fixed date: ${cleanDate}`);
        }
      }
    }
  }
}

function checkUnleashedDateFields() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const pos = fetchPurchaseOrders_({ pageSize: 1 });
    
    if (pos.length === 0) {
      ui.alert('No POs Found', 'No purchase orders returned from API.', ui.ButtonSet.OK);
      return;
    }
    
    const po = pos[0];
    
    // Log to console
    console.log('=== FULL PO OBJECT ===');
    console.log(JSON.stringify(po, null, 2));
    
    // Check all date-related fields
    let report = `PO: ${po.OrderNumber}\n\n`;
    report += `po.DeliveryDate: ${po.DeliveryDate || 'NOT FOUND'}\n`;
    report += `po.RequiredDate: ${po.RequiredDate || 'NOT FOUND'}\n`;
    report += `po.OrderDate: ${po.OrderDate || 'NOT FOUND'}\n`;
    
    if (po.PurchaseOrderLines && po.PurchaseOrderLines.length > 0) {
      const line = po.PurchaseOrderLines[0];
      report += `\nFirst Line:\n`;
      report += `line.DeliveryDate: ${line.DeliveryDate || 'NOT FOUND'}\n`;
      report += `line.RequiredDate: ${line.RequiredDate || 'NOT FOUND'}\n`;
    }
    
    report += `\n\nALL PO FIELDS:\n${Object.keys(po).join(', ')}`;
    
    ui.alert('API Fields Found', report, ui.ButtonSet.OK);
    console.log(report);
    
  } catch (error) {
    ui.alert('Error', error.message, ui.ButtonSet.OK);
  }
}

////////////////////////////// MAIN WORKFLOW //////////////////////////////
function pullAndSyncAll() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.alert('🚀 Starting Full Sync', 'This will pull POs, sync calendar events, and cleanup old events. Continue?', ui.ButtonSet.OK_CANCEL);
    
    // Pull purchase orders
    pullOpenPurchaseOrders();
    
    // Sync calendar events
    syncCalendarEventsFromSheet();
    
    // Cleanup old calendar events
    cleanupCalendarEvents();
    
    ui.alert('✅ Full Sync Complete', 'All operations completed successfully!', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Sync Failed', `Error during sync: ${error.message}`, ui.ButtonSet.OK);
  }
}
