/**
 * ===== Web App bootstrap =====
 * Deploy: Deploy → New deployment → Web app
 * Execute as: Me     •     Who has access: Your org (or Anyone with link)
 */

// REMOVE THESE DUPLICATE CONSTANTS - They're already in your main GS file
// const SHEET_NAME = 'OpenPOs';
// const FEED_SHEET = 'Activity Feed';
// const PERMIT_NUMBERS_SHEET = 'Permit Numbers';
// const ASSUMPTIONS_SHEET = 'OpenPOs Assumptions';
// const COMPLETED_EVENTS_SHEET = 'Completed Events';
// const FIXED_DATES_SHEET = 'Fixed Dates';
// const REQUIREMENTS_SHEET = 'OpenPOs Requirements';
// const DEFAULT_TIMEZONE = 'Australia/Brisbane';

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

/** Enhanced API wrapper for better error responses */
function apiCallWrapper(fn, ...args) {
  try {
    const result = fn(...args);
    return { success: true, data: result };
  } catch (error) {
    console.error(`API Error in ${fn.name}:`, error);
    return { 
      success: false, 
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/** Mobile API: list POs (lightweight) with simple search - FIXED FOR WEB APP */
function apiListPOs(query) {
  try {
    console.log('apiListPOs called with query:', query);
    
    // CRITICAL FIX: Use getActiveSpreadsheet() for web app compatibility
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('OpenPOs'); // Use string directly instead of constant
    
    if (!sh) {
      console.error('Sheet not found: OpenPOs');
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

/** Get single PO details for mobile detail view */
function apiGetPODetails(poNumber) {
  return apiCallWrapper(() => {
    const sh = _requireSheetRows_('OpenPOs');
    const values = sh.getDataRange().getValues();
    const headers = values[0];
    
    const idx = {
      order: headers.indexOf('OrderNumber'),
      status: headers.indexOf('OrderStatus'),
      supplier: headers.indexOf('SupplierName'),
      product: headers.indexOf('ProductDescription'),
      qty: headers.indexOf('OrderedQuantity'),
      comments: headers.indexOf('Comments'),
      depart: headers.indexOf('Depart Canada'),
      arrive: headers.indexOf('Arrive in Aus'),
      customs: headers.indexOf('Clears Customs'),
      irrad: headers.indexOf('Irradiation'),
      test: headers.indexOf('Testing'),
      release: headers.indexOf('Packaging and Release'),
      warehouse: headers.indexOf('WarehouseCode')
    };
    
    const poRow = values.find(row => String(row[idx.order]).trim() === poNumber.trim());
    if (!poRow) throw new Error(`PO ${poNumber} not found`);
    
    return {
      po: poNumber,
      status: poRow[idx.status],
      supplier: poRow[idx.supplier],
      product: poRow[idx.product],
      quantity: poRow[idx.qty],
      warehouse: poRow[idx.warehouse],
      comments: poRow[idx.comments],
      dates: {
        'Depart Canada': toISODateCell_(poRow[idx.depart]) || '',
        'Arrive in Aus': toISODateCell_(poRow[idx.arrive]) || '',
        'Clears Customs': toISODateCell_(poRow[idx.customs]) || '',
        'Irradiation': toISODateCell_(poRow[idx.irrad]) || '',
        'Testing': toISODateCell_(poRow[idx.test]) || '',
        'Packaging and Release': toISODateCell_(poRow[idx.release]) || ''
      }
    };
  });
}

/** Quick stats for dashboard */
function apiGetStats() {
  return apiCallWrapper(() => {
    const pos = apiListPOs('').data || [];
    const completedEvents = loadCompletedEvents_();
    const feed = apiGetFeed(50).data || [];
    
    let completedCount = 0;
    Object.values(completedEvents).forEach(events => {
      completedCount += events.size;
    });
    
    const recentActivity = feed.filter(f => {
      const hoursAgo = (new Date() - new Date(f.timestamp)) / (1000 * 60 * 60);
      return hoursAgo < 24;
    }).length;
    
    return {
      totalPOs: pos.length,
      completedMilestones: completedCount,
      recentActivity: recentActivity,
      statusBreakdown: countStatuses(pos)
    };
  });
}

function countStatuses(pos) {
  const statusCount = {};
  pos.forEach(p => {
    statusCount[p.status] = (statusCount[p.status] || 0) + 1;
  });
  return statusCount;
}

/** Activity feed (latest first) */
function apiGetFeed(limit) {
  return apiCallWrapper(() => {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Activity Feed');
    if (!sh || sh.getLastRow() < 2) return [];
    const vals = sh.getRange(2, 1, Math.min(sh.getLastRow()-1, 2000), 7).getValues();
    return vals.map(r => ({
      timestamp:  r[0] ? Utilities.formatDate(new Date(r[0]), 'Australia/Brisbane', 'yyyy-MM-dd HH:mm') : '',
      poNumber:  r[1],
      action: r[2],
      description: r[3],
      oldValue: r[4],
      newValue: r[5],
      user: r[6]
    })).slice(0, limit || 200);
  });
}

/** Core actions wired to your existing functions */
function apiPullAndSyncAll() {
  return apiCallWrapper(() => {
    // Import the function from main script
    const result = pullAndSyncAll();
    return { message: 'Full sync completed successfully' };
  });
}

/** Validate PO number format */
function validatePONumber(poNumber) {
  return typeof poNumber === 'string' && poNumber.trim().length > 0;
}

/** Validate milestone */
function validateMilestone(milestone) {
  const validMilestones = [
    'Depart Canada', 'Arrive in Aus', 'Clears Customs', 
    'Irradiation', 'Testing', 'Packaging and Release', 'BLS Delivery'
  ];
  return validMilestones.includes(milestone);
}

function apiCompleteEvent(po, milestone, comments) {
  return apiCallWrapper(() => {
    if (!validatePONumber(po)) {
      throw new Error('Invalid PO number');
    }
    if (!validateMilestone(milestone)) {
      throw new Error('Invalid milestone');
    }
    
    const user = Session.getEffectiveUser().getEmail();
    const res = markEventComplete_(po, milestone, user, comments || '');
    return { 
      message: `Successfully completed ${milestone} for PO ${po}`,
      details: res 
    };
  });
}

function apiUncompleteEvent(po, milestone, reason) {
  return apiCallWrapper(() => {
    if (!validatePONumber(po)) {
      throw new Error('Invalid PO number');
    }
    if (!validateMilestone(milestone)) {
      throw new Error('Invalid milestone');
    }
    
    const res = uncompleteEvent_(po, milestone, reason || '');
    return { 
      message: `Successfully uncompleted ${milestone} for PO ${po}`,
      details: res 
    };
  });
}

function apiSetFixedDate(po, milestone, dateStr) {
  return apiCallWrapper(() => {
    if (!validatePONumber(po)) {
      throw new Error('Invalid PO number');
    }
    if (!validateMilestone(milestone)) {
      throw new Error('Invalid milestone');
    }
    if (!dateStr.match(/^\d{4}-\d{2}-\d{2}$/)) {
      throw new Error('Invalid date format. Use YYYY-MM-DD');
    }
    
    const res = setFixedDate_(po, milestone, dateStr);
    return { 
      message: `Fixed date set for ${milestone} on PO ${po}`,
      details: res 
    };
  });
}

function apiClearFixedDate(po, milestone) {
  return apiCallWrapper(() => {
    if (!validatePONumber(po)) {
      throw new Error('Invalid PO number');
    }
    if (!validateMilestone(milestone)) {
      throw new Error('Invalid milestone');
    }
    
    const res = clearFixedDate_(po, milestone);
    return { 
      message: `Fixed date cleared for ${milestone} on PO ${po}`,
      details: res 
    };
  });
}

function apiAssignPermitNumber(po, permit) {
  return apiCallWrapper(() => {
    if (!validatePONumber(po)) {
      throw new Error('Invalid PO number');
    }
    if (!permit || typeof permit !== 'string' || permit.trim().length === 0) {
      throw new Error('Permit number cannot be empty');
    }
    
    const res = assignPermitNumber_(po, permit.trim());
    return { 
      message: `Permit number ${permit} assigned to PO ${po}`,
      details: res 
    };
  });
}

function apiGetPermits() {
  return apiCallWrapper(() => {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('Permit Numbers');
    const permits = [];
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
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
    
    return permits;
  });
}

function apiGetAssumptions() {
  return apiCallWrapper(() => {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('OpenPOs Assumptions');
    const assumptions = [];
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
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
    
    return assumptions;
  });
}

function apiUpdateAssumption(name, value) {
  return apiCallWrapper(() => {
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
    
    return { name: name, value: value };
  });
}

// WEB APP COMPATIBLE FUNCTIONS - These call the main functions
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
    // Call the main function
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('OpenPOs');
    if (sheet && sheet.getLastRow() > 1) {
      const dataRows = sheet.getLastRow() - 1;
      applyCompleteFormatting_(sheet, dataRows);
      applyPersistentCompletionFormatting_(sheet, dataRows);
      applyFixedDateFormatting_(sheet, dataRows);
    }
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

// Utility functions that need to be available
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
        return Utilities.formatDate(d, 'Australia/Brisbane', 'yyyy-MM-dd');
      }
    }
    
    // Handle simple YYYY-MM-DD strings
    if (cellValue.match(/^\d{4}-\d{2}-\d{2}$/)) return cellValue;
    
    // Handle ISO strings with time components
    const isoMatch = cellValue.match(/^(\d{4}-\d{2}-\d{2})T/);
    if (isoMatch) return isoMatch[1];
    
    // Fallback: try direct parsing
    const d = new Date(cellValue);
    return isNaN(d) ? '' : Utilities.formatDate(d, 'Australia/Brisbane', 'yyyy-MM-dd');
  }
  
  if (cellValue instanceof Date && !isNaN(cellValue)) {
    return Utilities.formatDate(cellValue, 'Australia/Brisbane', 'yyyy-MM-dd');
  }
  
  return '';
}
