/**
 * ============================================================================
 * NOTE TAKER APP - BACKEND (Code.gs)
 * ============================================================================
 * Handles web app routing, spreadsheet schema ingestion, campaign targeting,
 * zip code database lookups, email template retrieval, and feedback logging.
 */

// Global Configuration & Sheet IDs
const CONFIG = {
  // Spreadsheet ID is dynamically fetched from the bound file
  get SPREADSHEET() {
    return SpreadsheetApp.getActiveSpreadsheet();
  },
  SETTINGS_TAB: 'Settings',
  ZIP_LOOKUP_SS_ID: '1O7L6NVy7N4vS47IMqlGZmExdOD13l8e8kmAuXw2AAKw' // Configured in Settings!B9
};

/**
 * Serves the web app UI (Index.html / Notes.html)
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Note Taker App')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Standard Apps Script include helper for modular HTML/CSS/JS files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ============================================================================
 * 1. FORM SCHEMA & CAMPAIGN INGESTION
 * ========================================================================== */

/**
 * Retrieves the full schema dictionary for all forms.
 * Reads column configurations and associates active campaigns.
 */
function getFormSchema(department, agentExtension) {
  const ss = CONFIG.SPREADSHEET;
  const schema = {};

  // Form tabs defining schema structures (8 standard columns)
  const formTabs = [
    { key: 'inbound', name: 'Inbound' },
    { key: 'outbound', name: 'Outbound' },
    { key: 'residential', name: 'Residential' },
    { key: 'dealer', name: 'Dealer' },
    { key: 'addendum', name: 'Addendum' },
    { key: 'parts', name: 'Parts' },
    { key: 'mx', name: 'MX' },
    { key: 'mx_outbound', name: 'MX Outbound' },
    { key: 'retention', name: 'Retention' },
    { key: 'retention_outbound', name: 'Retention Outbound' },
    { key: 'login', name: 'login' }
  ];

  formTabs.forEach(tabInfo => {
    const sheet = ss.getSheetByName(tabInfo.name);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      schema[tabInfo.key] = [];
      schema[tabInfo.key + '_order'] = [];
      return;
    }

    data.shift(); // Remove header row

    const fields = data
      .filter(row => row[0] && String(row[0]).trim() !== '')
      .map(row => {
        return {
          id: String(row[0]).trim(),
          internalId: String(row[0]).trim(),
          label: String(row[1] || row[0]).trim(),
          displayLabel: String(row[1] || row[0]).trim(),
          type: String(row[2] || 'Text').trim().toLowerCase(),
          fieldType: String(row[2] || 'Text').trim(),
          options: parseOptions(row[3]),
          visible: parseBoolean(row[4], true),
          required: parseBoolean(row[5], false),
          dependsOn: String(row[6] || '').trim(),
          copyIfBlank: parseBoolean(row[7], false)
        };
      });

    schema[tabInfo.key] = fields;
    schema[tabInfo.key + '_order'] = fields;
  });

  // Inject departmental campaigns if department is provided
  if (department) {
    schema.campaigns = getCampaigns(department, agentExtension);
  }

  return schema;
}

/**
 * Loads campaigns from the department-specific campaign tab.
 * Includes the 10th column: Auto-Populate Reason.
 */
function getCampaigns(department, agentExtension) {
  const ss = CONFIG.SPREADSHEET;
  const tabName = 'Campaigns - ' + department;
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  data.shift(); // Remove header row

  return data
    .filter(row => {
      const campaignId = String(row[0] || '').trim();
      const isActive = parseBoolean(row[4], false);
      if (!campaignId || !isActive) return false;

      // Targeting evaluation (Global vs. Extension-Specific)
      const targetScope = String(row[5] || 'Global').trim().toLowerCase();
      if (targetScope === 'extension-specific' && agentExtension) {
        const targetExtensions = String(row[6] || '')
          .split(',')
          .map(ext => ext.trim());
        return targetExtensions.includes(String(agentExtension).trim());
      }
      return true;
    })
    .map(row => ({
      id: String(row[0]).trim(),
      internalId: String(row[0]).trim(),
      label: String(row[1] || row[0]).trim(),
      displayLabel: String(row[1] || row[0]).trim(),
      type: String(row[2] || 'Radio').trim().toLowerCase(),
      fieldType: String(row[2] || 'Radio').trim(),
      options: parseOptions(row[3] || 'Yes,No'),
      active: parseBoolean(row[4], true),
      targetScope: String(row[5] || 'Global').trim(),
      targetExtensions: String(row[6] || '').split(',').map(s => s.trim()).filter(Boolean),
      required: parseBoolean(row[7], false),
      notes: String(row[8] || '').trim(),
      // ---------------------------------------------------------------
      // COLUMN J: AUTO-POPULATE REASON
      // ---------------------------------------------------------------
      autoPopulateReason: String(row[9] || '').trim()
    }));
}

/* ============================================================================
 * 2. ZIP CODE LOOKUP (FOOTPRINT & DEALER SERVICE)
 * ========================================================================== */

/**
 * Validates zip codes against the Official Pro Zips and Dealer Service tabs.
 */
function checkZipCode(zip) {
  const cleanZip = String(zip || '').trim().padStart(5, '0');
  if (!cleanZip || cleanZip.length !== 5) {
    return { found: false, message: 'Invalid ZIP format.' };
  }

  let targetSs;
  try {
    targetSs = SpreadsheetApp.openById(CONFIG.ZIP_LOOKUP_SS_ID);
  } catch (err) {
    targetSs = CONFIG.SPREADSHEET;
  }

  // 1. Check Official Pro Zips (In Footprint)
  const proSheet = targetSs.getSheetByName('Official Pro Zips');
  if (proSheet) {
    const proZips = proSheet.getDataRange().getValues().flat().map(String);
    if (proZips.includes(cleanZip)) {
      return { found: true, type: 'in_footprint', label: '✅ In Footprint' };
    }
  }

  // 2. Check Official DS (Upgrade & Service)
  const dsSheet = targetSs.getSheetByName('Official DS (Upgrade & Service)');
  if (dsSheet) {
    const values = dsSheet.getDataRange().getValues();
    const numCols = values[0].length;

    for (let c = 0; c < numCols; c++) {
      const headerRow1 = String(values[0][c] || '').toUpperCase();
      const headerRow2 = values.length > 1 ? String(values[1][c] || '').toUpperCase() : '';
      const isNoTech = headerRow1.includes('NO TECH') || headerRow2.includes('NO TECH');

      for (let r = 2; r < values.length; r++) {
        if (String(values[r][c]).trim() === cleanZip) {
          if (isNoTech) {
            return { found: false, type: 'no_tech', label: '❌ No Tech Available' };
          }
          return { found: true, type: 'dealer_service', label: '✅ Dealer Service' };
        }
      }
    }
  }

  return { found: false, type: 'not_found', label: '❌ No Tech Available' };
}

/* ============================================================================
 * 3. QUICK LINKS, TEMPLATES, FEEDBACK & GOVERNANCE
 * ========================================================================== */

/**
 * Retrieves department-specific quick links
 */
function getQuickLinks(department) {
  const sheet = CONFIG.SPREADSHEET.getSheetByName('Links - ' + department);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  data.shift(); // Remove header

  return data
    .filter(r => parseBoolean(r[4], true) && r[0])
    .map(r => ({
      title: r[0],
      url: r[1],
      category: r[2],
      popupMode: parseBoolean(r[3], false),
      notes: r[5]
    }));
}

/**
 * Retrieves special email templates
 */
function getSpecialEmails() {
  const sheet = CONFIG.SPREADSHEET.getSheetByName('Special Emails');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  data.shift();

  return data
    .filter(r => r[0])
    .map(r => ({
      id: String(r[0]).trim(),
      brand: String(r[1]).trim(),
      subject: String(r[2]).trim(),
      body: String(r[3] || '').trim(),
      type: 'special'
    }));
}

/**
 * Appends feedback submitted by agents to the Feedback sheet
 */
function submitFeedback(payload) {
  const sheet = CONFIG.SPREADSHEET.getSheetByName('Feedback');
  if (!sheet) throw new Error('Feedback sheet not found.');

  const row = [
    new Date(),
    payload.agentName || '',
    payload.extension || '',
    payload.department || '',
    payload.category || 'General',
    payload.applicableForm || '',
    payload.details || ''
  ];

  sheet.appendRow(row);
  return { success: true };
}

/**
 * Fetches the system change log records for the UI modal
 */
function getChangeLog() {
  const sheet = CONFIG.SPREADSHEET.getSheetByName('Change Log');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  data.shift();

  return data
    .filter(r => parseBoolean(r[5], true) && r[0])
    .map(r => ({
      date: r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'MM/dd/yyyy') : String(r[0]),
      version: String(r[1]),
      applicableForm: String(r[2]),
      description: String(r[3]),
      updatedBy: String(r[4])
    }))
    .reverse();
}

/**
 * Verifies admin passcode from Settings!B2 (Admin Password)
 */
function verifyAdminPasscode(enteredPasscode) {
  const sheet = CONFIG.SPREADSHEET.getSheetByName(CONFIG.SETTINGS_TAB);
  if (!sheet) return false;

  const actualPassword = String(sheet.getRange('B2').getValue()).trim();
  return String(enteredPasscode).trim() === actualPassword;
}

/* ============================================================================
 * 4. INTERNAL UTILITY HELPERS
 * ========================================================================== */

function parseOptions(val) {
  if (!val) return [];
  if (Array.isArray(val)) return val;
  return String(val)
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);
}

function parseBoolean(val, defaultVal = false) {
  if (val === undefined || val === null || val === '') return defaultVal;
  if (typeof val === 'boolean') return val;
  const s = String(val).trim().toUpperCase();
  return s === 'TRUE' || s === 'YES' || s === '1';
}
