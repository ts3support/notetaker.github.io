/**
 * ============================================================================
 * NOTE TAKER APPLICATION - BACKEND ENGINE (Code.gs)
 * AAA Smart Home / A3 Smart Home
 * ============================================================================
 * Files in project: Code.gs, Login.html, Notes.html
 * ============================================================================
 * Handles:
 *  - Web app routing (doGet) for Login.html & Notes.html
 *  - Dynamic multi-department campaign evaluation & extension targeting
 *  - Dynamic:Campaigns placeholder resolution (Label: NON-COMM, Options: Yes,No)
 *  - Form schema parsing (Inbound, Outbound, Retention, MX, Tickets, Parts, etc.)
 *  - Google Docs integration (Helpful Links & Standard Email Templates)
 *  - Special emails & admin passcode validation
 *  - Footprint / Territory Zip Code lookup
 *  - Agent feedback submission & change log audit queries
 * ============================================================================
 */

// ============================================================================
// 1. ROUTING (doGet) - CASE-INSENSITIVE FOR LOGIN.HTML & NOTES.HTML
// ============================================================================

function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const requested = String(params.page || '').trim().toLowerCase();

  // If user requested notes (or already passed agent/extension params), load Notes.html
  // Otherwise, start on Login.html
  let targetFile = 'Login';

  if (requested === 'notes' || requested === 'main' || requested === 'app' || requested === 'index') {
    targetFile = 'Notes';
  } else if (requested === 'login') {
    targetFile = 'Login';
  } else if (params.agent || params.extension || params.ext || params.user) {
    targetFile = 'Notes';
  } else {
    targetFile = 'Login';
  }

  // Load template with strict fallback to exact project names (Notes.html / Login.html)
  let template;
  try {
    template = HtmlService.createTemplateFromFile(targetFile);
  } catch (err) {
    try {
      template = HtmlService.createTemplateFromFile(targetFile.toLowerCase());
    } catch (e2) {
      template = HtmlService.createTemplateFromFile(targetFile === 'Notes' ? 'Login' : 'Notes');
    }
  }

  // Pass URL and parameters to the HTML template
  template.appUrl = ScriptApp.getService().getUrl();
  template.params = params;
  template.agent = params.agent || params.name || '';
  template.extension = params.extension || params.ext || '';
  template.department = params.department || params.dept || '';
  template.role = params.role || params.role_tier || '';

  return template.evaluate()
    .setTitle('AAA Smart Home - Note Taker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  const candidates = [
    filename,
    filename.charAt(0).toUpperCase() + filename.slice(1),
    filename.toLowerCase(),
    filename.toUpperCase()
  ];
  for (let i = 0; i < candidates.length; i++) {
    try {
      return HtmlService.createHtmlOutputFromFile(candidates[i]).getContent();
    } catch (e) {}
  }
  return '';
}

// ============================================================================
// 2. SETTINGS HELPER
// ============================================================================

function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Settings');
  const settings = {};

  if (!sheet) return settings;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0] || '').trim();
    const val = data[i][1];
    if (key) {
      settings[key] = val;
    }
  }
  return settings;
}

// ============================================================================
// 3. DYNAMIC CAMPAIGN RESOLVER ENGINE
// ============================================================================

/**
 * Resolves active campaigns targeted to an agent by Department and Extension.
 *
 * @param {string} department - e.g., "Retention", "Tech Support", "Member Experience"
 * @param {string|number} extension - e.g., "1306228"
 * @returns {Array<Object>} List of active campaigns matching targeting rules
 */
function getActiveDepartmentCampaigns(department, extension) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const extStr = String(extension || '').trim();
  const deptStr = String(department || '').trim().toLowerCase();

  // Resolve matching department campaign tab
  let targetTab = '';
  if (deptStr.includes('retention')) {
    targetTab = 'Campaigns - Retention';
  } else if (deptStr.includes('tech') || deptStr.includes('support')) {
    targetTab = 'Campaigns - Tech Support';
  } else if (deptStr.includes('member') || deptStr.includes('mx') || deptStr.includes('experience')) {
    targetTab = 'Campaigns - Member Experience';
  } else {
    targetTab = 'Campaigns - ' + department;
  }

  let sheet = ss.getSheetByName(targetTab);
  // Fallback to Retention if sheet not found by explicit name
  if (!sheet) {
    sheet = ss.getSheetByName('Campaigns - Retention');
  }
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const colIndex = {
    id: headers.indexOf('campaign id'),
    label: headers.indexOf('display label'),
    type: headers.indexOf('field type'),
    options: headers.indexOf('options'),
    active: headers.indexOf('active'),
    scope: headers.indexOf('target scope'),
    extensions: headers.indexOf('target extensions'),
    required: headers.indexOf('required'),
    notes: headers.indexOf('notes')
  };

  const activeCampaigns = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Check Active boolean
    const isActive = row[colIndex.active] === true || String(row[colIndex.active]).trim().toUpperCase() === 'TRUE';
    if (!isActive) continue;

    const campaignId = String(row[colIndex.id] || '').trim();
    if (!campaignId) continue;

    const scope = String(row[colIndex.scope] || '').trim().toLowerCase();
    const rawExtensions = String(row[colIndex.extensions] || '').trim();
    const targetExtList = rawExtensions.split(',').map(e => e.trim());

    // Targeting rules:
    // 1. Department-Wide or extension list has "ALL"
    // 2. Extension-Specific and extension list contains agent's extension
    // 3. If extension is empty/unknown, allow Department-Wide campaigns to display by default
    let isTargeted = false;
    if (scope === 'department-wide' || targetExtList.includes('ALL')) {
      isTargeted = true;
    } else if (scope === 'extension-specific' && extStr && targetExtList.includes(extStr)) {
      isTargeted = true;
    }

    if (isTargeted) {
      activeCampaigns.push({
        campaignId: campaignId,
        id: campaignId,
        label: row[colIndex.label] || campaignId.toUpperCase(),
        type: row[colIndex.type] || 'Radio',
        options: String(row[colIndex.options] || 'Yes,No').split(',').map(o => o.trim()),
        required: row[colIndex.required] === true || String(row[colIndex.required]).trim().toUpperCase() === 'TRUE',
        targetScope: row[colIndex.scope] || 'Department-Wide',
        notes: row[colIndex.notes] || ''
      });
    }
  }

  return activeCampaigns;
}

// ============================================================================
// 4. LOGIN FORM CONFIG & VALIDATION HELPERS
// ============================================================================

function getLoginFields() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('login') || ss.getSheetByName('Login') || ss.getSheetByName('Index');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const fields = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      fields.push({
        id: String(row[0]).trim(),
        internalId: String(row[0]).trim(),
        label: String(row[1] || row[0]).trim(),
        type: String(row[2] || 'Text').trim(),
        options: String(row[3] || '').split(',').map(o => o.trim()).filter(Boolean),
        visible: row[4] === true || String(row[4]).toUpperCase() === 'TRUE',
        required: row[5] === true || String(row[5]).toUpperCase() === 'TRUE'
      });
    }
  }
  return fields;
}

function getLoginConfig() {
  return getLoginFields();
}

function validateLogin(agentName, extension, department) {
  return {
    success: true,
    agentName: agentName,
    extension: extension,
    department: department
  };
}

function getProjectExtensions() {
  const settings = getSettings();
  const raw = String(settings['Project 2026 Extensions'] || '');
  return raw.split(',').map(e => e.trim()).filter(Boolean);
}

function isProjectExtension(ext) {
  const exts = getProjectExtensions();
  return exts.includes(String(ext).trim());
}

// ============================================================================
// 5. FORM SCHEMA LOADER (getFormSchema) - DYNAMIC RESOLUTION
// ============================================================================

/**
 * Reads form definitions from all service tabs and dynamically resolves campaign settings.
 * Replaces "Dynamic:Campaigns" with the active campaign controls (Label: NON-COMM, Options: Yes,No).
 */
function getFormSchema(department, extension) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const extStr = String(extension || '').trim();
  const deptStr = String(department || '').trim();

  const schema = {
    campaigns: {}
  };

  const formSheets = [
    { key: 'login', tab: 'login', defaultDept: 'Tech Support' },
    { key: 'index', tab: 'login', defaultDept: 'Tech Support' },
    { key: 'inbound', tab: 'Inbound', defaultDept: 'Tech Support' },
    { key: 'outbound', tab: 'Outbound', defaultDept: 'Tech Support' },
    { key: 'retention', tab: 'Retention', defaultDept: 'Retention' },
    { key: 'retention_outbound', tab: 'Retention Outbound', defaultDept: 'Retention' },
    { key: 'residential', tab: 'Residential', defaultDept: 'Tech Support' },
    { key: 'dealer', tab: 'Dealer', defaultDept: 'Tech Support' },
    { key: 'addendum', tab: 'Addendum', defaultDept: 'Tech Support' },
    { key: 'parts', tab: 'Parts', defaultDept: 'Tech Support' },
    { key: 'mx', tab: 'MX', defaultDept: 'Member Experience' },
    { key: 'mx_outbound', tab: 'MX Outbound', defaultDept: 'Member Experience' }
  ];

  formSheets.forEach(item => {
    const sheet = ss.getSheetByName(item.tab);
    if (!sheet) {
      schema[item.key] = {};
      schema[item.key + '_order'] = [];
      return;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      schema[item.key] = {};
      schema[item.key + '_order'] = [];
      return;
    }

    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const idIdx = headers.indexOf('internal id');
    const labelIdx = headers.indexOf('display label');
    const typeIdx = headers.indexOf('field type');
    const optIdx = headers.indexOf('options');
    const visIdx = headers.indexOf('visible');
    const reqIdx = headers.indexOf('required');
    const depIdx = headers.indexOf('depends on');
    const copyIdx = headers.indexOf('copy if blank');

    const fieldsMap = {};
    const fieldsOrder = [];

    // Determine the department to evaluate for this specific form
    const effectiveDept = deptStr || item.defaultDept;

    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      const internalId = String(row[idIdx] || '').trim();
      if (!internalId) continue;

      let fieldLabel = String(row[labelIdx] || internalId).trim();
      let fieldType = String(row[typeIdx] || 'Text').trim();
      let fieldOptions = String(row[optIdx] || '').trim();
      let isVisible = row[visIdx] === true || String(row[visIdx]).toUpperCase() === 'TRUE';
      let isRequired = row[reqIdx] === true || String(row[reqIdx]).toUpperCase() === 'TRUE';

      // ======================================================================
      // INTERCEPT DYNAMIC CAMPAIGN PLACEHOLDER
      // ======================================================================
      if (internalId === 'campaign_toggle' || fieldOptions.toLowerCase().includes('dynamic:campaigns')) {
        const campaigns = getActiveDepartmentCampaigns(effectiveDept, extStr);

        if (campaigns && campaigns.length > 0) {
          const activeCampaign = campaigns[0];
          fieldLabel = activeCampaign.label || 'NON-COMM';
          fieldType = activeCampaign.type || 'Radio';
          fieldOptions = Array.isArray(activeCampaign.options) 
            ? activeCampaign.options.join(',') 
            : String(activeCampaign.options || 'Yes,No');
          isVisible = true;
          isRequired = activeCampaign.required === true;
        } else {
          // If no active or targeted campaign for this department/agent, hide the row
          isVisible = false;
          fieldOptions = '';
        }
      }

      const fieldObj = {
        id: internalId,
        internalId: internalId,
        label: fieldLabel,
        type: fieldType,
        options: fieldOptions,
        visible: isVisible,
        required: isRequired,
        dependsOn: String(row[depIdx] || '').trim(),
        copyIfBlank: row[copyIdx] === true || String(row[copyIdx]).toUpperCase() === 'TRUE'
      };

      fieldsMap[internalId] = fieldObj;
      fieldsOrder.push(fieldObj);
    }

    schema[item.key] = fieldsMap;
    schema[item.key + '_order'] = fieldsOrder;
  });

  // Attach department campaigns object
  schema.campaigns = {
    retention: getActiveDepartmentCampaigns('Retention', extStr),
    tech: getActiveDepartmentCampaigns('Tech Support', extStr),
    mx: getActiveDepartmentCampaigns('Member Experience', extStr)
  };

  if (deptStr) {
    schema.activeCampaigns = getActiveDepartmentCampaigns(deptStr, extStr);
  }

  return schema;
}

// ============================================================================
// 6. HELPFUL LINKS (GOOGLE DOC INTEGRATION)
// ============================================================================

function getHelpfulLinksFromDoc() {
  const settings = getSettings();
  const docId = settings['Doc ID - Helpful Links'] || '1R0ThNGCFYMvl28rJ-PwuCaFHsLV6chOeSSpgAFvVuzU';
  const links = [];

  try {
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    const paragraphs = body.getParagraphs();

    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];
      const text = p.getText().trim();
      if (!text) continue;

      const textObj = p.editAsText();
      const str = textObj.getText();
      let lastUrl = null;

      for (let ch = 0; ch < str.length; ch++) {
        const url = textObj.getLinkUrl(ch);
        if (url && url !== lastUrl) {
          lastUrl = url;
          let name = text;
          let popup = null;

          if (name.toLowerCase().includes('chatbot') || url.toLowerCase().includes('chatbot')) {
            popup = 'chatbot';
          }

          links.push({
            name: name,
            url: url,
            popup: popup
          });
          break;
        }
      }
    }
  } catch (err) {
    console.error('Error in getHelpfulLinksFromDoc:', err);
  }

  return links;
}

// ============================================================================
// 7. STANDARD & SPECIAL EMAIL TEMPLATES
// ============================================================================

function getEmailTemplatesHtml() {
  const settings = getSettings();
  const docId = settings['Doc ID - Standard Emails'] || '1CJtJskbXT-fY2qNfna9W-i-x3rlJl404W4bgMoKFm0c';

  try {
    const url = 'https://docs.google.com/feeds/download/documents/export/Export?id=' + docId + '&exportFormat=html';
    const params = {
      method: 'get',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(url, params);
    return response.getContentText();
  } catch (err) {
    console.error('Error in getEmailTemplatesHtml:', err);
    return '';
  }
}

function getSpecialEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Special Emails');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const list = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      list.push({
        id: row[0],
        brand: row[1] || 'AAA Smart Home',
        subject: row[2] || '',
        body: row[3] || ''
      });
    }
  }
  return list;
}

// ============================================================================
// 8. FOOTPRINT & DEALER SERVICE ZIP SEARCH
// ============================================================================

function checkZipCode(zipInput) {
  const zip = String(zipInput || '').trim();
  if (!zip) {
    return { found: false, message: 'Please enter a valid zip code.' };
  }

  const settings = getSettings();
  const zipSheetId = settings['Sheet ID - Zip Codes'] || '1O7L6NVy7N4vS47IMqlGZmExdOD13l8e8kmAuXw2AAKw';

  try {
    const zipSS = SpreadsheetApp.openById(zipSheetId);

    // Rule 1: Official Pro Zips
    const proSheet = zipSS.getSheetByName('Official Pro Zips');
    if (proSheet) {
      const proData = proSheet.getDataRange().getDisplayValues();
      for (let r = 0; r < proData.length; r++) {
        for (let c = 0; c < proData[r].length; c++) {
          if (proData[r][c].trim() === zip) {
            return {
              found: true,
              status: 'PRO',
              message: '✅ In Footprint (Pro Territory)'
            };
          }
        }
      }
    }

    // Rule 2: Official DS (Upgrade & Service)
    const dsSheet = zipSS.getSheetByName('Official DS (Upgrade & Service)') ||
                    zipSS.getSheetByName('Official DS') ||
                    zipSS.getSheetByName('Dealer Service');

    if (dsSheet) {
      const dsData = dsSheet.getDataRange().getDisplayValues();
      if (dsData.length > 2) {
        const numCols = dsData[0].length;
        for (let col = 0; col < numCols; col++) {
          const headerRow1 = String(dsData[0][col] || '').toUpperCase();
          const headerRow2 = String(dsData[1][col] || '').toUpperCase();
          const isNoTech = headerRow1.includes('NO TECH') || headerRow2.includes('NO TECH');

          for (let row = 2; row < dsData.length; row++) {
            if (dsData[row][col].trim() === zip) {
              if (isNoTech) {
                return {
                  found: false,
                  status: 'NO_TECH',
                  message: '❌ No Tech Available (Dealer No Tech Territory)'
                };
              } else {
                const territoryName = dsData[0][col] || dsData[1][col] || 'Dealer Service';
                return {
                  found: true,
                  status: 'DEALER',
                  message: '✅ Dealer Service (' + territoryName.trim() + ')'
                };
              }
            }
          }
        }
      }
    }

    return {
      found: false,
      status: 'NOT_FOUND',
      message: '❌ No Tech Available'
    };

  } catch (err) {
    console.error('Error in checkZipCode:', err);
    return {
      found: false,
      status: 'ERROR',
      message: 'Error searching zip footprint: ' + err.message
    };
  }
}

// ============================================================================
// 9. FEEDBACK & CHANGE LOG
// ============================================================================

function submitFeedback(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Feedback');
    if (!sheet) {
      return { success: false, message: 'Feedback sheet could not be found.' };
    }

    const timestamp = new Date();
    const agentName = payload.name || 'Anonymous';
    const extension = payload.extension || payload.phone || '-';
    const dept = payload.dept || payload.department || '-';
    const category = payload.category || 'General';
    const form = payload.section || payload.form || payload.applicableForm || 'General';
    const comments = payload.comments || payload.details || payload.message || '';

    sheet.appendRow([
      timestamp,
      agentName,
      extension,
      dept,
      category,
      form,
      comments
    ]);

    return { success: true, message: 'Feedback submitted successfully. Thank you!' };
  } catch (err) {
    console.error('Error in submitFeedback:', err);
    return { success: false, message: 'Error submitting feedback: ' + err.message };
  }
}

function getChangeLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Change Log');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const records = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const isVisible = row[5] === undefined || row[5] === true || String(row[5]).toUpperCase() === 'TRUE';
    if (!isVisible) continue;

    let dateFormatted = row[0];
    if (row[0] instanceof Date) {
      dateFormatted = Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'MM/dd/yyyy');
    }

    records.push({
      date: dateFormatted,
      version: row[1] || 'v1.x',
      form: row[2] || 'General',
      description: row[3] || '',
      updatedBy: row[4] || 'Admin'
    });
  }

  return records.reverse();
}

function validateAdminPasscode(enteredPasscode) {
  const settings = getSettings();
  const correctPasscode = String(settings['Admin Password'] || '8454').trim();
  return String(enteredPasscode).trim() === correctPasscode;
}
