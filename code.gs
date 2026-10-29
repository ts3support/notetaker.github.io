/**
 * ============================================================================
 * AAA SMART HOME - NOTE TAKER APP BACKEND (Code.gs)
 * ============================================================================
 */

/* ============================================================================
 * 1. UTILITY & HELPER FUNCTIONS
 * ========================================================================== */

function text_(v) {
  return v == null ? '' : String(v).trim();
}

function flag_(v, fallback) {
  if (typeof v === 'boolean') return v;
  const s = text_(v).toLowerCase();
  if (['true', 'yes', '1'].includes(s)) return true;
  if (['false', 'no', '0'].includes(s)) return false;
  return fallback === undefined ? false : fallback;
}

function list_(v) {
  return text_(v).split(/[,;\n]+/).map(text_).filter(Boolean);
}

function headers_(r) {
  return r.map(v => text_(v).toLowerCase().replace(/\s+/g, ' '));
}

function cell_(r, h, n) {
  const i = h.indexOf(n);
  return i < 0 ? '' : r[i];
}

function requireHeaders_(h, requiredList, tabName) {
  const missing = requiredList.filter(v => !h.includes(v));
  if (missing.length) {
    throw new Error(tabName + ': missing headers: ' + missing.join(', '));
  }
}

function department_(v) {
  const s = text_(v).toLowerCase().replace(/\s+/g, ' ');
  if (/^(tech support|technical support)(?:$|[ -])/.test(s)) return { key: 'tech', name: 'Tech Support' };
  if (/^retention(?:$|[ -])/.test(s)) return { key: 'retention', name: 'Retention' };
  if (/^(member experience|mx)(?:$|[ -])/.test(s)) return { key: 'mx', name: 'Member Experience' };
  return null;
}

function documentId_(v) {
  const s = text_(v);
  const m = s.match(/\/d\/([A-Za-z0-9_-]+)/) || s.match(/[?&]id=([A-Za-z0-9_-]+)/);
  const id = m ? m[1] : s;
  if (!/^[A-Za-z0-9_-]+$/.test(id)) throw new Error('Invalid Google document ID or URL.');
  return id;
}

/* ============================================================================
 * 2. WEB APP ROUTING (doGet & include)
 * ========================================================================== */

function doGet(e) {
  const p = (e && e.parameter) || {};
  const page = text_(p.page).toLowerCase();

  // Default to 'Notes' so the main application and your banners load immediately.
  let file = 'Notes';
  if (page === 'login') {
    file = 'Login';
  } else if (['notes', 'main', 'app', 'index'].includes(page)) {
    file = 'Notes';
  }

  let t;
  try {
    t = HtmlService.createTemplateFromFile(file);
  } catch (err) {
    try {
      t = HtmlService.createTemplateFromFile(file.toLowerCase());
    } catch (err2) {
      t = HtmlService.createTemplateFromFile('Index');
    }
  }

  t.appUrl = ScriptApp.getService().getUrl();
  t.params = p;
  t.agent = p.agent || p.name || '';
  t.extension = p.extension || p.ext || '';
  t.department = p.department || p.dept || '';
  t.role = p.role || p.role_tier || '';

  return t
    .evaluate()
    .setTitle('AAA Smart Home - Notes App')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  const n = text_(filename);
  for (const c of [n, n.charAt(0).toUpperCase() + n.slice(1), n.toLowerCase(), n.toUpperCase()]) {
    try {
      return HtmlService.createHtmlOutputFromFile(c).getContent();
    } catch (e) {}
  }
  return '';
}

/* ============================================================================
 * 3. SETTINGS & AUTHENTICATION
 * ========================================================================== */

function getSettings() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  const out = {};
  if (s) {
    s.getDataRange().getValues().slice(1).forEach(r => {
      if (text_(r[0])) out[text_(r[0])] = r[1];
    });
  }
  return out;
}

/**
 * Returns target extensions by reading Campaigns - Tech Support and Settings.
 */
function getProjectExtensions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Pull target extensions directly from Campaigns - Tech Support
  const campSheet = ss.getSheetByName('Campaigns - Tech Support');
  if (campSheet) {
    const data = campSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === 'non_comm' || String(data[i][1]).toUpperCase().includes('NON-COMM')) {
        const targetExts = String(data[i][6] || '');
        if (targetExts) return list_(targetExts);
      }
    }
  }

  // 2. Check Settings tab key 'Project 2026 Extensions' or B3 fallback
  const s = ss.getSheetByName('Settings');
  if (s) {
    const settings = getSettings();
    if (settings['Project 2026 Extensions']) return list_(settings['Project 2026 Extensions']);
  }

  return [];
}

function isProjectExtension(ext) {
  if (!ext) return false;
  const authorized = getProjectExtensions();
  return authorized.includes(String(ext).trim()) || authorized.includes('ALL');
}

function getLoginFields() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName('login') || ss.getSheetByName('Login') || ss.getSheetByName('Index');
  return s ? s.getDataRange().getValues().slice(1)
    .filter(r => text_(r[0]))
    .map(r => ({
      id: text_(r[0]),
      internalId: text_(r[0]),
      label: text_(r[1]) || text_(r[0]),
      type: text_(r[2]) || 'Text',
      options: list_(r[3]),
      visible: flag_(r[4]),
      required: flag_(r[5])
    })) : [];
}

function getLoginConfig() {
  return getLoginFields();
}

function validateLogin(name, extension, department) {
  const valid = !!text_(name) && !!text_(extension) && !!department_(department);
  return {
    success: valid,
    agentName: text_(name),
    extension: text_(extension),
    department: text_(department),
    message: valid ? '' : 'Enter an agent name, extension, and recognized department.'
  };
}

function validateAdminPasscode(value) {
  const correct = text_(getSettings()['Admin Password']);
  return correct !== '' && text_(value) === correct;
}

function verifyAdminPasscode(value) {
  return validateAdminPasscode(value);
}

/* ============================================================================
 * 4. CAMPAIGNS & SCHEMA INGESTION
 * ========================================================================== */

function getActiveDepartmentCampaigns(department, extension) {
  const d = department_(department);
  if (!d) return [];

  const tab = 'Campaigns - ' + d.name;
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tab);
  if (!s) return [];

  const data = s.getDataRange().getValues();
  if (data.length < 2) return [];

  const h = headers_(data[0]);
  requireHeaders_(h, ['campaign id', 'display label', 'field type', 'options', 'active', 'target scope', 'target extensions'], tab);

  const out = [];
  const seen = new Set();
  const ext = text_(extension);

  // Column J auto-populate reason header check
  const reasonHeader = h.find(col => col.includes('auto') && col.includes('reason')) || 
                       (h.includes('auto-populate reason') ? 'auto-populate reason' : 'auto populate reason');

  data.slice(1).forEach(r => {
    if (!flag_(cell_(r, h, 'active'))) return;
    const id = text_(cell_(r, h, 'campaign id'));
    if (!id) return;

    const scope = text_(cell_(r, h, 'target scope')).toLowerCase().replace(/[\s_]+/g, '-');
    const targets = list_(cell_(r, h, 'target extensions'));

    // Check scope targeting
    if (!(['department-wide', 'global'].includes(scope) || 
          targets.some(v => v.toUpperCase() === 'ALL') || 
          (scope === 'extension-specific' && ext !== '' && targets.includes(ext)))) {
      return;
    }

    if (seen.has(id)) throw new Error(tab + ': duplicate active campaign ID ' + id);
    seen.add(id);

    const options = list_(cell_(r, h, 'options'));

    out.push({
      id: id,
      campaignId: id,
      internalId: id,
      label: text_(cell_(r, h, 'display label')) || id,
      type: text_(cell_(r, h, 'field type')) || 'Radio',
      options: options.length ? options : ['Yes', 'No'],
      required: flag_(cell_(r, h, 'required')),
      targetScope: scope,
      notes: text_(cell_(r, h, 'notes')),
      autoPopulateReason: text_(cell_(r, h, reasonHeader)) // Column J
    });
  });

  return out;
}

function getCampaigns(department, extension) {
  return getActiveDepartmentCampaigns(department, extension);
}

function getFormSchema(department, extension) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ext = text_(extension);
  const d = department_(department);

  const campaigns = {
    tech: getActiveDepartmentCampaigns('Tech Support', ext),
    retention: getActiveDepartmentCampaigns('Retention', ext),
    mx: getActiveDepartmentCampaigns('Member Experience', ext)
  };

  const schema = {
    campaigns: campaigns,
    activeCampaigns: d ? campaigns[d.key] : [],
    authorizedExtensions: getProjectExtensions(),
    warnings: []
  };

  const specs = [
    ['login', 'login', 'tech'],
    ['index', 'login', 'tech'],
    ['inbound', 'Inbound', 'tech'],
    ['outbound', 'Outbound', 'tech'],
    ['retention', 'Retention', 'retention'],
    ['retention_outbound', 'Retention Outbound', 'retention'],
    ['residential', 'Residential', 'tech'],
    ['dealer', 'Dealer', 'tech'],
    ['addendum', 'Addendum', null],
    ['parts', 'Parts', 'tech'],
    ['mx', 'MX', 'mx'],
    ['mx_outbound', 'MX Outbound', 'mx']
  ];
  const cache = {};

  specs.forEach(spec => {
    const key = spec[0];
    const tab = spec[1];
    schema[key] = {};
    schema[key + '_order'] = [];

    if (!Object.prototype.hasOwnProperty.call(cache, tab)) {
      const s = ss.getSheetByName(tab) || (tab === 'login' ? (ss.getSheetByName('Login') || ss.getSheetByName('Index')) : null);
      cache[tab] = s ? s.getDataRange().getValues() : [];
    }

    const data = cache[tab];
    if (data.length < 2) return;

    const h = headers_(data[0]);
    requireHeaders_(h, ['internal id', 'display label', 'field type', 'options', 'visible', 'required'], tab);

    const owner = spec[2] || (d ? d.key : null);
    const available = owner ? campaigns[owner] : [];

    data.slice(1).forEach(r => {
      const source = text_(cell_(r, h, 'internal id'));
      if (!source) return;

      const options = text_(cell_(r, h, 'options'));
      const dynamic = source === 'campaign_toggle' || /dynamic:campaigns/i.test(options);
      const id = dynamic ? 'campaign_toggle' : source;

      const f = {
        id: id,
        internalId: id,
        label: text_(cell_(r, h, 'display label')) || source,
        type: text_(cell_(r, h, 'field type')) || 'Text',
        options: options,
        visible: flag_(cell_(r, h, 'visible')),
        required: flag_(cell_(r, h, 'required')),
        dependsOn: text_(cell_(r, h, 'depends on')),
        copyIfBlank: flag_(cell_(r, h, 'copy if blank'))
      };

      if (dynamic) {
        f.isCampaign = true;
        f.placement = 'top_controls';
        f.sourceId = source;
        f.campaignId = '';
        f.autoPopulateReason = '';

        if (available.length) {
          const c = available[0];
          f.campaignId = c.campaignId;
          f.label = c.label;
          f.type = c.type;
          f.options = c.options.join(',');
          f.required = c.required;
          f.autoPopulateReason = c.autoPopulateReason || '';
          if (available.length > 1) {
            schema.warnings.push(tab + ': multiple campaigns matched; using first row: ' + c.campaignId);
          }
        } else {
          f.visible = false;
          f.required = false;
          f.options = '';
        }
      }

      if (schema[key][id]) {
        throw new Error(tab + ': duplicate resolved Internal ID ' + id);
      }
      schema[key][id] = f;
      schema[key + '_order'].push(f);
    });
  });

  return schema;
}

/* ============================================================================
 * 5. HELPFUL LINKS & DOC IMPORTS
 * ========================================================================== */

function helpfulLinksSheet_(department) {
  const d = department_(department);
  if (!d) throw new Error('Helpful Links: unrecognized department.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = ['Links - ' + d.name, 'Links-' + d.name];
  if (d.key === 'mx') names.push('Links - MX', 'Links-MX');

  for (const name of names) {
    const s = ss.getSheetByName(name);
    if (s) return s;
  }
  throw new Error('Worksheet not found: ' + names[0]);
}

function getHelpfulLinks(department) {
  const d = department_(department);
  const s = helpfulLinksSheet_(department);
  const data = s.getDataRange().getDisplayValues();
  if (!data.length || data.every(r => r.every(v => !text_(v)))) return [];

  const norm = v => text_(v).toLowerCase().replace(/[_-]+/g, ' ').replace(/\s+/g, ' ');
  const h = data[0].map(norm);
  const find = names => {
    for (const n of names) {
      const i = h.indexOf(norm(n));
      if (i >= 0) return i;
    }
    return -1;
  };

  const c = {
    id: find(['link id', 'resource id']),
    title: find(['link title', 'title', 'name']),
    url: find(['url', 'link url']),
    active: find(['active']),
    category: find(['category']),
    description: find(['description']),
    instructions: find(['agent instructions', 'instructions']),
    popup: find(['popup mode']),
    sort: find(['sort order'])
  };

  const missing = [];
  if (c.title < 0) missing.push('Link Title');
  if (c.url < 0) missing.push('URL');
  if (c.active < 0) missing.push('Active');
  if (missing.length) throw new Error(s.getName() + ': missing headers: ' + missing.join(', '));

  const v = (r, i) => (i < 0 ? '' : text_(r[i]));
  const out = [];
  const seen = new Set();

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!flag_(v(r, c.active))) continue;

    const title = v(r, c.title);
    const url = v(r, c.url);
    if (!title || !/^https?:\/\/[^\s/]+(?:[/?#][^\s]*)?$/i.test(url)) {
      console.warn(s.getName() + ': skipped invalid active row ' + (i + 1));
      continue;
    }

    const id = v(r, c.id) || d.key + '_row_' + (i + 1);
    if (seen.has(id)) throw new Error(s.getName() + ': duplicate Link ID ' + id);
    seen.add(id);

    const popup = ['popup', 'modal', 'chatbot'].includes(v(r, c.popup).toLowerCase());
    const raw = v(r, c.sort);
    const number = raw === '' ? NaN : Number(raw);
    const instructions = v(r, c.instructions);

    out.push({
      id: id,
      name: title,
      title: title,
      url: url,
      category: v(r, c.category) || 'General',
      description: v(r, c.description),
      instructions: instructions,
      agentInstructions: instructions,
      department: d.name,
      popupMode: popup ? 'popup' : 'new_tab',
      popup: popup ? 'chatbot' : null,
      sortOrder: Number.isFinite(number) ? number : i + 1,
      sourceRow: i + 1,
      active: true
    });
  }

  out.sort((a, b) => a.sortOrder - b.sortOrder || a.sourceRow - b.sourceRow);
  return out;
}

function getQuickLinks(department) {
  return getHelpfulLinks(department);
}

function getHelpfulLinksFromDoc(department) {
  if (text_(department)) return getHelpfulLinks(department);
  const id = text_(getSettings()['Doc ID - Helpful Links']);
  if (!id) return [];
  return extractDocLinks_(DocumentApp.openById(documentId_(id)).getBody()).map(l => ({
    name: l.title,
    title: l.title,
    url: l.url,
    category: 'General',
    popup: /chatbot/i.test(l.title + ' ' + l.url) ? 'chatbot' : null
  }));
}

function extractDocLinks_(body) {
  const out = [];
  function visit(e) {
    if (e.getType() === DocumentApp.ElementType.TEXT) {
      const t = e.asText();
      const content = t.getText();
      const starts = t.getTextAttributeIndices();
      let url = '';
      let title = '';

      function flush() {
        if (/^https?:\/\//i.test(url)) {
          out.push({ title: text_(title) || url, url: url });
        }
        url = '';
        title = '';
      }

      starts.forEach((start, i) => {
        const end = i + 1 < starts.length ? starts[i + 1] : content.length;
        const next = text_(t.getLinkUrl(start));
        if (next !== url) {
          flush();
          url = next;
        }
        if (next) title += content.substring(start, end);
      });
      flush();
      return;
    }
    if (typeof e.getNumChildren === 'function') {
      for (let i = 0; i < e.getNumChildren(); i++) visit(e.getChild(i));
    }
  }
  visit(body);
  return out;
}

function importLinksFromGoogleDoc() {
  return importDepartmentHelpfulLinks_('Tech Support');
}
function importTechSupportLinksFromGoogleDoc() {
  return importDepartmentHelpfulLinks_('Tech Support');
}
function importRetentionLinksFromGoogleDoc() {
  return importDepartmentHelpfulLinks_('Retention');
}
function importMemberExperienceLinksFromGoogleDoc() {
  return importDepartmentHelpfulLinks_('Member Experience');
}

function importDepartmentHelpfulLinks_(department) {
  const d = department_(department);
  if (!d) throw new Error('Unrecognized import department.');

  const settings = getSettings();
  const setting = 'Doc ID - Helpful Links - ' + d.name;
  const configured = text_(settings[setting]) || (d.key === 'tech' ? '1R0ThNGCFYMvl28rJ-PwuCaFHsLV6chOeSSpgAFvVuzU' : '');
  if (!configured) throw new Error('Missing Settings value: ' + setting);

  const id = documentId_(configured);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const s = helpfulLinksSheet_(d.name);
    const width = s.getLastColumn();
    if (!width) throw new Error(s.getName() + ' needs a header row.');

    const h = headers_(s.getRange(1, 1, 1, width).getDisplayValues()[0]);
    requireHeaders_(h, ['link title', 'url', 'active'], s.getName());

    const seen = new Set();
    const last = s.getLastRow();
    if (last > 1) {
      s.getRange(2, h.indexOf('url') + 1, last - 1, 1).getDisplayValues().forEach(r => {
        if (text_(r[0])) seen.add(text_(r[0]));
      });
    }

    const extracted = extractDocLinks_(DocumentApp.openById(id).getBody());
    const rows = [];
    let duplicates = 0;

    extracted.forEach(link => {
      if (seen.has(link.url)) {
        duplicates++;
        return;
      }
      seen.add(link.url);

      const row = new Array(width).fill('');
      const put = (name, value) => {
        const i = h.indexOf(name);
        if (i >= 0) row[i] = value;
      };

      put('link title', /^[=+@-]/.test(link.title) ? "'" + link.title : link.title);
      put('url', link.url);
      put('category', 'Guides & Reference');
      put('popup mode', 'New Tab');
      put('active', true);
      put('notes', 'Imported from ' + d.name + ' Google Doc');
      put('link id', Utilities.getUuid());
      rows.push(row);
    });

    if (rows.length) {
      const first = s.getLastRow() + 1;
      const needed = first + rows.length - 1;
      if (needed > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), needed - s.getMaxRows());
      s.getRange(first, 1, rows.length, width).setValues(rows);
    }

    const result = {
      department: d.name,
      sourceDocumentId: id,
      destination: s.getName(),
      imported: rows.length,
      duplicatesSkipped: duplicates
    };
    console.log(JSON.stringify(result));
    ss.toast(d.name + ': imported ' + rows.length + ' new links; skipped ' + duplicates + ' duplicates.', 'Helpful Links Import', 8);
    return result;
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================================
 * 6. EMAIL TEMPLATES (Standard & Special Emails)
 * ========================================================================== */

function getEmailTemplatesHtml(department) {
  const settings = getSettings();
  const d = department_(department);
  if (text_(department) && !d) throw new Error('Unrecognized email department.');

  const source = (d ? text_(settings['Doc ID - Standard Emails - ' + d.name]) : '') || text_(settings['Doc ID - Standard Emails']);
  if (!source) throw new Error('No standard email document configured for ' + (d ? d.name : 'the shared library') + '.');

  const response = UrlFetchApp.fetch('https://docs.google.com/document/d/' + encodeURIComponent(documentId_(source)) + '/export?format=html', {
    method: 'get',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  if (status < 200 || status >= 300) throw new Error('Email export failed (HTTP ' + status + '). Check document ID and permissions.');
  return response.getContentText();
}

/**
 * Universal getter for Special Emails supporting multiple call signatures:
 * getSpecialEmails(department, campaignId) OR getSpecialEmails(extension)
 */
function getSpecialEmails(departmentOrExtension, campaignId) {
  if (text_(departmentOrExtension) && /^\d+$/.test(text_(departmentOrExtension))) {
    return readSpecialEmails_('', campaignId, departmentOrExtension);
  }
  return readSpecialEmails_(departmentOrExtension, campaignId);
}

function getSpecialEmailTemplates(extension, department, campaignId) {
  return readSpecialEmails_(department, campaignId, extension);
}

/**
 * Reads from the updated 7-column 'Special Emails' sheet:

 * A: Template ID | B: Department | C: Campaign ID | D: Brand | E: Subject | F: Template Body | G: Active
 */
function readSpecialEmails_(department, campaignId, extension) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Special Emails');
  if (!s) return [];

  const data = s.getDataRange().getValues();
  if (data.length < 2) return [];

  const h = headers_(data[0]);
  const rows = data.slice(1);

  const targetDept = department ? text_(department).toLowerCase() : '';
  const targetCamp = campaignId ? text_(campaignId).toLowerCase() : '';
  const ext = text_(extension);

  // If extension is passed, verify against targeted extensions if non-empty
  if (ext && !isProjectExtension(ext)) {
    return [];
  }

  return rows
    .filter(r => {
      const id = text_(cell_(r, h, 'template id'));
      if (!id) return false;

      // 1. Must be Active (Column G)
      const isActive = flag_(cell_(r, h, 'active'), true);
      if (!isActive) return false;

      // 2. Department check (Column B) — matches dept, 'global', 'all', or blank
      if (targetDept) {
        const rowDept = text_(cell_(r, h, 'department')).toLowerCase();
        if (rowDept && !['all', 'global', targetDept].includes(rowDept)) {
          return false;
        }
      }

      // 3. Campaign ID check (Column C)
      if (targetCamp) {
        const rowCamp = text_(cell_(r, h, 'campaign id')).toLowerCase();
        if (rowCamp && rowCamp !== targetCamp) {
          return false;
        }
      }

      return true;
    })
    .map(r => ({
      id: text_(cell_(r, h, 'template id')),                         // Col A
      department: text_(cell_(r, h, 'department')),                 // Col B
      campaignId: text_(cell_(r, h, 'campaign id')),                 // Col C
      brand: text_(cell_(r, h, 'brand')) || 'AAA Smart Home',       // Col D
      subject: text_(cell_(r, h, 'subject')),                       // Col E
      body: text_(cell_(r, h, 'template body')),                     // Col F
      active: flag_(cell_(r, h, 'active'), true),                   // Col G
      type: 'special'
    }));
}

/* ============================================================================
 * 7. ZIP LOOKUP, FEEDBACK & CHANGELOG
 * ========================================================================== */

function checkZipCode(input) {
  const raw = text_(input);
  if (!/^\d{1,5}$/.test(raw)) return { found: false, status: 'INVALID', message: 'Invalid ZIP format.' };
  const zip = raw.padStart(5, '0');

  try {
    const settings = getSettings();
    const ss = SpreadsheetApp.openById(settings['Sheet ID - Zip Codes'] || '1O7L6NVy7N4vS47IMqlGZmExdOD13l8e8kmAuXw2AAKw');

    // 1. Check Official Pro Zips
    const pro = ss.getSheetByName('Official Pro Zips');
    if (pro && pro.getDataRange().getDisplayValues().some(r => r.some(v => v.trim() === zip))) {
      return { found: true, status: 'PRO', message: '✅ In Footprint (Pro Territory)' };
    }

    // 2. Check Official DS (Upgrade & Service)
    const ds = ss.getSheetByName('Official DS (Upgrade & Service)') || ss.getSheetByName('Official DS') || ss.getSheetByName('Dealer Service');
    if (ds) {
      const data = ds.getDataRange().getDisplayValues();
      if (data.length > 2) {
        for (let c = 0; c < data[0].length; c++) {
          const noTech = /NO TECH/i.test(data[0][c] + ' ' + data[1][c]);
          for (let r = 2; r < data.length; r++) {
            if (data[r][c].trim() !== zip) continue;
            if (noTech) return { found: false, status: 'NO_TECH', message: '❌ No Tech Available (Dealer No Tech Territory)' };
            return { found: true, status: 'DEALER', message: '✅ Dealer Service (' + text_(data[0][c] || data[1][c] || 'Dealer Service') + ')' };
          }
        }
      }
    }

    return { found: false, status: 'NOT_FOUND', message: '❌ No Tech Available' };
  } catch (e) {
    console.error(e);
    return { found: false, status: 'ERROR', message: 'Error searching zip footprint: ' + e.message };
  }
}

function submitFeedback(payload) {
  try {
    const p = payload || {};
    const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Feedback');
    if (!s) return { success: false, message: 'Feedback sheet could not be found.' };

    const comments = text_(p.comments || p.details || p.message);
    if (!comments) return { success: false, message: 'Please enter comments or details.' };

    const safe = v => {
      const t = text_(v);
      return /^[=+@-]/.test(t) ? "'" + t : t;
    };

    s.appendRow([
      new Date(),
      safe(p.name || p.agentName || 'Anonymous'),
      safe(p.extension || p.phone || '-'),
      safe(p.dept || p.department || '-'),
      safe(p.category || 'General'),
      safe(p.section || p.form || p.applicableForm || 'General'),
      safe(comments)
    ]);
    return { success: true, message: 'Feedback submitted successfully. Thank you!' };
  } catch (e) {
    return { success: false, message: 'Error submitting feedback: ' + e.message };
  }
}

function getChangeLog() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Change Log');
  return s ? s.getDataRange().getValues().slice(1)
    .filter(r => flag_(r[5], true))
    .map(r => ({
      date: r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'MM/dd/yyyy') : text_(r[0]),
      version: text_(r[1]) || 'v1.x',
      form: text_(r[2]) || 'General',
      applicableForm: text_(r[2]) || 'General',
      description: text_(r[3]),
      updatedBy: text_(r[4]) || 'Admin'
    }))
    .reverse() : [];
}

function parseOptions(value) {
  return Array.isArray(value) ? value : list_(value);
}

function parseBoolean(value, fallback = false) {
  return flag_(value, fallback);
}
