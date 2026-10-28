/**
 * Note Taker backend: department campaigns, links, email templates and forms.
 * Campaign control ID remains campaign_toggle; campaignId is separate metadata.
 * Department resources do not silently fall back to another department.
 */
function doGet(e) {
  const p = e && e.parameter || {};
  const page = String(p.page || '').trim().toLowerCase();
  let file = 'Login';
  if (['notes','main','app','index'].includes(page)) file = 'Notes';
  else if (page !== 'login' && (p.agent || p.extension || p.ext || p.user)) file = 'Notes';
  let t;
  try { t = HtmlService.createTemplateFromFile(file); }
  catch (err) { t = HtmlService.createTemplateFromFile(file.toLowerCase()); }
  t.appUrl = ScriptApp.getService().getUrl();
  t.params = p;
  t.agent = p.agent || p.name || '';
  t.extension = p.extension || p.ext || '';
  t.department = p.department || p.dept || '';
  t.role = p.role || p.role_tier || '';
  return t.evaluate().setTitle('AAA Smart Home - Note Taker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

function include(filename) {
  const name = String(filename || '');
  for (const candidate of [name, name.charAt(0).toUpperCase() + name.slice(1), name.toLowerCase(), name.toUpperCase()]) {
    try { return HtmlService.createHtmlOutputFromFile(candidate).getContent(); } catch (err) {}
  }
  return '';
}

function text_(v) { return v === null || v === undefined ? '' : String(v).trim(); }
function flag_(v, fallback) {
  if (v === true || v === false) return v;
  const s = text_(v).toLowerCase();
  if (['true','yes','1'].includes(s)) return true;
  if (['false','no','0'].includes(s)) return false;
  return fallback === undefined ? false : fallback;
}
function list_(v) { return text_(v).split(/[,;\n]+/).map(text_).filter(Boolean); }
function department_(value) {
  const s = text_(value).toLowerCase().replace(/\s+/g,' ');
  if (/^(tech support|technical support)(?:$|[ -])/.test(s)) return {key:'tech',name:'Tech Support'};
  if (/^(member experience|mx)(?:$|[ -])/.test(s)) return {key:'mx',name:'Member Experience'};
  if (/^retention(?:$|[ -])/.test(s)) return {key:'retention',name:'Retention'};
  return null;
}
function headers_(row) { return row.map(v => text_(v).toLowerCase().replace(/\s+/g,' ')); }
function cell_(row, headers, name) {
  const i = headers.indexOf(name);
  return i < 0 ? '' : row[i];
}
function requireHeaders_(headers, names, tab) {
  const missing = names.filter(n => !headers.includes(n));
  if (missing.length) throw new Error(tab + ': missing headers: ' + missing.join(', '));
}
function getSettings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  const result = {};
  if (!sheet) return result;
  sheet.getDataRange().getValues().slice(1).forEach(row => {
    const key = text_(row[0]);
    if (key) result[key] = row[1];
  });
  return result;
}

function getActiveDepartmentCampaigns(department, extension) {
  const dept = department_(department);
  if (!dept) return [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tab = 'Campaigns - ' + dept.name;
  const sheet = ss.getSheetByName(tab);
  if (!sheet) return [];
  return readCampaigns_(sheet.getDataRange().getValues(), tab, extension);
}
function readCampaigns_(data, tab, extension) {
  if (data.length < 2) return [];
  const h = headers_(data[0]);
  requireHeaders_(h,['campaign id','display label','field type','options','active','target scope','target extensions'],tab);
  const ext = text_(extension);
  const found = new Set();
  const result = [];
  data.slice(1).forEach(row => {
    if (!flag_(cell_(row,h,'active'))) return;
    const id = text_(cell_(row,h,'campaign id'));
    if (!id) return;
    const scope = text_(cell_(row,h,'target scope')).toLowerCase().replace(/[\s_]+/g,'-');
    const targets = list_(cell_(row,h,'target extensions'));
    const targeted = scope === 'department-wide' || targets.some(v => v.toUpperCase() === 'ALL') ||
      (scope === 'extension-specific' && ext !== '' && targets.includes(ext));
    if (!targeted) return;
    if (found.has(id)) throw new Error(tab + ': duplicate active campaign ID: ' + id);
    found.add(id);
    result.push({
      id:id, campaignId:id,
      label:text_(cell_(row,h,'display label')) || id,
      type:text_(cell_(row,h,'field type')) || 'Radio',
      options:list_(cell_(row,h,'options')).length ? list_(cell_(row,h,'options')) : ['Yes','No'],
      required:flag_(cell_(row,h,'required')),
      targetScope:scope,
      notes:text_(cell_(row,h,'notes'))
    });
  });
  return result;
}

function getLoginFields() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('login') || ss.getSheetByName('Login') || ss.getSheetByName('Index');
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1).filter(row => text_(row[0])).map(row => ({
    id:text_(row[0]), internalId:text_(row[0]), label:text_(row[1]) || text_(row[0]),
    type:text_(row[2]) || 'Text', options:list_(row[3]), visible:flag_(row[4]), required:flag_(row[5])
  }));
}
function getLoginConfig() { return getLoginFields(); }
function validateLogin(agentName, extension, department) {
  // Compatibility response only; this is not identity authentication.
  const valid = !!text_(agentName) && !!text_(extension) && !!department_(department);
  return {success:valid,agentName:text_(agentName),extension:text_(extension),department:text_(department),
    message:valid ? '' : 'Enter an agent name, extension, and recognized department.'};
}
function getProjectExtensions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  // Preserve the original Settings!B3 extension list used for special-email access.
  return sheet ? list_(sheet.getRange('B3').getDisplayValue()) : [];
}
function isProjectExtension(ext) { return getProjectExtensions().includes(text_(ext)); }

function getFormSchema(department, extension) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ext = text_(extension);
  const agentDept = department_(department);
  const campaigns = {
    tech:getActiveDepartmentCampaigns('Tech Support',ext),
    retention:getActiveDepartmentCampaigns('Retention',ext),
    mx:getActiveDepartmentCampaigns('Member Experience',ext)
  };
  const schema = {campaigns:campaigns,activeCampaigns:agentDept ? campaigns[agentDept.key] : [],
    authorizedExtensions:getProjectExtensions(),warnings:[]};
  const forms = [
    ['login','login','tech'],['index','login','tech'],
    ['inbound','Inbound','tech'],['outbound','Outbound','tech'],
    ['retention','Retention','retention'],['retention_outbound','Retention Outbound','retention'],
    ['residential','Residential','tech'],['dealer','Dealer','tech'],['addendum','Addendum',null],
    ['parts','Parts','tech'],['mx','MX','mx'],['mx_outbound','MX Outbound','mx']
  ];
  const sheetData = {};
  forms.forEach(spec => {
    const key = spec[0], tab = spec[1];
    schema[key] = {};
    schema[key + '_order'] = [];
    if (!Object.prototype.hasOwnProperty.call(sheetData,tab)) {
      const sheet = ss.getSheetByName(tab) || (tab === 'login' ? ss.getSheetByName('Login') || ss.getSheetByName('Index') : null);
      sheetData[tab] = sheet ? sheet.getDataRange().getValues() : [];
    }
    const data = sheetData[tab];
    if (data.length < 2) return;
    const h = headers_(data[0]);
    requireHeaders_(h,['internal id','display label','field type','options','visible','required'],tab);
    // Each form uses its owning department; shared addendum uses the signed-in department.
    const owner = spec[2] || (agentDept ? agentDept.key : null);
    const available = owner ? campaigns[owner] : [];
    data.slice(1).forEach(row => {
      const sourceId = text_(cell_(row,h,'internal id'));
      if (!sourceId) return;
      const rawOptions = text_(cell_(row,h,'options'));
      const dynamic = sourceId === 'campaign_toggle' || /dynamic:campaigns/i.test(rawOptions);
      const field = {
        id:dynamic ? 'campaign_toggle' : sourceId,
        internalId:dynamic ? 'campaign_toggle' : sourceId,
        label:text_(cell_(row,h,'display label')) || sourceId,
        type:text_(cell_(row,h,'field type')) || 'Text',
        options:rawOptions,visible:flag_(cell_(row,h,'visible')),required:flag_(cell_(row,h,'required')),
        dependsOn:text_(cell_(row,h,'depends on')),copyIfBlank:flag_(cell_(row,h,'copy if blank'))
      };
      if (dynamic) {
        field.isCampaign = true;
        field.placement = 'top_controls';
        field.sourceId = sourceId;
        field.campaignId = '';
        if (available.length) {
          const campaign = available[0];
          field.campaignId = campaign.campaignId;
          field.label = campaign.label;
          field.type = campaign.type;
          field.options = campaign.options.join(',');
          // Keep the form row's Visible switch authoritative.
          field.required = campaign.required;
          if (available.length > 1) schema.warnings.push(tab + ': multiple campaigns match; using the first row: ' + campaign.campaignId);
        } else {
          field.visible = false;
          field.required = false;
          field.options = '';
        }
      }
      if (schema[key][field.id]) throw new Error(tab + ': duplicate resolved Internal ID: ' + field.id);
      schema[key][field.id] = field;
      schema[key + '_order'].push(field);
    });
  });
  return schema;
}

function getHelpfulLinks(department) {
  const dept = department_(department);
  if (!dept) return [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tab = 'Links - ' + dept.name;
  const sheet = ss.getSheetByName(tab);
  if (!sheet) {
    // Optional explicit shared document; no fixed or other-department fallback.
    const shared = text_(getSettings()['Doc ID - Helpful Links']);
    return shared ? readHelpfulLinksDoc_(shared) : [];
  }
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = headers_(data[0]);
  requireHeaders_(h,['link title','url','active'],tab);
  const links = [];
  data.slice(1).forEach(row => {
    if (!flag_(cell_(row,h,'active'))) return;
    const name = text_(cell_(row,h,'link title'));
    const url = text_(cell_(row,h,'url'));
    if (!name || !/^https?:\/\//i.test(url)) return;
    const popup = text_(cell_(row,h,'popup mode')).toLowerCase();
    links.push({name:name,url:url,category:text_(cell_(row,h,'category')) || 'General',
      popup:['popup','modal','chatbot'].includes(popup) || /chatbot/i.test(name) ? 'chatbot' : null});
  });
  return links;
}
function getHelpfulLinksFromDoc(department) {
  // Compatibility wrapper: old callers can supply a department to use the new routing.
  if (text_(department)) return getHelpfulLinks(department);
  const doc = text_(getSettings()['Doc ID - Helpful Links']);
  return doc ? readHelpfulLinksDoc_(doc) : [];
}
function documentId_(value) {
  const s = text_(value);
  const match = s.match(/\/d\/([A-Za-z0-9_-]+)/) || s.match(/[?&]id=([A-Za-z0-9_-]+)/);
  const id = match ? match[1] : s;
  if (!/^[A-Za-z0-9_-]+$/.test(id)) throw new Error('Invalid Google document ID or URL.');
  return id;
}
function readHelpfulLinksDoc_(value) {
  const doc = DocumentApp.openById(documentId_(value));
  const links = [];
  doc.getBody().getParagraphs().forEach(p => {
    const name = p.getText().trim();
    if (!name) return;
    const text = p.editAsText();
    for (let i = 0; i < text.getText().length; i++) {
      const url = text.getLinkUrl(i);
      if (url && /^https?:\/\//i.test(url)) {
        links.push({name:name,url:url,category:'General',popup:/chatbot/i.test(name + ' ' + url) ? 'chatbot' : null});
        break;
      }
    }
  });
  return links;
}

function getEmailTemplatesHtml(department) {
  const settings = getSettings();
  const dept = department_(department);
  if (text_(department) && !dept) throw new Error('Unrecognized email-template department.');
  const specific = dept ? text_(settings['Doc ID - Standard Emails - ' + dept.name]) : '';
  const configured = specific || text_(settings['Doc ID - Standard Emails']);
  if (!configured) throw new Error('No standard email document configured for ' + (dept ? dept.name : 'the shared email library') + '.');
  const id = documentId_(configured);
  const url = 'https://docs.google.com/document/d/' + encodeURIComponent(id) + '/export?format=html';
  const response = UrlFetchApp.fetch(url,{
    method:'get',headers:{Authorization:'Bearer ' + ScriptApp.getOAuthToken()},muteHttpExceptions:true
  });
  const status = response.getResponseCode();
  if (status < 200 || status >= 300) throw new Error('Email document export failed (HTTP ' + status + '). Check the document ID and access permissions.');
  return response.getContentText();
}

function getSpecialEmailTemplates(extension) {
  if (!isProjectExtension(extension)) return [];
  return readSpecialEmails_();
}
function getSpecialEmails(extension) {
  // Keep the legacy public name, but require the same extension check.
  return getSpecialEmailTemplates(extension);
}
function readSpecialEmails_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Special Emails');
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1).filter(row => text_(row[0])).map(row => ({
    id:text_(row[0]),brand:text_(row[1]) || 'AAA Smart Home',subject:text_(row[2]),body:text_(row[3])
  }));
}

function checkZipCode(zipInput) {
  const zip = text_(zipInput);
  if (!zip) return {found:false,message:'Please enter a valid zip code.'};
  const settings = getSettings();
  const sheetId = settings['Sheet ID - Zip Codes'] || '1O7L6NVy7N4vS47IMqlGZmExdOD13l8e8kmAuXw2AAKw';
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const pro = ss.getSheetByName('Official Pro Zips');
    if (pro) {
      const rows = pro.getDataRange().getDisplayValues();
      if (rows.some(row => row.some(v => v.trim() === zip))) return {found:true,status:'PRO',message:'✅ In Footprint (Pro Territory)'};
    }
    const ds = ss.getSheetByName('Official DS (Upgrade & Service)') || ss.getSheetByName('Official DS') || ss.getSheetByName('Dealer Service');
    if (ds) {
      const data = ds.getDataRange().getDisplayValues();
      if (data.length > 2) {
        for (let col = 0; col < data[0].length; col++) {
          const noTech = /NO TECH/i.test(data[0][col] + ' ' + data[1][col]);
          for (let row = 2; row < data.length; row++) {
            if (data[row][col].trim() !== zip) continue;
            if (noTech) return {found:false,status:'NO_TECH',message:'❌ No Tech Available (Dealer No Tech Territory)'};
            const territory = data[0][col] || data[1][col] || 'Dealer Service';
            return {found:true,status:'DEALER',message:'✅ Dealer Service (' + territory.trim() + ')'};
          }
        }
      }
    }
    return {found:false,status:'NOT_FOUND',message:'❌ No Tech Available'};
  } catch (err) {
    console.error('Error in checkZipCode:',err);
    return {found:false,status:'ERROR',message:'Error searching zip footprint: ' + err.message};
  }
}

function submitFeedback(payload) {
  try {
    const p = payload || {};
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Feedback');
    if (!sheet) return {success:false,message:'Feedback sheet could not be found.'};
    const comments = text_(p.comments || p.details || p.message);
    if (!comments) return {success:false,message:'Please enter comments or details.'};
    const safe = v => {
      const s = text_(v);
      return /^[=+@-]/.test(s) ? "'" + s : s;
    };
    sheet.appendRow([new Date(),safe(p.name || 'Anonymous'),safe(p.extension || p.phone || '-'),
      safe(p.dept || p.department || '-'),safe(p.category || 'General'),
      safe(p.section || p.form || p.applicableForm || 'General'),safe(comments)]);
    return {success:true,message:'Feedback submitted successfully. Thank you!'};
  } catch (err) { return {success:false,message:'Error submitting feedback: ' + err.message}; }
}
function getChangeLog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Change Log');
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1).filter(row => flag_(row[5],true)).map(row => ({
    date:row[0] instanceof Date ? Utilities.formatDate(row[0],Session.getScriptTimeZone(),'MM/dd/yyyy') : text_(row[0]),
    version:text_(row[1]) || 'v1.x',form:text_(row[2]) || 'General',description:text_(row[3]),updatedBy:text_(row[4]) || 'Admin'
  })).reverse();
}
function validateAdminPasscode(enteredPasscode) {
  const configured = text_(getSettings()['Admin Password']);
  return configured !== '' && text_(enteredPasscode) === configured;
}
