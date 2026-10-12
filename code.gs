// ==========================================
// WEB APP ROUTING LOGIC
// ==========================================
function doGet(e) {
  const pageName = (e.parameter.page || '').toLowerCase();  
  const pages = {
    notes: 'Notes',
    residential: 'Residential',
    dealer: 'Dealer',
    addendum: 'Addendum'
  };  
  
  const page = pages[pageName] || 'Index';  
  
  const template = HtmlService.createTemplateFromFile(page);
  template.appUrl = ScriptApp.getService().getUrl();  
  
  return template
    .evaluate()
    .setTitle('Note Taker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// EMAIL TEMPLATES LOGIC
// ==========================================
function getEmailTemplatesHtml() {
  const docId = '1CJtJskbXT-fY2qNfna9W-i-x3rlJl404W4bgMoKFm0c'; 
  DriveApp.getFileById(docId); 
  
  const url = "https://docs.google.com/feeds/download/documents/export/Export?id=" + docId + "&exportFormat=html";
  const options = {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  return response.getContentText();
}

// ==========================================
// DYNAMIC HELPFUL LINKS LOGIC
// ==========================================
function getHelpfulLinksFromDoc() {
  const docId = '1R0ThNGCFYMvl28rJ-PwuCaFHsLV6chOeSSpgAFvVuzU';
  const doc = DocumentApp.openById(docId);
  
  const top10Names = [
    "Tech Support Opportunities",
    "Chatbot",
    "Common Parts List",
    "Test Mode SOP",
    "DS Zip Codes",
    "Tech Support Tier I Guide ",
    "Known Issues Tracker",
    "Waived Fee Guideline",
    "Call Routing Guide",
    "Actual/Incident Form"      
  ];
  
  const extractedLinks = {};
  
  function extractLink(textObj) {
    const textStr = textObj.getText().trim();
    if (!textStr) return;
    for (let i = 0; i < textStr.length; i++) {
      const url = textObj.getLinkUrl(i);
      if (url) {
        extractedLinks[textStr] = url;
        break;
      }
    }
  }
  
  doc.getBody().getParagraphs().forEach(p => extractLink(p.editAsText()));
  doc.getBody().getTables().forEach(table => {
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      for (let c = 0; c < row.getNumCells(); c++) {
        extractLink(row.getCell(c).editAsText());
      }
    }
  });
  
  const finalLinks = [];
  top10Names.forEach(name => {
    let foundUrl = null;
    for (let key in extractedLinks) {
      if (key.toLowerCase().includes(name.toLowerCase())) {
        foundUrl = extractedLinks[key];
        break;
      }
    }
    if (foundUrl) {
      finalLinks.push({
        name: name,
        url: foundUrl,
        popup: name.toLowerCase() === 'chatbot' ? 'chatbot' : ''
      });
    }
  });
  
  finalLinks.push({
    name: "📄 VIEW ALL LINKS",
    url: "https://docs.google.com/document/d/" + docId + "/edit",
    popup: ""
  });
  
  return finalLinks;
}

// ==========================================
// ZIP CODE SEARCH LOGIC
// ==========================================
function checkZipCode(zip) {
  if (!zip) return { found: false, message: "No zip code provided." };
  const ssId = '1O7L6NVy7N4vS47IMqlGZmExdOD13l8e8kmAuXw2AAKw'; 
  
  try {
    const ss = SpreadsheetApp.openById(ssId);
    const cleanZip = zip.trim();
    
    let textFinder = ss.createTextFinder(cleanZip).matchEntireCell(true);
    let results = textFinder.findAll();
    
    if (results.length === 0) {
      textFinder = ss.createTextFinder(cleanZip);
      results = textFinder.findAll().filter(r => r.getValue().toString().trim() === cleanZip);
    }
    
    if (results.length > 0) {
      let isProZip = false;
      let isDSZip = false;
      
      for (let i = 0; i < results.length; i++) {
        const r = results[i];
        const sheet = r.getSheet();
        const tabName = sheet.getName();
        const col = r.getColumn();
        
        if (tabName.includes("Official Pro Zips")) {
          isProZip = true;
        } else if (tabName.includes("Official DS (Upgrade & Service)")) {
          const header1 = sheet.getRange(1, col).getValue().toString().toUpperCase();
          if (!header1.includes("NO TECH")) {
            isDSZip = true;
          }
        }
      }
      
      if (isProZip) {
        return { found: true, message: "✅ In Footprint" };
      } else if (isDSZip) {
        return { found: true, message: "✅ Dealer Service" };
      } else {
        return { found: false, message: "❌ No Tech Available" };
      }
    } else {
      return { found: false, message: "❌ No Tech Available" };
    }
  } catch (e) {
    return { found: false, message: "Error searching zip codes. Please try again." };
  }
}

// ==========================================
// DYNAMIC FORM SCHEMA LOGIC
// ==========================================
function getFormSchema() {
  const sheetId = '1UdgMgfNU82pXRnC2MP_UsAD-HGxIf2A1h5EFf3-aIWY'; 
  const ss = SpreadsheetApp.openById(sheetId);
  const tabNames = ['Inbound', 'Outbound', 'Residential', 'Dealer', 'Addendum'];
  const schema = {};
  
  tabNames.forEach(tabName => {
    const sheet = ss.getSheetByName(tabName);
    const key = tabName.toLowerCase();
    schema[key] = {};
    
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
          schema[key][data[i][0].toString().trim()] = {
            label: data[i][1].toString().trim(),
            type: data[i][2].toString().trim(),
            options: data[i][3].toString().trim(),
            visible: data[i][4].toString().trim()
          };
        }
      }
    }
  });
  
  return schema;
}

// ==========================================
// SECURE PASSCODE LOGIC
// ==========================================
function getAdminPasscode() {
  const sheetId = '1UdgMgfNU82pXRnC2MP_UsAD-HGxIf2A1h5EFf3-aIWY'; 
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const settingsSheet = ss.getSheetByName('Settings');
    if (settingsSheet) {
      const passcode = settingsSheet.getRange('B1').getValue().toString().trim();
      return passcode;
    }
  } catch(e) {}
  return "Admin8454";
}

// ==========================================
// FEEDBACK SUBMISSION LOGIC
// ==========================================
function submitFeedback(data) {
  const sheetId = '1UdgMgfNU82pXRnC2MP_UsAD-HGxIf2A1h5EFf3-aIWY';
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    let feedbackSheet = ss.getSheetByName('Feedback');
    
    if (!feedbackSheet) {
      feedbackSheet = ss.insertSheet('Feedback');
      feedbackSheet.appendRow([
        'Timestamp',
        'Agent Name',
        'Extension',
        'Department',
        'Category',
        'Applicable Form',
        'Details / Comments'
      ]);
      feedbackSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#f3f3f3');
      feedbackSheet.setFrozenRows(1);
    }
    
    feedbackSheet.appendRow([
      new Date(),
      data.name || '',
      data.phone || '',
      data.dept || '',
      data.category || '',
      data.section || '',
      data.details || ''
    ]);
    
    return { success: true, message: 'Thank you! Your feedback has been submitted.' };
  } catch (err) {
    return { success: false, message: 'Failed to submit feedback: ' + err.toString() };
  }
}
