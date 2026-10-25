  // ==========================================
  // CONFIGURATION & HELPER UTILITIES
  // ==========================================

  /**
   * Case-insensitive tab name lookup.
   */
  function getSheetByNameInsensitive_(ss, name) {
    const wanted = String(name).trim().toLowerCase();
    return ss.getSheets().find(function(sheet) {
      return sheet.getName().trim().toLowerCase() === wanted;
    }) || null;
  }

  /**
   * Opens the master configuration spreadsheet explicitly.
   */
  function getAppSpreadsheet() {
    const masterDocId = '1a3qJOmQbn2eywwOEjA_vhLY3yCFU3j3qHrQ9hEvyEDo';
    try {
      return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(masterDocId);
    } catch (err) {
      return SpreadsheetApp.openById(masterDocId);
    }
  }

  /**
   * Converts spreadsheet settings into booleans.
   */
  function isEnabledSetting(value, defaultValue) {
    const fallback = defaultValue === true;
    if (value === true) return true;
    if (value === false) return false;

    const normalized = String(value === undefined || value === null ? '' : value).trim().toLowerCase();
    if (normalized === 'true' || normalized === 'yes') return true;
    if (normalized === 'false' || normalized === 'no') return false;
    return fallback;
  }

  /**
   * Converts a spreadsheet value into trimmed text. Preserves numeric zero.
   */
  function cleanSettingText(value) {
    return value === undefined || value === null ? '' : String(value).trim();
  }

  /**
   * Reads authorized extensions dynamically from Settings!B3.
   */
  function getAuthorizedExtensions_(ss) {
    const spreadsheet = ss || getAppSpreadsheet();
    const settingsSheet = getSheetByNameInsensitive_(spreadsheet, 'Settings');
    if (!settingsSheet) return [];

    return cleanSettingText(settingsSheet.getRange('B3').getValue())
      .split(',')
      .map(function(extension) {
        return extension.trim();
      })
      .filter(Boolean);
  }

  // ==========================================
  // WEB APP ROUTING LOGIC (LOGIN FIRST)
  // ==========================================

  function doGet(e) {
    const pageParam = cleanSettingText(
      e && e.parameter ? e.parameter.page : ''
    ).toLowerCase();

    // If ?page=notes or ?page=index is explicitly passed, load Notes. Otherwise ALWAYS start at Login!
    const isNotes = (pageParam === 'notes' || pageParam === 'index');
    const candidateNames = isNotes
      ? ['Notes', 'notes', 'Notes.html', 'notes.html']
      : ['Login', 'login', 'Login.html', 'login.html'];

    let template = null;
    for (let i = 0; i < candidateNames.length; i++) {
      try {
        template = HtmlService.createTemplateFromFile(candidateNames[i]);
        break;
      } catch (err) {}
    }

    if (!template) {
      try {
        template = HtmlService.createTemplateFromFile(isNotes ? 'Notes' : 'Login');
      } catch (e2) {
        template = HtmlService.createTemplateFromFile('Login');
      }
    }

    template.appUrl = ScriptApp.getService().getUrl();
    return template
      .evaluate()
      .setTitle('Note Taker - AAA Smart Home')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ==========================================
  // DYNAMIC FORM SCHEMA GENERATOR
  // ==========================================

  /**
   * Reads all active form tabs + login configuration.
   */
  function getFormSchema() {
    const ss = getAppSpreadsheet();

    // Explicitly supported tabs: Includes login, core forms, and new departments
    const supportedTabs = [
      'login',
      'index',
      'inbound',
      'outbound',
      'residential',
      'dealer',
      'addendum',
      'parts',
      'retention',
      'mx'
    ];

    const sheets = ss.getSheets().filter(function(sheet) {
      const name = sheet.getName().trim().toLowerCase();
      return supportedTabs.includes(name);
    });

    const schema = {};

    // Administrative tabs excluded from form generation
    const excludedSystemTabs = [
      'settings',
      'change log',
      'feedback',
      'special emails',
      'admin guide'
    ];

    sheets.forEach(function(sheet) {
      const tabName = sheet.getName().trim();
      const cleanKey = tabName.toLowerCase();

      if (excludedSystemTabs.includes(cleanKey)) {
        return;
      }

      schema[cleanKey] = {};
      schema[cleanKey + '_order'] = [];

      const data = sheet.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const id = cleanSettingText(row[0]);
        if (!id) continue;

        const rawType = cleanSettingText(row[2]).toLowerCase() || 'text';
        const isVisible = isEnabledSetting(row[4], true);
        const isRequired = isEnabledSetting(row[5], false);
        const isCopyBlank = isEnabledSetting(row[7], false);

        const fieldObj = {
          id: id,
          label: cleanSettingText(row[1]) || id,
          type: rawType,
          options: cleanSettingText(row[3]),
          visible: isVisible ? 'TRUE' : 'FALSE',
          required: isRequired,
          dependsOn: cleanSettingText(row[6]),
          copyIfBlank: isCopyBlank,
          rowOrder: i
        };

        schema[cleanKey][id] = fieldObj;
        schema[cleanKey + '_order'].push(fieldObj);
      }
    });

    // Ensure both schema.login and schema.index are accessible for backwards compatibility
    if (schema.login && !schema.index) {
      schema.index = schema.login;
      schema.index_order = schema.login_order;
    } else if (schema.index && !schema.login) {
      schema.login = schema.index;
      schema.login_order = schema.index_order;
    }

    try {
      schema.authorizedExtensions = getAuthorizedExtensions_(ss);
    } catch (err) {
      Logger.log('Error reading authorized extensions: ' + err);
      schema.authorizedExtensions = [];
    }

    return schema;
  }

  // ==========================================
  // STANDARD EMAIL TEMPLATES (WHOLE TEAM)
  // ==========================================

  /**
   * Fetches the standard email templates Google Doc as HTML.
   */
  function getEmailTemplatesHtml() {
    const docId = '1CJtJskbXT-fY2qNfna9W-i-x3rlJl404W4bgMoKFm0c';

    try {
      DriveApp.getFileById(docId);
      const url = 'https://docs.google.com/feeds/download/documents/export/Export?id=' + docId + '&exportFormat=html';
      const options = {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();

      if (code === 200) {
        return response.getContentText();
      }

      return `
        <div style="padding:16px; color:#b91c1c; font-family:sans-serif;">
          <strong>Error Loading Templates:</strong> Server returned HTTP status ${code}.
          Please verify document permissions for: <code>${docId}</code>.
        </div>
      `;
    } catch (err) {
      Logger.log('Error loading standard email templates: ' + err);
      return `
        <div style="padding:16px; color:#b91c1c; font-family:sans-serif;">
          <strong>Error Loading Email Templates:</strong> ${escapeHtml_(err.message || String(err))}
        </div>
      `;
    }
  }

  /**
   * Escapes plain text before inserting it into HTML.
   */
  function escapeHtml_(value) {
    return String(value === undefined || value === null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  // ==========================================
  // RESTRICTED SPECIAL EMAIL TEMPLATES (PROJECT 2026)
  // ==========================================

  function getSpecialEmailTemplates(userExtension) {
    try {
      const ss = getAppSpreadsheet();
      const authorizedExts = getAuthorizedExtensions_(ss);
      const cleanUserExt = cleanSettingText(userExtension);

      if (!authorizedExts.includes(cleanUserExt)) {
        return [];
      }

      const sheet = getSheetByNameInsensitive_(ss, 'Special Emails');
      if (!sheet) return [];

      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return [];

      const templates = [];
      for (let i = 1; i < data.length; i++) {
        const templateId = cleanSettingText(data[i][0]);
        const brand = cleanSettingText(data[i][1]);
        const subject = cleanSettingText(data[i][2]);
        const body = cleanSettingText(data[i][3]);

        if (!templateId && !brand) continue;

        const displayName = `🔒 Project 2026 - ${brand}`;
        const formattedBody = body
          .split(/\r?\n\r?\n/)
          .map(function(paragraph) {
            return '<p>' + escapeHtml_(paragraph).replace(/\r?\n/g, '<br>') + '</p>';
          })
          .join('<br>');

        const fullHtml = `
          <div style="font-family:Arial,sans-serif; line-height:1.5;">
            <p><strong>SUBJECT: ${escapeHtml_(subject)}</strong></p>
            <br>
            ${formattedBody}
          </div>
        `;

        templates.push({
          id: templateId,
          brand: brand,
          subject: subject,
          name: displayName,
          html: fullHtml,
          body: body
        });
      }

      return templates;
    } catch (err) {
      Logger.log('Error in getSpecialEmailTemplates: ' + err);
      return [];
    }
  }

  // ==========================================
  // DYNAMIC HELPFUL LINKS
  // ==========================================

  function getHelpfulLinksFromDoc() {
    const docId = '1R0ThNGCFYMvl28rJ-PwuCaFHsLV6chOeSSpgAFvVuzU';
    const viewAllLink = {
      name: '📄 VIEW ALL LINKS',
      url: 'https://docs.google.com/document/d/' + docId + '/edit',
      popup: ''
    };

    try {
      const doc = DocumentApp.openById(docId);
      const top10Names = [
        'Tech Support Opportunities',
        'Chatbot',
        'Common Parts List',
        'Test Mode SOP',
        'DS Zip Codes',
        'Tech Support Tier I Guide',
        'Known Issues Tracker',
        'Waived Fee Guideline',
        'Call Routing Guide',
        'Actual/Incident Form'
      ];

      const extractedLinks = {};
      function extractLink(textObj) {
        const rawText = textObj.getText();
        const displayText = rawText.trim();
        if (!displayText) return;

        for (let i = 0; i < rawText.length; i++) {
          const url = textObj.getLinkUrl(i);
          if (url) {
            extractedLinks[displayText] = url;
            break;
          }
        }
      }

      doc.getBody().getParagraphs().forEach(function(paragraph) {
        extractLink(paragraph.editAsText());
      });

      doc.getBody().getTables().forEach(function(table) {
        for (let r = 0; r < table.getNumRows(); r++) {
          const row = table.getRow(r);
          for (let c = 0; c < row.getNumCells(); c++) {
            extractLink(row.getCell(c).editAsText());
          }
        }
      });

      const finalLinks = [];
      top10Names.forEach(function(name) {
        let foundUrl = null;
        for (const key in extractedLinks) {
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

      finalLinks.push(viewAllLink);
      return finalLinks;
    } catch (err) {
      Logger.log('Error extracting helpful links: ' + err);
      return [viewAllLink];
    }
  }

  // ==========================================
  // ZIP CODE SEARCH
  // ==========================================

  function checkZipCode(zip) {
    const cleanZip = cleanSettingText(zip);
    if (!cleanZip) {
      return { found: false, message: 'No zip code provided.' };
    }

    const ssId = '1O7L6NVy7N4vS47IMqlGZmExdOD13l8e8kmAuXw2AAKw';

    try {
      const ss = SpreadsheetApp.openById(ssId);
      let textFinder = ss.createTextFinder(cleanZip).matchEntireCell(true);
      let results = textFinder.findAll();

      if (results.length === 0) {
        textFinder = ss.createTextFinder(cleanZip);
        results = textFinder.findAll().filter(function(range) {
          return cleanSettingText(range.getValue()) === cleanZip;
        });
      }

      if (results.length === 0) {
        return { found: false, message: '❌ No Tech Available' };
      }

      let isProZip = false;
      let isDSZip = false;

      for (let i = 0; i < results.length; i++) {
        const range = results[i];
        const sheet = range.getSheet();
        const tabName = sheet.getName();
        const col = range.getColumn();

        if (tabName.includes('Official Pro Zips')) {
          isProZip = true;
        } else if (tabName.includes('Official DS (Upgrade & Service)')) {
          const header = cleanSettingText(sheet.getRange(1, col).getValue()).toUpperCase();
          if (!header.includes('NO TECH')) {
            isDSZip = true;
          }
        }
      }

      if (isProZip) return { found: true, message: '✅ In Footprint' };
      if (isDSZip) return { found: true, message: '✅ Dealer Service' };
      return { found: false, message: '❌ No Tech Available' };
    } catch (err) {
      return { found: false, message: 'Error searching zip codes: ' + (err.message || String(err)) };
    }
  }

  // ==========================================
  // ADMIN PASSCODE & FEEDBACK
  // ==========================================

  function getAdminPasscode() {
    try {
      const ss = getAppSpreadsheet();
      const settingsSheet = getSheetByNameInsensitive_(ss, 'Settings');
      if (settingsSheet) {
        const passcode = cleanSettingText(settingsSheet.getRange('B1').getValue());
        return passcode || '8454';
      }
    } catch (err) {
      Logger.log('Error reading admin passcode: ' + err);
    }
    return '8454';
  }

  function submitFeedback(data) {
    try {
      data = data || {};
      const ss = getAppSpreadsheet();
      let feedbackSheet = getSheetByNameInsensitive_(ss, 'Feedback');

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
        data.phone || data.extension || '',
        data.dept || '',
        data.category || '',
        data.section || data.applicableForm || '',
        data.details || data.comments || ''
      ]);

      return { success: true, message: 'Thank you! Your feedback has been submitted.' };
    } catch (err) {
      return { success: false, message: 'Failed to submit feedback: ' + String(err) };
    }
  }

  // ==========================================
  // CHANGE LOG
  // ==========================================

  function getChangeLog() {
    try {
      const ss = getAppSpreadsheet();
      const sheet = getSheetByNameInsensitive_(ss, 'Change Log');
      if (!sheet) return [];

      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return [];

      const tz = ss.getSpreadsheetTimeZone();
      const visibleRows = data.slice(1).filter(function(row) {
        const isVisible = isEnabledSetting(row[5], false);
        const hasContent = row[0] || row[3];
        return isVisible && Boolean(hasContent);
      });

      return visibleRows.map(function(row) {
        return {
          date: row[0] instanceof Date
            ? Utilities.formatDate(row[0], tz, 'yyyy-MM-dd')
            : cleanSettingText(row[0]),
          version: cleanSettingText(row[1]),
          form: cleanSettingText(row[2]),
          description: cleanSettingText(row[3]),
          updatedBy: cleanSettingText(row[4])
        };
      });
    } catch (err) {
      Logger.log('Error reading change log: ' + err);
      return [];
    }
  }

  // ==========================================
  // AUTOMATIC CHANGE LOG TRIGGER
  // ==========================================

  function onEdit(e) {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    const excludedTabs = ['Change Log', 'Feedback', 'Special Emails'];

    if (excludedTabs.includes(sheetName)) return;
    if (e.range.getRow() === 1) return;

    const ss = e.source || getAppSpreadsheet();
    const logSheet = getSheetByNameInsensitive_(ss, 'Change Log');
    if (!logSheet) return;

    const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    const user = Session.getActiveUser().getEmail() || 'Admin';
    const cellNotation = e.range.getA1Notation();

    let fieldIdentifier = '';
    try {
      const row = e.range.getRow();
      const idVal = cleanSettingText(sheet.getRange(row, 1).getValue());
      const labelVal = cleanSettingText(sheet.getRange(row, 2).getValue());
      fieldIdentifier = labelVal || idVal || `Row ${row}`;
    } catch (err) {
      fieldIdentifier = `Row ${e.range.getRow()}`;
    }

    let description;
    if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) {
      description = `Updated range ${cellNotation} beginning at [${fieldIdentifier}]. Multiple cells were edited; individual old/new values are not available.`;
    } else {
      const oldValue = e.oldValue !== undefined ? `"${e.oldValue}"` : '(blank)';
      const newValue = e.value !== undefined ? `"${e.value}"` : '(blank)';
      description = `Updated [${fieldIdentifier}] in ${cellNotation}: changed from ${oldValue} to ${newValue}`;
    }

    logSheet.insertRowBefore(2);
    logSheet.getRange(2, 6).insertCheckboxes();
    logSheet.getRange(2, 1, 1, 6).setValues([[
      timestamp,
      'v1.x',
      sheetName,
      description,
      user,
      true
    ]]);
  }
