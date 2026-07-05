// ============================================================
// HR AUTOMATION - ATOMS DIGITAL SOLUTIONS
// WEB APP VERSION
// ============================================================

const CONFIG = {
  sheets: {
    data: 'Data',
    rolesResponsibilities: 'Roles_Responsibilities',
    subjectBody: 'Subject_Body',
    letterBody: 'Letter_body',
    logs: 'Logs'
  },
  headerImages: {
    letterHeader: '15Xkigpy0YeZKY05vaeC4N6MHXGFgIoMB',
    certificateHeader: '1sh_QfM98oeZN1y0OyZjpS7LOLM3TtJ9b'
  },
  certificateTypes: [
    'Certification of Appreciation',
    'Certification of Completion'
  ]
};

// ============================================================
// WEB APP ENTRY POINTS
// ============================================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('HR Agent — Atoms Digital Solutions')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function handleAgentRequest(paramsJson) {
  try {
    const params = JSON.parse(paramsJson);
    const action = params.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === 'getLetterTypes') {
      const sheet = ss.getSheetByName(CONFIG.sheets.letterBody);
      const data = sheet.getDataRange().getValues().slice(1).map(r => r[0]).filter(v => v);
      return { success: true, data: data };
    }

    if (action === 'getRoles') {
      const sheet = ss.getSheetByName(CONFIG.sheets.rolesResponsibilities);
      const data = sheet.getDataRange().getValues().slice(1).map(r => r[0]).filter(v => v);
      return { success: true, data: data };
    }

    if (action === 'previewLetter') {
      const formData = params.formData;
      const getValue = (col) => { const val = formData[col]; return val ? val.toString().trim() : ''; };
      const letterType = getValue('Letter Type');
      const role = getValue('Role');
      const responsibilities = getResponsibilities(ss, role);
      const getV = (col) => {
        if (col === 'RESPONSIBILITES' || col === 'RESPONSIBILITIES') return responsibilities;
        if (col === 'HR Team') return getValue('HR Name');
        return getValue(col);
      };
      const { subject, emailBody } = getSubjectAndBody(ss, letterType);
      return {
        success: true,
        data: {
          subject: replacePlaceholders(subject, getV),
          emailBody: replacePlaceholders(emailBody, getV),
          recipientEmail: getValue('Email'),
          recipientName: getValue('Full Name')
        }
      };
    }

    if (action === 'sendLetter') {
      const formData = params.formData;
      const getValue = (col) => { const val = formData[col]; return val ? val.toString().trim() : ''; };
      const letterType = getValue('Letter Type');
      const role = getValue('Role');
      const fullName = getValue('Full Name');
      const email = getValue('Email').replace(/\s+/g, '').toLowerCase();
      if (!email || !email.includes('@')) throw new Error('Invalid email: ' + email);

      const responsibilities = getResponsibilities(ss, role);
      const getV = (col) => {
        if (col === 'RESPONSIBILITES' || col === 'RESPONSIBILITIES') return responsibilities;
        if (col === 'HR Team') return getValue('HR Name');
        return getValue(col);
      };

      const { subject, emailBody } = getSubjectAndBody(ss, letterType);
      const finalSubject = replacePlaceholders(subject, getV);
      const finalEmailBody = replacePlaceholders(emailBody, getV);

      const pdfBlob = generatePDF(letterType, fullName, getV, ss);
      pdfBlob.setName(`${role || letterType} - ${fullName}.pdf`);

      const startDateVal = getValue('Start Date');
      let year = new Date().getFullYear().toString();
      if (startDateVal) { const p = new Date(startDateVal); if (!isNaN(p)) year = p.getFullYear().toString(); }

      const savedPdf = getOrCreatePdfFolder(letterType, year).createFile(pdfBlob);
      savedPdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const pdfUrl = savedPdf.getUrl();

      GmailApp.sendEmail(email, finalSubject, finalEmailBody, {
        htmlBody: buildHtmlEmail(finalEmailBody, letterType),
        attachments: [pdfBlob],
        name: 'Atoms Digital Solutions',
        cc: formData['CC'] || ''
      });

      // Save to Data sheet
      const dataSheet = ss.getSheetByName(CONFIG.sheets.data);
      if (dataSheet && dataSheet.getLastColumn() > 0) {
        const headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
        const row = headers.map(h => formData[h] ? formData[h].toString().trim() : '');
        const statusIdx = headers.indexOf('Status');
        if (statusIdx >= 0) row[statusIdx] = 'Sent';
        dataSheet.appendRow(row);
      }

      saveLog(ss, fullName, letterType, 'Sent', 'Sent via Web App', pdfUrl);
      return { success: true, data: { message: 'Email sent!' } };
    }

    return { success: false, error: 'Unknown action: ' + action };
  } catch(e) {
    Logger.log('handleAgentRequest ERROR: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================
// ON OPEN MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏢 HR Automation')
    .addItem('Sync Columns', 'syncColumns')
    .addItem('Send Emails', 'sendEmails')
    .addItem('Preview Email', 'previewEmail')
    .addItem('Fix Dropdowns', 'fixDropdowns')
    .addItem('Beautify Sheet', 'beautifySheet')
    .addItem('View Logs', 'viewLogs')
    .addToUi();
}

// ============================================================
// SYNC COLUMNS
// ============================================================
function syncColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const subjectBodySheet = ss.getSheetByName(CONFIG.sheets.subjectBody);
  const letterBodySheet = ss.getSheetByName(CONFIG.sheets.letterBody);
  const dataSheet = ss.getSheetByName(CONFIG.sheets.data);

  const placeholders = new Set();

  // Get placeholders from Subject_Body
  const sbData = subjectBodySheet.getDataRange().getValues();
  for (let i = 1; i < sbData.length; i++) {
    extractPlaceholders(sbData[i][1]).forEach(p => placeholders.add(p));
    extractPlaceholders(sbData[i][2]).forEach(p => placeholders.add(p));
  }

  // Get placeholders from Letter_Body
  const lbData = letterBodySheet.getDataRange().getValues();
  for (let i = 1; i < lbData.length; i++) {
    extractPlaceholders(lbData[i][1]).forEach(p => placeholders.add(p));
  }

  // Remove system placeholders
  placeholders.delete('RESPONSIBILITES');
  placeholders.delete('HR Team');

  // Get existing columns in Data sheet
  const syncLastCol = dataSheet.getLastColumn();
  const existingHeaders = syncLastCol > 0 ? dataSheet.getRange(1, 1, 1, syncLastCol).getValues()[0] : [];

  // Add missing columns
  let added = 0;
  placeholders.forEach(placeholder => {
    if (!existingHeaders.includes(placeholder)) {
      const newCol = dataSheet.getLastColumn() + 1;
      dataSheet.getRange(1, newCol).setValue(placeholder);
      added++;
    }
  });

  // Add required columns if missing (in order)
  const requiredCols = ['Full Name', 'Email', 'Letter Type', 'Role', 'Employment Type', 'Mode', 'Start Date', 'End Date', 'Date', 'HR Name', 'Designation', 'Compensation', 'Tenure', 'Department Name', 'Reason 1', 'Reason 2', 'Status'];
  requiredCols.forEach(col => {
    const currentLastCol = dataSheet.getLastColumn();
    const currentHeaders = currentLastCol > 0 ? dataSheet.getRange(1, 1, 1, currentLastCol).getValues()[0] : [];
    if (!currentHeaders.includes(col)) {
      dataSheet.getRange(1, dataSheet.getLastColumn() + 1).setValue(col);
    }
  });

  // Add Status column if missing
  const updatedLastCol = dataSheet.getLastColumn();
  const updatedHeaders = updatedLastCol > 0 ? dataSheet.getRange(1, 1, 1, updatedLastCol).getValues()[0] : [];
  if (!updatedHeaders.includes('Status')) {
    dataSheet.getRange(1, dataSheet.getLastColumn() + 1).setValue('Status');
  }

  SpreadsheetApp.getUi().alert(`✅ Sync Complete!\n${added} new columns added to Data sheet.`);
}

function extractPlaceholders(text) {
  if (!text) return [];
  const matches = text.match(/\{\{([^}]+)\}\}/g) || [];
  return matches.map(m => m.replace(/\{\{|\}\}/g, '').trim());
}

// ============================================================
// FIX DROPDOWNS
// ============================================================
function fixDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CONFIG.sheets.data);

  const lastCol = dataSheet.getLastColumn();
  if (lastCol < 1) {
    SpreadsheetApp.getUi().alert('Data sheet empty! Sync Columns first run cheyyi.');
    return;
  }

  const headers = dataSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Letter Type dropdown
  const letterBodySheet = ss.getSheetByName(CONFIG.sheets.letterBody);
  const lbLastRow = letterBodySheet.getLastRow();
  if (lbLastRow >= 2) {
    const letterTypes = letterBodySheet.getRange(2, 1, lbLastRow - 1, 1).getValues().flat().filter(v => v);
    const letterTypeCol = headers.indexOf('Letter Type') + 1;
    if (letterTypeCol > 0 && letterTypes.length > 0) {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(letterTypes, true)
        .build();
      dataSheet.getRange(2, letterTypeCol, 1000).setDataValidation(rule);
    }
  }

  // Role dropdown
  const rolesSheet = ss.getSheetByName(CONFIG.sheets.rolesResponsibilities);
  const rLastRow = rolesSheet.getLastRow();
  if (rLastRow >= 2) {
    const roles = rolesSheet.getRange(2, 1, rLastRow - 1, 1).getValues().flat().filter(v => v);
    const roleCol = headers.indexOf('Role') + 1;
    if (roleCol > 0 && roles.length > 0) {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(roles, true)
        .build();
      dataSheet.getRange(2, roleCol, 1000).setDataValidation(rule);
    }
  }

  // Employment Type dropdown
  const empTypeCol = headers.indexOf('Employment Type') + 1;
  if (empTypeCol > 0) {
    const empRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Internship', 'Full-Time', 'Work from Home', 'Work from Office'], true)
      .build();
    dataSheet.getRange(2, empTypeCol, 1000).setDataValidation(empRule);
  }

  // Mode dropdown
  const modeCol = headers.indexOf('Mode') + 1;
  if (modeCol > 0) {
    const modeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Work from Home', 'Work from Office'], true)
      .build();
    dataSheet.getRange(2, modeCol, 1000).setDataValidation(modeRule);
  }

  SpreadsheetApp.getUi().alert('✅ Dropdowns fixed!');
}

// ============================================================
// SEND EMAILS
// ============================================================
function sendEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CONFIG.sheets.data);
  const data = dataSheet.getDataRange().getValues();
  const headers = data[0];

  let sent = 0, failed = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const statusCol = headers.indexOf('Status');
    const status = row[statusCol];

    if (status === 'Sent') continue;
    if (!row[headers.indexOf('Email')]) continue;

    try {
      processRow(row, headers, ss);
      dataSheet.getRange(i + 1, statusCol + 1).setValue('Sent');
      sent++;
    } catch (e) {
      dataSheet.getRange(i + 1, statusCol + 1).setValue('Error');
      saveLog(ss, row[headers.indexOf('Full Name')], row[headers.indexOf('Letter Type')], 'Error', e.message);
      failed++;
    }
  }

  SpreadsheetApp.getUi().alert(`✅ Done!\nSent: ${sent}\nFailed: ${failed}`);
}

// ============================================================
// PROCESS ROW
// ============================================================
function processRow(row, headers, ss) {
  const getValue = (col) => {
    const idx = headers.indexOf(col);
    if (idx < 0) return '';
    const val = row[idx];
    if (!val) return '';
    // Format dates properly
    if (val instanceof Date) {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd MMM yyyy');
    }
    return val.toString();
  };

  const letterType = getValue('Letter Type');
  const role = getValue('Role');
  const email = getValue('Email');
  const fullName = getValue('Full Name');

  // Get Subject & Email Body
  const { subject, emailBody } = getSubjectAndBody(ss, letterType);

  // Get Responsibilities
  const responsibilities = getResponsibilities(ss, role);

  // Extend getValue to include responsibilities
  const getValueExtended = (col) => {
    if (col === 'RESPONSIBILITES' || col === 'RESPONSIBILITIES') return responsibilities;
    if (col === 'HR Team') return getValue('HR Name');
    return getValue(col);
  };

  // Replace placeholders in email
  const finalEmailBody = replacePlaceholders(emailBody, getValueExtended);
  const finalSubject = replacePlaceholders(subject, getValueExtended);

  // Generate PDF — passes getValue so Doc placeholders get replaced
  const pdfBlob = generatePDF(letterType, fullName, getValueExtended, ss);
  pdfBlob.setName(`${role} - ${fullName}.pdf`);

  // Get year from Start Date, fallback to current year
  const startDateVal = getValue('Start Date');
  let year = new Date().getFullYear().toString();
  if (startDateVal) {
    const parsed = new Date(startDateVal);
    if (!isNaN(parsed)) year = parsed.getFullYear().toString();
  }

  // Save PDF to Drive with folder structure
  const savedPdf = getOrCreatePdfFolder(letterType, year).createFile(pdfBlob);
  savedPdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const pdfUrl = savedPdf.getUrl();

  // Build beautiful HTML email body
  const htmlEmailBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { margin: 0; padding: 0; background: #f4f6f9; font-family: Arial, sans-serif; }
    .wrapper { max-width: 620px; margin: 30px auto; background: #ffffff; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 20px rgba(0,0,0,0.08); }
    .header { background: linear-gradient(135deg, #1F4E79, #2E75B6); padding: 30px 40px; text-align: center; }
    .header h1 { color: #ffffff; margin: 0; font-size: 22px; letter-spacing: 1px; }
    .header p { color: #BDD7EE; margin: 6px 0 0; font-size: 13px; }
    .body { padding: 35px 40px; color: #333333; font-size: 14px; line-height: 1.8; }
    .body p { margin: 0 0 14px; }
    .attachment-box { background: #EBF3FB; border-left: 4px solid #2E75B6; border-radius: 6px; padding: 14px 18px; margin: 24px 0; display: flex; align-items: center; gap: 12px; }
    .attachment-box span { font-size: 28px; }
    .attachment-box div { font-size: 13px; color: #1F4E79; font-weight: bold; }
    .attachment-box small { color: #555; font-weight: normal; display: block; }
    .footer { background: #F4F6F9; padding: 20px 40px; text-align: center; border-top: 1px solid #E0E0E0; }
    .footer p { margin: 0; font-size: 12px; color: #888888; }
    .footer strong { color: #1F4E79; }
  </style>
</head>
<body>
  <div class="wrapper">
    <div class="header">
      <h1>🏢 Atoms Digital Solutions</h1>
      <p>HR Communication</p>
    </div>
    <div class="body">
      ${finalEmailBody.replace(/\n/g, '<br>')}
      <div class="attachment-box">
        <span>📄</span>
        <div>${letterType} - ${fullName}.pdf <small>Please find your letter attached.</small></div>
      </div>
    </div>
    <div class="footer">
      <p><strong>Atoms Digital Solutions</strong><br>This is an automated email. Please do not reply directly.</p>
    </div>
  </div>
</body>
</html>`;

  // Send Email
  GmailApp.sendEmail(email, finalSubject, '', {
    htmlBody: htmlEmailBody,
    attachments: [pdfBlob],
    name: 'Atoms Digital Solutions'
  });

  // Save Log with PDF URL
  saveLog(ss, fullName, letterType, 'Sent', 'Email sent successfully', pdfUrl);
}

// ============================================================
// GET SUBJECT & BODY
// ============================================================
function getSubjectAndBody(ss, letterType) {
  const sheet = ss.getSheetByName(CONFIG.sheets.subjectBody);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === letterType) {
      return {
        subject: data[i][1],
        emailBody: data[i][2]
      };
    }
  }
  throw new Error(`Letter Type "${letterType}" not found in Subject_Body sheet`);
}

// ============================================================
// GET LETTER BODY (Google Doc URL or plain text)
// ============================================================
function getLetterBody(ss, letterType) {
  const sheet = ss.getSheetByName(CONFIG.sheets.letterBody);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === letterType) {
      const cellValue = data[i][1].toString().trim();
      if (cellValue.startsWith('https://docs.google.com')) {
        const docId = extractDocId(cellValue);
        if (docId) return fetchDocContent(cellValue, docId);
      }
      return cellValue;
    }
  }
  throw new Error(`Letter Type "${letterType}" not found in Letter_body sheet`);
}

function extractDocId(url) {
  const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

function fetchDocContent(url, docId) {
  try {
    if (url.includes('/presentation/')) {
      const presentation = SlidesApp.openById(docId);
      let content = '';
      presentation.getSlides().forEach(slide => {
        slide.getShapes().forEach(shape => {
          if (shape.getText) {
            content += shape.getText().asString() + '\n';
          }
        });
      });
      return content;
    } else {
      const doc = DocumentApp.openById(docId);
      return doc.getBody().getText();
    }
  } catch(e) {
    throw new Error(`Doc fetch failed: ${e.message}`);
  }
}

// ============================================================
// GET RESPONSIBILITIES
// ============================================================
function getResponsibilities(ss, role) {
  const sheet = ss.getSheetByName(CONFIG.sheets.rolesResponsibilities);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === role) {
      const raw = data[i][1].toString().trim();
      const lines = raw.split('\n')
        .map(line => line.trim())
        .filter(line => line)
        .map(line => {
          // Remove existing bullets if any
          line = line.replace(/^[•\-\*]\s*/, '');
          return '• ' + line;
        });
      return lines.join('\n');
    }
  }
  return '';
}

// ============================================================
// REPLACE PLACEHOLDERS
// ============================================================
function replacePlaceholders(text, getValue) {
  if (!text) return '';
  return text.replace(/\{\{([^}]+)\}\}/g, (match, key) => {
    const value = getValue(key.trim());
    return value !== undefined && value !== '' ? value : match;
  });
}

// ============================================================
// GENERATE PDF
// ============================================================
function generatePDF(letterType, fullName, getValue, ss) {
  const sheet = ss.getSheetByName(CONFIG.sheets.letterBody);
  const data = sheet.getDataRange().getValues();
  let docUrl = '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === letterType) {
      docUrl = data[i][1].toString().trim();
      break;
    }
  }

  if (docUrl.startsWith('https://docs.google.com')) {
    const docId = extractDocId(docUrl);
    if (docId) {
      try {
        const tempCopy = DriveApp.getFileById(docId).makeCopy(`TEMP_${Date.now()}`);
        const tempId = tempCopy.getId();

        if (docUrl.includes('/presentation/')) {
          const pres = SlidesApp.openById(tempId);
          const slides = pres.getSlides();
          
          slides.forEach(slide => {
            slide.getPageElements().forEach(element => {
              try {
                const type = element.getPageElementType();
                if (type === SlidesApp.PageElementType.SHAPE) {
                  const shape = element.asShape();
                  const tf = shape.getText();
                  const text = tf.asString();
                  const phs = [...new Set(text.match(/\{\{[^}]+\}\}/g) || [])];
                  phs.forEach(ph => {
                    const key = ph.replace(/\{\{|\}\}/g, '').trim();
                    const value = getValue(key) || '';
                    shape.getText().replaceAllText(ph, value);
                  });
                } else if (type === SlidesApp.PageElementType.TABLE) {
                  const table = element.asTable();
                  for (let r = 0; r < table.getNumRows(); r++) {
                    for (let c = 0; c < table.getNumColumns(); c++) {
                      try {
                        const tf = table.getCell(r, c).getText();
                        const text = tf.asString();
                        const phs = [...new Set(text.match(/\{\{[^}]+\}\}/g) || [])];
                        phs.forEach(ph => {
                          const key = ph.replace(/\{\{|\}\}/g, '').trim();
                          tf.replaceAllText(ph, getValue(key) || '');
                        });
                      } catch(e) {}
                    }
                  }
                }
              } catch(e) {
                Logger.log('Shape error: ' + e.message);
              }
            });
          });
          pres.saveAndClose();
        } else {
          // Google Docs — replace placeholders
          const tempDoc = DocumentApp.openById(tempId);
          const body = tempDoc.getBody();
          
          // Get all placeholders from full text
          const fullText = body.getText();
          const phs = [...new Set(fullText.match(/\{\{[^}]+\}\}/g) || [])];
          
          // Replace each placeholder — replaceText handles split runs automatically
          phs.forEach(ph => {
            const key = ph.replace(/\{\{|\}\}/g, '').trim();
            const value = getValue(key) || '';
            // Build regex that matches literal {{ and }} with the key
            const regex = '\\{\\{' + key.split('').map(c => {
              return /[.*+?^${}()|[\]\\]/.test(c) ? '\\' + c : c;
            }).join('') + '\\}\\}';
            body.replaceText(regex, value);
          });

          // Clear all text background/highlight colors
          const numChildren = body.getNumChildren();
          for (let c = 0; c < numChildren; c++) {
            try {
              const child = body.getChild(c);
              const type = child.getType();
              if (type === DocumentApp.ElementType.PARAGRAPH) {
                const para = child.asParagraph();
                const numParts = para.getNumChildren();
                for (let p = 0; p < numParts; p++) {
                  try {
                    const part = para.getChild(p);
                    if (part.getType() === DocumentApp.ElementType.TEXT) {
                      part.asText().setBackgroundColor(null);
                    }
                  } catch(e) {}
                }
              }
            } catch(e) {}
          }

          tempDoc.saveAndClose();
        }

        const pdfBlob = DriveApp.getFileById(tempId).getAs('application/pdf');
        pdfBlob.setName(`${letterType} - ${fullName}.pdf`);
        DriveApp.getFileById(tempId).setTrashed(true);
        return pdfBlob;
      } catch(e) {
        Logger.log('Doc PDF error: ' + e.message);
        throw new Error('PDF generation failed: ' + e.message);
      }
    }
  }

  // Fallback HTML PDF
  let headerHtml = '';
  try {
    const isCertificate = CONFIG.certificateTypes.includes(letterType);
    const headerId = isCertificate ? CONFIG.headerImages.certificateHeader : CONFIG.headerImages.letterHeader;
    const headerBlob = DriveApp.getFileById(headerId).getBlob();
    const headerBase64 = Utilities.base64Encode(headerBlob.getBytes());
    const headerMime = headerBlob.getContentType();
    headerHtml = `<img src="data:${headerMime};base64,${headerBase64}" style="width:100%;display:block;" />`;
  } catch(e) { headerHtml = ''; }

  const html = `<html><head><style>body{font-family:Arial,sans-serif;font-size:12px;margin:0;padding:0;}.content{padding:30px 50px;white-space:pre-wrap;line-height:1.8;}</style></head><body><div>${headerHtml}</div><div class="content">${letterBody.replace(/\n/g, '<br>')}</div></body></html>`;
  const blob = Utilities.newBlob(html, 'text/html', 'letter.html');
  const pdfBlob = blob.getAs('application/pdf');
  pdfBlob.setName(`${letterType} - ${fullName}.pdf`);
  return pdfBlob;
}

// ============================================================
// PREVIEW EMAIL (ALL ROWS SIDEBAR)
// ============================================================
function previewEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CONFIG.sheets.data);
  const data = dataSheet.getDataRange().getValues();
  const headers = data[0];

  if (data.length < 2) {
    SpreadsheetApp.getUi().alert('No data found in Data sheet!');
    return;
  }

  // Build preview data for all rows
  let rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const getValue = (col) => {
      const idx = headers.indexOf(col);
      if (idx < 0) return '';
      const val = row[idx];
      if (!val) return '';
      if (val instanceof Date) {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd MMM yyyy');
      }
      return val.toString();
    };

    const fullName = getValue('Full Name');
    const email = getValue('Email');
    const letterType = getValue('Letter Type');
    const status = getValue('Status');
    if (!fullName && !email) continue;

    const role = getValue('Role');
    const responsibilities = getResponsibilities(ss, role);
    const getValueExtended = (col) => {
      if (col === 'RESPONSIBILITES' || col === 'RESPONSIBILITIES') return responsibilities;
      if (col === 'HR Team') return getValue('HR Name');
      return getValue(col);
    };

    let finalSubject = '', finalEmailBody = '';
    try {
      const { subject, emailBody } = getSubjectAndBody(ss, letterType);
      finalSubject = replacePlaceholders(subject, getValueExtended);
      finalEmailBody = replacePlaceholders(emailBody, getValueExtended);
    } catch(e) {
      finalSubject = 'Error: ' + e.message;
      finalEmailBody = '';
    }

    // Generate PDF preview URL
    let pdfUrl = '';
    try {
      const pdfBlob = generatePDF(letterType, fullName, getValueExtended, ss);
      const tempFile = DriveApp.getRootFolder().createFile(pdfBlob);
      tempFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      pdfUrl = tempFile.getUrl();
    } catch(e) {
      pdfUrl = '';
    }

    rows.push({
      rowNum: i + 1,
      fullName, email, letterType, status,
      subject: finalSubject,
      body: finalEmailBody,
      pdfName: `${letterType} - ${fullName}.pdf`,
      pdfUrl: pdfUrl
    });
  }

  // Build HTML
  let cards = rows.map((r, idx) => {
    const statusColor = r.status === 'Sent' ? '#c6f6d5' : r.status === 'Error' ? '#fed7d7' : '#fefcbf';
    const statusText = r.status || 'Pending';
    return `
      <div class="card" id="card${idx}">
        <div class="card-header" onclick="toggle(${idx})">
          <span class="row-num">Row ${r.rowNum}</span>
          <span class="name">${r.fullName}</span>
          <span class="badge" style="background:${statusColor}">${statusText}</span>
          <span class="letter-type">${r.letterType}</span>
          <span class="arrow" id="arrow${idx}">▼</span>
        </div>
        <div class="card-body" id="body${idx}" style="display:none;">
          <div class="field"><span class="lbl">To</span><div class="val">${r.email}</div></div>
          <div class="field"><span class="lbl">Subject</span><div class="val">${r.subject}</div></div>
          <div class="field"><span class="lbl">Email Body</span><div class="val scroll">${r.body.replace(/\n/g, '<br>')}</div></div>
          <div class="field"><span class="lbl">Attachment</span>
            <div class="attachment">📄 <b>${r.pdfName}</b>${r.pdfUrl ? ` &nbsp;<a href="${r.pdfUrl}" target="_blank">👁 Preview PDF</a>` : ''}</div>
          </div>
        </div>
      </div>
    `;
  }).join('');

  const html = `
    <html>
      <head>
        <style>
          * { box-sizing: border-box; }
          body { font-family: Arial, sans-serif; font-size: 12px; margin: 0; padding: 10px; background: #f5f5f5; }
          h3 { margin: 0 0 10px; font-size: 14px; color: #333; }
          .card { background: #fff; border: 1px solid #ddd; border-radius: 6px; margin-bottom: 8px; overflow: hidden; }
          .card-header { display: flex; align-items: center; gap: 8px; padding: 10px 12px; cursor: pointer; background: #fafafa; flex-wrap: wrap; }
          .card-header:hover { background: #f0f0f0; }
          .row-num { color: #888; font-size: 11px; min-width: 40px; }
          .name { font-weight: bold; flex: 1; }
          .badge { padding: 2px 8px; border-radius: 10px; font-size: 10px; font-weight: bold; }
          .letter-type { color: #555; font-size: 11px; }
          .arrow { margin-left: auto; color: #888; }
          .card-body { padding: 12px; border-top: 1px solid #eee; }
          .field { margin-bottom: 10px; }
          .lbl { font-weight: bold; color: #555; font-size: 10px; text-transform: uppercase; display: block; margin-bottom: 3px; }
          .val { background: #f9f9f9; border: 1px solid #eee; border-radius: 4px; padding: 8px; white-space: pre-wrap; }
          .scroll { max-height: 150px; overflow-y: auto; }
          .attachment { background: #fff8e1; border: 1px solid #ffe082; border-radius: 4px; padding: 8px; }
        </style>
        <script>
          function toggle(idx) {
            const body = document.getElementById('body' + idx);
            const arrow = document.getElementById('arrow' + idx);
            if (body.style.display === 'none') {
              body.style.display = 'block';
              arrow.textContent = '▲';
            } else {
              body.style.display = 'none';
              arrow.textContent = '▼';
            }
          }
        </script>
      </head>
      <body>
        <h3>📧 Email Preview — ${rows.length} row(s)</h3>
        ${cards}
      </body>
    </html>
  `;

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(750).setHeight(600),
    '📧 Email Preview'
  );
}

// ============================================================
// BEAUTIFY SHEET
// ============================================================
function beautifySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CONFIG.sheets.data);
  if (!dataSheet) return;

  const lastCol = dataSheet.getLastColumn();
  const lastRow = dataSheet.getLastRow();
  if (lastCol < 1) return;

  // Header row styling
  const headerRange = dataSheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground('#1F4E79');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  dataSheet.setRowHeight(1, 36);

  // Freeze header row
  dataSheet.setFrozenRows(1);

  // Data rows styling
  if (lastRow > 1) {
    for (let i = 2; i <= lastRow; i++) {
      const rowRange = dataSheet.getRange(i, 1, 1, lastCol);
      if (i % 2 === 0) {
        rowRange.setBackground('#EBF3FB');
      } else {
        rowRange.setBackground('#FFFFFF');
      }
      rowRange.setFontColor('#333333');
      rowRange.setFontSize(10);
      rowRange.setVerticalAlignment('middle');
      dataSheet.setRowHeight(i, 28);
    }

    // Status column color coding
    const headers = dataSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const statusCol = headers.indexOf('Status') + 1;
    if (statusCol > 0) {
      for (let i = 2; i <= lastRow; i++) {
        const statusCell = dataSheet.getRange(i, statusCol);
        const status = statusCell.getValue();
        if (status === 'Sent') {
          statusCell.setBackground('#C6F6D5');
          statusCell.setFontColor('#276749');
          statusCell.setFontWeight('bold');
        } else if (status === 'Error') {
          statusCell.setBackground('#FED7D7');
          statusCell.setFontColor('#9B2335');
          statusCell.setFontWeight('bold');
        } else if (status === 'Pending') {
          statusCell.setBackground('#FEFCBF');
          statusCell.setFontColor('#744210');
          statusCell.setFontWeight('bold');
        }
      }
    }
  }

  // Auto resize columns
  for (let c = 1; c <= lastCol; c++) {
    dataSheet.autoResizeColumn(c);
  }

  // Border for all data
  if (lastRow > 0) {
    dataSheet.getRange(1, 1, lastRow, lastCol)
      .setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  }

  // Beautify Logs sheet too
  const logsSheet = ss.getSheetByName(CONFIG.sheets.logs);
  if (logsSheet) {
    const lLastCol = logsSheet.getLastColumn();
    const lLastRow = logsSheet.getLastRow();
    if (lLastCol > 0) {
      const lHeader = logsSheet.getRange(1, 1, 1, lLastCol);
      lHeader.setBackground('#2E75B6');
      lHeader.setFontColor('#FFFFFF');
      lHeader.setFontWeight('bold');
      lHeader.setFontSize(11);
      lHeader.setHorizontalAlignment('center');
      logsSheet.setFrozenRows(1);
      logsSheet.setRowHeight(1, 36);

      if (lLastRow > 1) {
        for (let i = 2; i <= lLastRow; i++) {
          const r = logsSheet.getRange(i, 1, 1, lLastCol);
          r.setBackground(i % 2 === 0 ? '#EBF3FB' : '#FFFFFF');
          r.setFontSize(10);
          logsSheet.setRowHeight(i, 26);
        }
      }
      for (let c = 1; c <= lLastCol; c++) logsSheet.autoResizeColumn(c);
    }
  }

  SpreadsheetApp.getUi().alert('✅ Sheet beautified!');
}

// ============================================================
// SAVE LOG
// ============================================================
function saveLog(ss, name, letterType, status, message, pdfUrl) {
  const logsSheet = ss.getSheetByName(CONFIG.sheets.logs);
  if (!logsSheet) return;
  logsSheet.appendRow([
    new Date(),
    name,
    letterType,
    status,
    message,
    pdfUrl || ''
  ]);
}

// ============================================================
// VIEW LOGS
// ============================================================
function viewLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logsSheet = ss.getSheetByName(CONFIG.sheets.logs);
  if (!logsSheet) {
    SpreadsheetApp.getUi().alert('Logs sheet not found!');
    return;
  }
  ss.setActiveSheet(logsSheet);
}

// ============================================================
// TEST LETTER BODY DETECTION
// ============================================================
function testLetterBodyDetection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheets.letterBody);
  const data = sheet.getDataRange().getValues();
  
  let msg = 'Letter Body Detection:\n\n';
  for (let i = 1; i < data.length; i++) {
    const type = data[i][0];
    const value = data[i][1].toString().trim();
    if (!type) continue;
    
    if (value.startsWith('https://docs.google.com')) {
      const docId = extractDocId(value);
      msg += `✅ ${type}\n   → Doc ID: ${docId}\n\n`;
    } else if (value) {
      msg += `📝 ${type}\n   → Plain text\n\n`;
    } else {
      msg += `❌ ${type}\n   → Empty!\n\n`;
    }
  }
  
  SpreadsheetApp.getUi().alert(msg);
}

// ============================================================
// GET OR CREATE PDF FOLDER
// ============================================================
function buildHtmlEmail(plainBody, letterType) {
  const lines = plainBody.split('\n').map(line => {
    if (!line.trim()) return '<br>';
    return `<p style="margin:0 0 10px 0;color:#374151;font-size:14px;line-height:1.7;">${line}</p>`;
  }).join('');

  return `<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f3f4f6;font-family:'Segoe UI',Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f3f4f6;padding:30px 0;">
<tr><td align="center">
<table width="600" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.08);">
  <tr>
    <td style="background:linear-gradient(135deg,#1a56db,#1e3a8a);padding:28px 36px;">
      <h1 style="margin:0;color:#fff;font-size:22px;font-weight:700;">Atoms Digital Solutions</h1>
      <p style="margin:4px 0 0;color:rgba(255,255,255,0.8);font-size:13px;">Guntur, Andhra Pradesh, India</p>
    </td>
  </tr>
  <tr>
    <td style="background:#eff6ff;padding:12px 36px;border-bottom:1px solid #dbeafe;">
      <span style="display:inline-block;background:#1a56db;color:#fff;font-size:12px;font-weight:600;padding:4px 14px;border-radius:20px;">${letterType}</span>
    </td>
  </tr>
  <tr><td style="padding:28px 36px;">${lines}</td></tr>
  <tr><td style="padding:0 36px;"><hr style="border:none;border-top:1px solid #e5e7eb;margin:0;"></td></tr>
  <tr>
    <td style="padding:18px 36px 26px;background:#f9fafb;text-align:center;">
      <p style="margin:0;font-size:12px;color:#9ca3af;line-height:1.6;">
        Official communication from <strong>Atoms Digital Solutions Pvt. Ltd.</strong><br>
        Please find the official letter attached as a PDF.
      </p>
    </td>
  </tr>
</table>
</td></tr>
</table>
</body></html>`;
}

function getOrCreatePdfFolder(letterType, year) {
  // Get or create HR Letters root folder
  let hrFolder;
  const hrFolders = DriveApp.getFoldersByName('HR Letters');
  if (hrFolders.hasNext()) {
    hrFolder = hrFolders.next();
  } else {
    hrFolder = DriveApp.createFolder('HR Letters');
  }

  // Get or create Letter Type folder inside HR Letters
  let letterFolder;
  const letterFolders = hrFolder.getFoldersByName(letterType);
  if (letterFolders.hasNext()) {
    letterFolder = letterFolders.next();
  } else {
    letterFolder = hrFolder.createFolder(letterType);
  }

  // Get or create Year folder inside Letter Type folder
  let yearFolder;
  const yearFolders = letterFolder.getFoldersByName(year);
  if (yearFolders.hasNext()) {
    yearFolder = yearFolders.next();
  } else {
    yearFolder = letterFolder.createFolder(year);
  }

  return yearFolder;
}

function setPendingStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CONFIG.sheets.data);
  const lastCol = dataSheet.getLastColumn();
  if (lastCol < 1) return;

  const headers = dataSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const statusCol = headers.indexOf('Status') + 1;
  if (statusCol === 0) return;

  const lastRow = dataSheet.getLastRow();
  for (let i = 2; i <= lastRow; i++) {
    const status = dataSheet.getRange(i, statusCol).getValue();
    if (!status) {
      dataSheet.getRange(i, statusCol).setValue('Pending');
    }
  }
}
