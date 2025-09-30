/**
 * Mail Merge Google Apps Script
 * - Draft / Send all or selected rows
 * - Schedule selected rows for Today or Tomorrow
 *
 * CONFIG: adjust constants below to match your sheet columns / preferences
 */

// ====== CONFIG - edit these ======
const RECIPIENT_COL   = "EMAIL";        // header name for recipient(s) (comma-separated allowed)
const EMAIL_SENT_COL  = "DRAFT DATE";   // header name to record status/date
const SUBJECT_COL     = "SUBJECT";      // header name for per-row subject (optional)
const DRAFT_SUBJECT   = "TEST";         // default subject used to locate Gmail draft template
const SCHEDULE_HOUR   = 9;              // default scheduled hour (24h)
const SCHEDULE_MINUTE = 30;             // default scheduled minute
// =====================================

/** Add custom menu to the active spreadsheet UI */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
    .addSubMenu(ui.createMenu('All Rows')
      .addItem('📝 Draft Emails', 'draftAllEmails')
      .addItem('📤 Send Emails Now', 'sendAllEmailsNow')
      .addItem('⏰ Schedule Send (Tomorrow 9 AM)', 'scheduleAllEmailsTomorrow'))
    .addSubMenu(ui.createMenu('Selected Rows')
      .addItem('📝 Draft Selected Emails', 'draftSelectedEmails')
      .addItem('📤 Send Selected Emails Now', 'sendSelectedEmailsNow')
      .addItem('⏰ Schedule Selected (Today)', 'scheduleSelectedEmailsToday')
      .addItem('⏰ Schedule Selected (Tomorrow)', 'scheduleSelectedEmailsTomorrow'))
    .addSeparator()
    .addItem('🗑️ Remove All Triggers', 'deleteAllTriggers')
    .addItem('📋 View Active Triggers', 'viewActiveTriggers')
    .addToUi();
}


// -------------------- All Rows --------------------

function draftAllEmails() {
  const subjectLine = Browser.inputBox("Draft Subject", "Made by ❤️ Manish Chaudhary", Browser.Buttons.OK_CANCEL);
  if (subjectLine === "cancel" || subjectLine == "") return;
  processEmails(subjectLine, 'draft', null);
  SpreadsheetApp.getUi().alert('✅ Drafts created successfully!');
}

function sendAllEmailsNow() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('⚠️ Confirm Send', 'Are you sure you want to SEND emails to all unprocessed rows immediately?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    const subjectLine = Browser.inputBox("Draft Subject", "Made by ❤️ Manish Chaudhary", Browser.Buttons.OK_CANCEL);
    if (subjectLine === "cancel" || subjectLine == "") return;
    processEmails(subjectLine, 'send', null);
    ui.alert('✅ Emails sent successfully!');
  }
}

function scheduleAllEmailsTomorrow() {
  createOneTimeTrigger('runScheduledAllEmails', "tomorrow");
  SpreadsheetApp.getUi().alert(`✅ Emails scheduled for tomorrow at ${String(SCHEDULE_HOUR).padStart(2,'0')}:${String(SCHEDULE_MINUTE).padStart(2,'0')}`);
}


// -------------------- Selected Rows --------------------

function draftSelectedEmails() {
  const selectedRows = getSelectedRowIndices();
  if (selectedRows.length === 0) { SpreadsheetApp.getUi().alert('⚠️ Please select at least one row to process.'); return; }
  const subjectLine = Browser.inputBox("Draft Subject", "Made by ❤️ Manish Chaudhary", Browser.Buttons.OK_CANCEL);
  if (subjectLine === "cancel" || subjectLine == "") return;
  processEmails(subjectLine, 'draft', selectedRows);
  SpreadsheetApp.getUi().alert(`✅ Drafts created for ${selectedRows.length} selected row(s)!`);
}

function sendSelectedEmailsNow() {
  const selectedRows = getSelectedRowIndices();
  if (selectedRows.length === 0) { SpreadsheetApp.getUi().alert('⚠️ Please select at least one row to process.'); return; }
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('⚠️ Confirm Send', `Are you sure you want to SEND emails to ${selectedRows.length} selected row(s) immediately?`, ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    const subjectLine = Browser.inputBox("Draft Subject", "Made by ❤️ Manish Chaudhary", Browser.Buttons.OK_CANCEL);
    if (subjectLine === "cancel" || subjectLine == "") return;
    processEmails(subjectLine, 'send', selectedRows);
    ui.alert(`✅ Emails sent to ${selectedRows.length} recipient(s)!`);
  }
}

// Schedule selected for TODAY at SCHEDULE_HOUR:SCHEDULE_MINUTE
function scheduleSelectedEmailsToday() {
  const selectedRows = getSelectedRowIndices();
  if (selectedRows.length === 0) { SpreadsheetApp.getUi().alert('⚠️ Please select at least one row to process.'); return; }
  PropertiesService.getScriptProperties().setProperty('scheduledRows', JSON.stringify(selectedRows));
  createOneTimeTrigger('runScheduledSelectedEmails', "today");
  SpreadsheetApp.getUi().alert(`✅ ${selectedRows.length} email(s) scheduled for today at ${String(SCHEDULE_HOUR).padStart(2,'0')}:${String(SCHEDULE_MINUTE).padStart(2,'0')}`);
}

// Schedule selected for TOMORROW at SCHEDULE_HOUR:SCHEDULE_MINUTE
function scheduleSelectedEmailsTomorrow() {
  const selectedRows = getSelectedRowIndices();
  if (selectedRows.length === 0) { SpreadsheetApp.getUi().alert('⚠️ Please select at least one row to process.'); return; }
  PropertiesService.getScriptProperties().setProperty('scheduledRows', JSON.stringify(selectedRows));
  createOneTimeTrigger('runScheduledSelectedEmails', "tomorrow");
  SpreadsheetApp.getUi().alert(`✅ ${selectedRows.length} email(s) scheduled for tomorrow at ${String(SCHEDULE_HOUR).padStart(2,'0')}:${String(SCHEDULE_MINUTE).padStart(2,'0')}`);
}


// -------------------- Core processing --------------------

/**
 * Process rows: create drafts or send email
 * - subjectLine: subject of the draft template to fetch (from Gmail Drafts)
 * - mode: 'draft' or 'send'
 * - selectedRows: array of 0-based indices into the data-array (null => all rows)
 */
function processEmails(subjectLine, mode = 'draft', selectedRows = null) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert('⚠️ No data rows found.');
    return;
  }
  const heads = data.shift(); // header row

  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
  const recipientColIdx = heads.indexOf(RECIPIENT_COL);
  const subjectColIdx = heads.indexOf(SUBJECT_COL);

  if (emailSentColIdx === -1 || recipientColIdx === -1 || subjectColIdx === -1) {
    throw new Error('Required columns not found. Please check column headers: ' + EMAIL_SENT_COL + ', ' + RECIPIENT_COL + ', ' + SUBJECT_COL);
  }

  // convert rows into objects keyed by header
  const obj = data.map(r => (heads.reduce((o,k,i) => { o[k] = r[i] || ''; return o; }, {})));
  const out = [];

  obj.forEach(function(row, index){
    // If selectedRows supplied, only process those indices (these indices are 0-based relative to data array)
    if (selectedRows !== null && !selectedRows.includes(index)) {
      out.push([row[EMAIL_SENT_COL]]);
      return;
    }

    if (!row[EMAIL_SENT_COL]) { // only process rows not previously processed
      try {
        // support multiple recipients in a single cell (comma-separated)
        let recipients = (row[RECIPIENT_COL] || "")
          .split(",")
          .map(e => e.trim())
          .filter(e => e.length > 0);

        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        const invalids = recipients.filter(e => !emailRegex.test(e));
        if (invalids.length > 0) {
          out.push([`Error: Invalid email(s) → "${invalids.join(", ")}"`]);
          return;
        }

        if (recipients.length === 0) {
          out.push([`Error: No recipient found`]);
          return;
        }

        const recipientString = recipients.join(",");

        // merge placeholders
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);
        const emailSubject = row[SUBJECT_COL] || emailTemplate.message.subject || DRAFT_SUBJECT;

        const options = {
          htmlBody: msgObj.html,
          attachments: emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages
        };

        if (mode === 'send') {
          GmailApp.sendEmail(recipientString, emailSubject, msgObj.text, options);
          out.push([`Sent: ${new Date().toLocaleString()}`]);
        } else {
          GmailApp.createDraft(recipientString, emailSubject, msgObj.text, options);
          out.push([`Draft: ${new Date().toLocaleString()}`]);
        }
      } catch (e) {
        out.push([`Error: ${e.message}`]);
      }
    } else {
      // keep existing status (so we don't overwrite)
      out.push([row[EMAIL_SENT_COL]]);
    }
  });

  // write back results starting at row 2 in the Email status column
  sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out);
}


// -------------------- Scheduler helpers --------------------

function runScheduledAllEmails() {
  processEmails(DRAFT_SUBJECT, 'send', null);
}

function runScheduledSelectedEmails() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const selectedRows = JSON.parse(scriptProperties.getProperty('scheduledRows') || '[]');
  if (selectedRows.length > 0) {
    processEmails(DRAFT_SUBJECT, 'send', selectedRows);
    scriptProperties.deleteProperty('scheduledRows');
  }
}

function createOneTimeTrigger(functionName, day = "today") {
  const triggerDate = new Date();
  triggerDate.setHours(SCHEDULE_HOUR, SCHEDULE_MINUTE, 0, 0);

  // If user selected "tomorrow" or the scheduled time has already passed for today => schedule next day
  if (day === "tomorrow" || triggerDate.getTime() <= new Date().getTime()) {
    triggerDate.setDate(triggerDate.getDate() + 1);
  }

  ScriptApp.newTrigger(functionName).timeBased().at(triggerDate).create();
}


// -------------------- Trigger management --------------------

function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  SpreadsheetApp.getUi().alert(`✅ Deleted ${triggers.length} trigger(s).`);
}

function viewActiveTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) { SpreadsheetApp.getUi().alert('No active triggers found.'); return; }
  let message = 'Active Triggers:\n\n';
  triggers.forEach((trigger,i) => {
    message += `${i+1}. Function: ${trigger.getHandlerFunction()}\n   Type: ${trigger.getEventType()}\n\n`;
  });
  SpreadsheetApp.getUi().alert(message);
}


// -------------------- Helper functions --------------------

/**
 * Return an array of 0-based indices (relative to data rows) for the user's selected rows.
 * Assumes first data row is row 2 (row 1 = headers).
 */
function getSelectedRowIndices() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const selection = sheet.getSelection();
  if (!selection) return [];
  const rangeList = selection.getActiveRangeList();
  if (!rangeList) return [];
  const ranges = rangeList.getRanges();

  const indices = [];
  ranges.forEach(range => {
    const startRow = range.getRow();
    const numRows = range.getNumRows();
    for (let i = 0; i < numRows; i++) {
      // convert sheet row number to 0-based data index (data starts at row 2)
      const rowIndex = startRow + i - 2;
      if (rowIndex >= 0) indices.push(rowIndex);
    }
  });
  // unique indices
  return [...new Set(indices)];
}

/**
 * Read a Gmail draft that exactly matches subject_line and return
 * { message: {subject, text, html}, attachments:[], inlineImages:{} }
 *
 * Notes:
 * - Attempts to retrieve inline attachments and map them to cids used in html body.
 * - Falls back gracefully if inline images can't be matched.
 */
function getGmailTemplateFromDrafts_(subject_line) {
  try {
    const drafts = GmailApp.getDrafts();
    const draft = drafts.filter(d => d.getMessage().getSubject() === subject_line)[0];
    if (!draft) throw new Error("Can't find Gmail draft with subject: " + subject_line);

    const msg = draft.getMessage();

    // get inline images and regular attachments (use options to separate)
    const allInlineImages = msg.getAttachments({includeInlineImages: true, includeAttachments: false}) || [];
    const attachments = msg.getAttachments({includeInlineImages: false}) || [];

    const htmlBody = msg.getBody();
    const textBody = msg.getPlainBody();

    // Build object mapping of inline images by name (file name) and by cid (if available).
    const imgObj = {};
    allInlineImages.forEach(img => {
      // key by name
      try { imgObj[img.getName()] = img; } catch (e) { /* continue */ }
    });

    // Extract cid references like <img src="cid:some-id" alt="name">
    const imgexp = /<img[\s\S]*?src=["']cid:([^"']+)["'][\s\S]*?(?:alt=["']([^"']+)["'])?[\s\S]*?>/g;
    const matches = [...htmlBody.matchAll(imgexp)];
    const inlineImagesObj = {};

    matches.forEach(match => {
      const cid = match[1];           // cid used in the HTML (image key)
      const alt  = match[2] || null;  // optional alt text (sometimes used as filename)

      // try to find a matching blob from attachments by several heuristics:
      // 1) exact cid === filename
      // 2) alt text (if present) as filename
      // 3) first available inline image (last resort)
      let blob = imgObj[cid] || (alt ? imgObj[alt] : null);
      if (!blob) {
        // try to match by comparing names ignoring common prefixes/suffixes
        const keys = Object.keys(imgObj);
        for (let k=0; k<keys.length && !blob; k++) {
          if (keys[k].indexOf(cid) !== -1 || cid.indexOf(keys[k]) !== -1) blob = imgObj[keys[k]];
        }
      }
      if (!blob && keys && keys.length > 0) blob = imgObj[keys[0]]; // last resort
      if (blob) inlineImagesObj[cid] = blob;
    });

    return {
      message: { subject: subject_line, text: textBody, html: htmlBody },
      attachments: attachments,
      inlineImages: inlineImagesObj
    };
  } catch (e) {
    throw new Error("Error getting template: " + e.message);
  }
}

/**
 * Replace placeholders {{FIELD}} in the template with values from 'data' object.
 * Template is an object: {text:..., html:...}
 */
function fillInTemplateFromObject_(template, data) {
  let template_string = JSON.stringify(template);
  template_string = template_string.replace(/{{\s*([^{}]+)\s*}}/g, (_, fieldName) => {
    // allow fallback to empty string if missing
    return escapeData_(data[fieldName] || "");
  });
  return JSON.parse(template_string);
}

function escapeData_(str) {
  return String(str || "");
}
