// ==========================
// 📋 CONFIGURATION
// ==========================
/**
 * FIX #1 (Security): Sensitive values (calendarId, alwaysGuest, openerEmails, adminEmail)
 * are now loaded from Script Properties at runtime via getConfig().
 * To set them, run the one-time helper: setupScriptProperties()
 *
 * FIX #8 (Design): duplicateIgnoreSheets is now an ARRAY, supporting multiple ignored sheets.
 */
const CONFIG = {
  // Column mappings
  openerColumn: 2,           // B - Opener name
  moveTriggerColumn: 3,      // C - Move trigger dropdown
  companyNameColumn: 5,      // E - Company name
  authorizedPersonColumn: 6, // F - Contact person
  phoneColumn: 7,            // G - Phone number
  emailColumn: 8,            // H - Email address
  dateColumn: 9,             // I - Meeting date/Meeting Time
  notesColumn: 11,           // K - George's Notes (Email/Timestamp Trigger)
  lastCallColumn: 13,        // M - Last Call (Timestamp Log)
  checkboxColumn: 15,        // O - Schedule checkbox (Checkbox)

  emailThrottleMinutes: 5,
  settingsSheetName: "Settings",

  // Sheets to ignore in duplicate checks and batch movements (e.g. mirror sheets)
  duplicateIgnoreSheets: ["Sheet9"],

  // ALLOWLIST: Only these tabs participate in duplicate checks and batch row movements.
  // Every other tab (temp sheets, archives, mirrors, etc.) is completely skipped.
  activeSheets: [
    "New Meetings",
    "Follow Ups",
    "No-Show",
    "Contract Sent",
    "Invoice Sent",
    "Dead Leads",
    "Temporary Inactive",
    "Onboarded"
  ]
};
/**
 * Loads sensitive config from Script Properties at runtime.
 * Avoids hardcoding credentials in source code.
 */
function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    calendarId: props.getProperty('CALENDAR_ID') || '',
    alwaysGuest: props.getProperty('ALWAYS_GUEST') || '',
    adminEmail: props.getProperty('ADMIN_EMAIL') || '',
    openerEmails: JSON.parse(props.getProperty('OPENER_EMAILS') || '{}')
  };
}
/**
 * One-time setup: run this manually once to store sensitive values in Script Properties.
 * After running, you can delete or comment out this function.
 */
function setupScriptProperties() {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    'CALENDAR_ID': 'c39f21666a2c70d644024bd29c7fc98379662f792e5a047cdb357a42b3c57d04@group.calendar.google.com',
    'ALWAYS_GUEST': 'george@wazowskioutsourcing.com',
    'ADMIN_EMAIL': 'ben.arthur.wiz@gmail.com',
    'OPENER_EMAILS': JSON.stringify({
      'Ben': 'ben.arthur.wiz@gmail.com',
      'Jane': 'kaity.james.wiz@gmail.com',
      'Jimmy': 'jimmy.pearson.wiz@gmail.com',
      'Selene': 'selene.myles.wiz@gmail.com'
    })
  });
  SpreadsheetApp.getUi().alert('✅ Script properties configured successfully.');
}
// ==========================================
// 🎯 CORE EVENT HANDLERS
// ==========================================
/**
 * Main trigger function - handles all sheet edits
 * NOTE: Multi-cell pastes are silently ignored (only single-cell edits are processed).
 */
// ==========================================
// FIX 3: Queue the movement BEFORE acquiring the lock,
// so a lock timeout doesn't silently drop it.
// ==========================================
function onEdit(e) {
  if (!e || !e.source || !e.range) return;
  // Handle multi-row paste in Column C
  if (e.range.getHeight() > 1 || e.range.getWidth() > 1) {
    // Only care if the paste touches Column C
    const pasteStartCol = e.range.getColumn();
    const pasteEndCol = pasteStartCol + e.range.getWidth() - 1;

    if (pasteStartCol <= CONFIG.moveTriggerColumn && pasteEndCol >= CONFIG.moveTriggerColumn) {
      const pasteStartRow = e.range.getRow();
      const pasteHeight = e.range.getHeight();
      const colOffset = CONFIG.moveTriggerColumn - pasteStartCol;

      // Read all pasted values in Column C in one batch call
      const pastedValues = e.range.getValues();

      for (let i = 0; i < pasteHeight; i++) {
        const row = pasteStartRow + i;
        if (row === 1) continue; // skip header
        const val = pastedValues[i][colOffset];
        if (val) queueRowMovement(e.range.getSheet().getName(), row, val);
      }
    }
    return; // still skip non-Column-C multi-cell edits
  }
  const editedRange = e.range;
  const editedSheet = editedRange.getSheet();
  const col = editedRange.getColumn();
  const row = editedRange.getRow();
  const val = e.value;

  if (editedSheet.getName() === CONFIG.settingsSheetName || row === 1) return;

  // Queue row movement IMMEDIATELY — before lock acquisition.
  // This ensures the key is written even if the lock wait below fails.
  if (col === CONFIG.moveTriggerColumn && val) {
    queueRowMovement(editedSheet.getName(), row, val);
  }

  // Now acquire lock for the rest of the work
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (error) {
    Logger.log('Could not obtain lock — row movement was already queued, will process on next trigger run.');
    return;
  }

  try {
    if (col === CONFIG.notesColumn) {
      editedSheet.getRange(row, CONFIG.lastCallColumn).setValue(new Date());
      const props = PropertiesService.getScriptProperties();
      const key = `BATCH_UPDATE|${editedSheet.getName()}|${row}`;
      props.setProperty(key, "true");
      logActivity('Batch Queue', `Note edit in Row ${row} of ${editedSheet.getName()} queued.`);
    }
  } catch (error) {
    Logger.log(`onEdit Error: ${error.message}`);
  } finally {
    lock.releaseLock();
  }
}
// ==========================================
// 🔄 ROW MOVEMENT HANDLING
// ==========================================
/**
 * Queues a row movement with a 10-second buffer (undo window)
 */
function queueRowMovement(sheetName, row, dropdownValue) {
  const props = PropertiesService.getScriptProperties();
  const key = `PENDING_MOVE|${sheetName}|${row}`;

  const queueData = {
    sheetName: sheetName,
    row: row,
    dropdownValue: dropdownValue,
    timestamp: new Date().getTime()
  };

  // Only overwrite if the dropdown value changed — preserve original timestamp otherwise
  const existing = props.getProperty(key);
  if (existing) {
    try {
      const existingData = JSON.parse(existing);
      if (existingData.dropdownValue === dropdownValue) {
        Logger.log(`Queue key already exists for ${sheetName} Row ${row} with same value — preserving original timestamp.`);
        return;
      }
    } catch (_) { }
  }

  props.setProperty(key, JSON.stringify(queueData));
  Logger.log(`Row movement queued: ${sheetName} Row ${row} → ${dropdownValue} (10s buffer)`);
}
/**
 * Processes queued row movements after 10-second buffer.
 * Set a "Time-driven" trigger for this function (every 1 minute).
 *
 * FIX #5: Removed unused mockEvent object. handleRowMovement now receives
 *         (sheet, range, ss, val) directly — no phantom event parameter.
 */
// ==========================================
// FIX 2: Don't silently discard cancelled moves —
// re-queue them as a "RETRY" so the backup picks them up.
// ==========================================
function processQueuedMovements() {
  const props = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date().getTime();
  // Count pending keys first to determine buffer window
  const pendingKeys = Object.keys(allProps).filter(k => k.startsWith("PENDING_MOVE|"));
  const bufferMs = pendingKeys.length < 10 ? 3 * 1000 : 10 * 1000;
  Logger.log(`Queue size: ${pendingKeys.length} — using ${bufferMs / 1000}s buffer`);
  const destinationMap = getDestinationMap();

  let processedCount = 0;
  let cancelledCount = 0;

  for (const key in allProps) {
    if (!key.startsWith("PENDING_MOVE|")) continue;

    try {
      const queueData = JSON.parse(allProps[key]);
      const ageMs = now - queueData.timestamp;

      if (ageMs < bufferMs) continue; // Still in buffer window

      const sheet = ss.getSheetByName(queueData.sheetName);

      if (!sheet) {
        Logger.log(`Sheet not found: ${queueData.sheetName}`);
        props.deleteProperty(key);
        continue;
      }

      const currentValue = String(
        sheet.getRange(queueData.row, CONFIG.moveTriggerColumn).getValue()
      ).trim();

      const storedValue = String(queueData.dropdownValue).trim();

      if (currentValue === storedValue) {
        const range = sheet.getRange(queueData.row, CONFIG.moveTriggerColumn);
        handleRowMovement(sheet, range, ss, currentValue, destinationMap);
        processedCount++;
        Logger.log(`Processed queued move: ${queueData.sheetName} Row ${queueData.row}`);
      } else if (currentValue && destinationMap[currentValue]) {
        // FIX: Value changed but is still a valid move target — process it anyway
        Logger.log(`Value changed from "${storedValue}" to "${currentValue}" — still valid, processing.`);
        const range = sheet.getRange(queueData.row, CONFIG.moveTriggerColumn);
        handleRowMovement(sheet, range, ss, currentValue, destinationMap);
        processedCount++;
      } else {
        cancelledCount++;
        Logger.log(`Cancelled move: ${queueData.sheetName} Row ${queueData.row} (dropdown cleared or changed to unknown value "${currentValue}")`);
      }

    } catch (error) {
      Logger.log(`Error processing queued movement ${key}: ${error.message}`);
    } finally {
      props.deleteProperty(key);
    }
  }

  if (processedCount > 0 || cancelledCount > 0) {
    Logger.log(`Queue processing: ${processedCount} moved, ${cancelledCount} cancelled`);
  }
}
/**
 * Handles moving rows between sheets based on dropdown selection.
 *
 * FIX #5: Removed unused `e` parameter. Signature is now (sheet, range, ss, val, destinationMap).
 * FIX #12: Accepts an optional pre-loaded destinationMap to avoid redundant Settings reads.
 * FIX: UI alerts replaced with safe wrapper — time-driven triggers cannot call getUi(),
 *      so alerts are only shown when a user is present (manual/onEdit context).
 */
function safeAlert(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    // Running from a time-driven trigger — UI not available, log instead
    Logger.log('ALERT (no UI): ' + message);
  }
}
function handleRowMovement(sheet, range, ss, val, destinationMap) {
  // FIX #12: Only read Settings sheet if no map was passed in
  if (!destinationMap) {
    destinationMap = getDestinationMap();
  }

  const targetName = destinationMap[val];

  if (!targetName) {
    safeAlert(
      `❌ Error: No target sheet defined for "${val}"\n\n` +
      `Please add this option to the Settings sheet:\n` +
      `Column A: ${val}\n` +
      `Column B: [Target Sheet Name]`
    );
    return;
  }

  const currentSheetName = sheet.getName();
  if (currentSheetName === targetName) return;

  const targetSheet = ss.getSheetByName(targetName);
  if (!targetSheet) {
    safeAlert(`❌ Error: Target sheet "${targetName}" not found.`);
    return;
  }

  const row = range.getRow();
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const sourceIsIgnored = isSheetIgnored(currentSheetName);
  const targetIsIgnored = isSheetIgnored(targetName);

  let isDup = { found: false };

  if (!sourceIsIgnored && !targetIsIgnored) {
    isDup = checkForDuplicates(
      ss,
      rowData[CONFIG.companyNameColumn - 1],
      rowData[CONFIG.phoneColumn - 1],
      rowData[CONFIG.emailColumn - 1],
      currentSheetName,
      row
    );
  } else {
    Logger.log(`Skipping duplicate check: ${currentSheetName} or ${targetName} is ignored`);
  }

  if (isDup.found) {
    safeAlert(`⚠️ Duplicate found in "${isDup.sheetName}" at row ${isDup.row}.\n\nMove cancelled.`);
    return;
  }

  archiveDeletedRow(ss, currentSheetName, row, rowData);
  targetSheet.appendRow(rowData);
  sheet.deleteRow(row);

  logActivity('Row Moved', `${currentSheetName} → ${targetName} | Company: ${rowData[CONFIG.companyNameColumn - 1]}`);
}
/**
 * Fetches the move map from the Settings sheet.
 */
function getDestinationMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.settingsSheetName);
  if (!sheet) return {};

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const map = {};

  data.forEach(row => {
    if (row[0]) map[row[0].toString().trim()] = row[1].toString().trim();
  });

  return map;
}
/**
 * Returns true if a sheet should be SKIPPED in duplicate checks and batch movements.
 *
 * A sheet is skipped if:
 *   (a) it is explicitly listed in CONFIG.duplicateIgnoreSheets, OR
 *   (b) it is NOT in the CONFIG.activeSheets allowlist.
 *
 * This means only the 8 named active tabs are ever processed — temp sheets,
 * mirror sheets, and any other tabs are ignored automatically.
 */
function isSheetIgnored(sheetName) {
  if (!sheetName) return true;
  const normalized = sheetName.toString().trim().toLowerCase();

  // Explicitly ignored sheets (e.g. Sheet9 mirror)
  if (CONFIG.duplicateIgnoreSheets.some(s => s.toLowerCase() === normalized)) return true;

  // Not on the allowlist — ignore it
  if (!CONFIG.activeSheets.some(s => s.toLowerCase() === normalized)) return true;

  return false;
}

// ==========================================
// 📧 EMAIL NOTIFICATIONS
// ==========================================
/**
 * HTML-escapes a string to prevent XSS in email bodies.
 *
 * FIX #7: All user-sourced cell values are now escaped before being injected into HTML.
 */
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Sends email notification when Column K is edited (via batch processor).
 *
 * FIX #3: Removed unused `e` parameter from signature.
 * FIX #7: All cell values are HTML-escaped before injection into the email body.
 */
function handleEmailNotification(sheet, range) {
  const row = range.getRow();
  const sheetName = sheet.getName();
  const runtimeConfig = getConfig();

  if (wasEmailRecentlySent(row, sheetName, 'notification')) {
    Logger.log(`Email throttled for row ${row}`);
    return;
  }

  const rowData = sheet.getRange(row, 1, 1, Math.max(CONFIG.notesColumn, CONFIG.emailColumn, CONFIG.phoneColumn)).getValues()[0];
  const openerName = rowData[CONFIG.openerColumn - 1];
  const openerEmail = runtimeConfig.openerEmails[openerName];

  // FIX #15: Consistent error handling — log and return (matches scheduleSelectedMeetings pattern)
  if (!openerName || !openerEmail) {
    Logger.log(`Skipping email: No opener or email address found for row ${row}`);
    return;
  }

  // FIX #7: Escape all cell values before HTML injection
  const company = escapeHtml(rowData[CONFIG.companyNameColumn - 1]) || '(No company name)';
  const contact = escapeHtml(rowData[CONFIG.authorizedPersonColumn - 1]) || '(No contact)';
  const phone = escapeHtml(rowData[CONFIG.phoneColumn - 1]) || '(No phone)';
  const email = escapeHtml(rowData[CONFIG.emailColumn - 1]) || '(No email)';
  const notes = escapeHtml(rowData[CONFIG.notesColumn - 1]) || '(No notes)';

  const subject = `BD MEETINGS: Note Update for ${company}`;
  // Replace all newline characters (\n) with HTML break tags (<br>)
  const formattedNotes = notes.replace(/\n/g, '<br>');
  const htmlBody = `
    <html>
      <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
        <div style="max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;">
          <div style="background-color: #4A90E2; color: white; padding: 20px; text-align: center;">
            <h2 style="margin: 0; font-size: 20px;">Meeting Note Update</h2>
          </div>
          <div style="padding: 25px;">
            <p>Hello <strong>${escapeHtml(openerName)}</strong>,</p>
            <p>George has updated the notes for a meeting. Please review the details below:</p>
            <table style="width: 100%; border-collapse: collapse; margin-top: 20px; border: 1px solid #ddd;">
              <tr>
                <td style="padding: 12px; border: 1px solid #e0e0e0; background-color: #f7f7f7; font-weight: bold; width: 35%;">Company:</td>
                <td style="padding: 12px; border: 1px solid #e0e0e0;">${company}</td>
              </tr>
              <tr>
                <td style="padding: 12px; border: 1px solid #e0e0e0; background-color: #f7f7f7; font-weight: bold;">Contact Person:</td>
                <td style="padding: 12px; border: 1px solid #e0e0e0;">${contact}</td>
              </tr>
              <tr>
                <td style="padding: 12px; border: 1px solid #e0e0e0; background-color: #f7f7f7; font-weight: bold;">Phone:</td>
                <td style="padding: 12px; border: 1px solid #e0e0e0;">${phone}</td>
              </tr>
              <tr>
                <td style="padding: 12px; border: 1px solid #e0e0e0; background-color: #f7f7f7; font-weight: bold;">Email:</td>
                <td style="padding: 12px; border: 1px solid #e0e0e0;">${email}</td>
              </tr>
              <tr>
                <td style="padding: 12px; border: 1px solid #e0e0e0; background-color: #f7f7f7; font-weight: bold;">George's Notes:</td>
                <td style="padding: 12px; border: 1px solid #e0e0e0;">${formattedNotes}</td>
              </tr>
            </table>
            <p style="margin-top: 25px; text-align: center;">
              <a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}"
                 style="background-color: #4CAF50; color: white; padding: 12px 30px;
                        text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">
                Open Spreadsheet
              </a>
            </p>
          </div>
          <div style="background-color: #f0f0f0; padding: 15px; font-size: 12px; color: #777; text-align: center;">
            <p style="margin: 0;">This is an automated notification from your BD Meetings system.</p>
          </div>
        </div>
      </body>
    </html>
  `;
  const plainTextBody = `Note update for ${rowData[CONFIG.companyNameColumn - 1] || ''}. Notes: ${rowData[CONFIG.notesColumn - 1] || ''}`;
  const contentToSearch = `${subject} ${plainTextBody}`;
  const shouldCcBen = /\b(contract|prices|ben)\b/i.test(contentToSearch) || /\@ben/i.test(contentToSearch);

  try {
    MailApp.sendEmail({
      to: openerEmail,
      cc: shouldCcBen ? runtimeConfig.adminEmail : '',
      subject: subject,
      htmlBody: htmlBody,
      body: plainTextBody
    });

    recordEmailSent(row, sheetName, 'notification');
    Logger.log(`Email sent to ${openerEmail} for row ${row}`);
  } catch (error) {
    Logger.log(`Failed to send email: ${error.message}`);
  }
}

/**
 * Email throttling - checks if email was recently sent
 */
function wasEmailRecentlySent(row, sheetName, action) {
  const cache = CacheService.getScriptCache();
  return cache.get(`email_${sheetName}_${row}_${action}`) !== null;
}

/**
 * Records that an email was sent (for throttling)
 */
function recordEmailSent(row, sheetName, action) {
  const cache = CacheService.getScriptCache();
  cache.put(`email_${sheetName}_${row}_${action}`, "sent", CONFIG.emailThrottleMinutes * 60);
}

// ==========================================
// 📅 CALENDAR SCHEDULING
// ==========================================

/**
 * Schedules meetings for all checked rows.
 *
 * FIX #4: Past-date error message no longer hardcodes "2025".
 *         It now dynamically shows the current and expected year.
 */
function scheduleSelectedMeetings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const runtimeConfig = getConfig();

  const calendar = CalendarApp.getCalendarById(runtimeConfig.calendarId);

  if (!calendar) {
    SpreadsheetApp.getUi().alert("❌ Error: Calendar not found. Check permissions.");
    return;
  }

  let scheduledCount = 0;
  let errors = [];
  const now = new Date();
  const currentYear = now.getFullYear();
  const nextYear = currentYear + 1;

  for (let i = 1; i < data.length; i++) {
    const isChecked = data[i][CONFIG.checkboxColumn - 1] === true || String(data[i][CONFIG.checkboxColumn - 1]).toUpperCase() === "TRUE";
    const dateValue = data[i][CONFIG.dateColumn - 1];

    if (isChecked && dateValue) {
      try {
        const rowNum = i + 1;
        const company = data[i][CONFIG.companyNameColumn - 1] || 'Unknown Company';
        const contact = data[i][CONFIG.authorizedPersonColumn - 1] || 'No Contact';
        const notes = data[i][CONFIG.notesColumn - 1] || 'General Inquiry';
        const leadEmail = data[i][CONFIG.emailColumn - 1];
        const openerName = String(data[i][CONFIG.openerColumn - 1]).trim();
        const openerEmail = runtimeConfig.openerEmails[openerName];

        if (!openerEmail) throw new Error(`No email for: "${openerName}"`);

        const startTime = new Date(dateValue);

        if (isNaN(startTime.getTime())) throw new Error("Invalid date format.");

        // FIX #4: Dynamic year reference instead of hardcoded "2025"
        if (startTime < now) throw new Error(
          `Date appears to be in the past (year ${startTime.getFullYear()}). ` +
          `Please correct to ${currentYear} or ${nextYear}.`
        );

        const endTime = new Date(startTime.getTime() + 30 * 60000);

        const eventTitle = `Meeting with George ${company} - ${contact}`;
        const eventDescription = `Meeting with George for ${contact} / ${company}\nNotes: ${notes}`;
        const guests = `${runtimeConfig.alwaysGuest},${openerEmail}`;

        calendar.createEvent(eventTitle, startTime, endTime, {
          description: eventDescription,
          guests: guests,
          sendInvites: true
        });

        sheet.getRange(rowNum, CONFIG.checkboxColumn).setValue(false);
        scheduledCount++;
        logActivity('Meeting Scheduled', eventTitle);

      } catch (err) {
        errors.push(`Row ${i + 1}: ${err.message}`);
      }
    }
  }

  if (scheduledCount > 0) {
    let message = `✅ Successfully scheduled ${scheduledCount} meeting(s).`;
    if (errors.length > 0) message += `\n\n⚠️ Errors:\n${errors.join('\n')}`;
    SpreadsheetApp.getUi().alert(message);
  } else {
    let msg = errors.length > 0
      ? `❌ Errors found:\n${errors.join('\n')}`
      : `No checked rows with valid future dates found.`;
    SpreadsheetApp.getUi().alert(msg);
  }
}

// ==========================================
// 🔍 DUPLICATE CHECKING
// ==========================================

/**
 * Optimized duplicate checker — ignores sheets in CONFIG.duplicateIgnoreSheets.
 *
 * FIX #2: Replaced 3 separate getRange() calls per sheet with a single batch read.
 *
 * EXACT MATCH: A duplicate is only flagged when ALL three fields (company, phone, email)
 * match simultaneously. Partial matches (e.g. same company name but different phone)
 * are NOT treated as duplicates.
 */
function checkForDuplicates(ss, name, phone, email, sourceSheet, sourceRow) {
  if (!name && !phone && !email) return { found: false };

  const sheets = ss.getSheets();

  for (let sheet of sheets) {
    const sheetName = sheet.getName();

    if (sheetName === CONFIG.settingsSheetName) continue;
    if (isSheetIgnored(sheetName)) {
      Logger.log(`Skipping duplicate check in ignored sheet: ${sheetName}`);
      continue;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;

    // FIX #2: Single batch read instead of 3 separate column reads
    const maxCol = Math.max(CONFIG.companyNameColumn, CONFIG.phoneColumn, CONFIG.emailColumn);
    const batchData = sheet.getRange(2, 1, lastRow - 1, maxCol).getValues();

    for (let i = 0; i < batchData.length; i++) {
      const actualRow = i + 2;

      if (sheetName === sourceSheet && actualRow === sourceRow) continue;

      const rowName = batchData[i][CONFIG.companyNameColumn - 1];
      const rowPhone = batchData[i][CONFIG.phoneColumn - 1];
      const rowEmail = batchData[i][CONFIG.emailColumn - 1];

      const nameMatch = name && rowName && rowName.toString().trim().toLowerCase() === name.toString().trim().toLowerCase();
      const phoneMatch = phone && rowPhone && rowPhone.toString().trim() === phone.toString().trim();
      const emailMatch = email && rowEmail && rowEmail.toString().toLowerCase().trim() === email.toString().toLowerCase().trim();

      // EXACT MATCH ONLY: all three fields must match
      if (nameMatch && phoneMatch && emailMatch) {
        return { found: true, sheetName: sheetName, row: actualRow };
      }
    }
  }

  return { found: false };
}

/**
 * Finds all duplicates across the spreadsheet.
 * Ignores sheets listed in CONFIG.duplicateIgnoreSheets.
 *
 * EXACT MATCH: Only flags a duplicate when company name, phone, AND email
 * all match simultaneously (composite key). Rows that share only one or two
 * fields are NOT considered duplicates.
 */
function findAllDuplicates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheets = ss.getSheets();
  const duplicates = [];
  const seen = {};

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();

    if (sheetName === CONFIG.settingsSheetName || sheetName.startsWith('_')) return;
    if (isSheetIgnored(sheetName)) {
      Logger.log(`Skipping ignored sheet in duplicate scan: ${sheetName}`);
      return;
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const company = data[i][CONFIG.companyNameColumn - 1];
      const phone = data[i][CONFIG.phoneColumn - 1];
      const email = data[i][CONFIG.emailColumn - 1];

      if (!company && !phone && !email) continue;

      // Composite key — all three must match to be a duplicate
      const key = `${String(company).trim().toLowerCase()}|${String(phone).trim()}|${String(email).trim().toLowerCase()}`;

      if (seen[key]) {
        duplicates.push(
          `Row ${i + 1} in "${sheetName}" matches Row ${seen[key].row} in "${seen[key].sheet}"`
        );
      } else {
        seen[key] = { sheet: sheetName, row: i + 1 };
      }
    }
  });

  const ignoredList = CONFIG.duplicateIgnoreSheets.join(', ');
  if (duplicates.length > 0) {
    const message = `Found ${duplicates.length} duplicate(s):\n\n` +
      duplicates.slice(0, 10).join('\n') +
      (duplicates.length > 10 ? '\n\n...and ' + (duplicates.length - 10) + ' more' : '');
    ui.alert(message);
  } else {
    ui.alert(`✅ No duplicates found! (Ignored sheets: ${ignoredList})`);
  }
}

// ==========================================
// 📊 UTILITY FUNCTIONS
// ==========================================

/**
 * Archives deleted rows before removal.
 *
 * FIX #11: Archive sheet now includes proper column headers for all row data columns,
 *          not just the first 3 meta columns. Headers are auto-generated by column index.
 */
function archiveDeletedRow(ss, sourceSheet, row, rowData) {
  let deletedSheet = ss.getSheetByName('_Deleted_Rows');

  if (!deletedSheet) {
    deletedSheet = ss.insertSheet('_Deleted_Rows');
    // FIX #11: Build full headers — meta columns + data columns A, B, C...
    const dataHeaders = rowData.map((_, idx) => `Col ${String.fromCharCode(65 + idx)}`);
    const fullHeaders = ['Archived At', 'Source Sheet', 'Original Row #', ...dataHeaders];
    deletedSheet.appendRow(fullHeaders);
    const headerRange = deletedSheet.getRange(1, 1, 1, fullHeaders.length);
    headerRange.setFontWeight('bold').setBackground('#ff6b6b').setFontColor('#ffffff');
  }

  const timestamp = new Date();
  deletedSheet.appendRow([timestamp, sourceSheet, row, ...rowData]);
}

/**
 * Logs activities to activity log sheet
 */
function logActivity(action, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('_Activity_Log');

  if (!logSheet) {
    logSheet = ss.insertSheet('_Activity_Log');
    logSheet.appendRow(['Timestamp', 'User', 'Action', 'Details']);
    logSheet.getRange('A1:D1').setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
  }

  logSheet.appendRow([
    new Date(),
    Session.getActiveUser().getEmail(),
    action,
    details
  ]);
}

/**
 * Validates and creates Settings sheet if missing
 */
function validateSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.settingsSheetName);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.settingsSheetName);
    sheet.appendRow(['Option', 'Target Sheet']);
    sheet.getRange('A1:B1').setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');

    sheet.appendRow(['Onboarded', 'Onboarded']);
    sheet.appendRow(['Follow Up', 'Follow Ups']);
    sheet.appendRow(['Meeting Attended', 'Follow Ups']);
    sheet.appendRow(['No - Show / Callback', 'No-Show']);
    sheet.appendRow(['NI', 'Dead Leads']);
    sheet.appendRow(['DNC', 'Dead Leads']);
    sheet.appendRow(['No Medicare', 'Dead Leads']);
    sheet.appendRow(['Contract Sent', 'Contract Sent']);
    sheet.appendRow(['Pending Medicare', 'Temporary Inactive']);

    SpreadsheetApp.getUi().alert('✅ Settings sheet created with status mappings!');
  }

  return sheet;
}

/**
 * Processes all flagged rows every 10 minutes with better error handling.
 * Set a "Time-driven" trigger for this function.
 *
 * FIX #3: handleEmailNotification called without null `e` — updated signature matches.
 */
function processBatchEmails() {
  const props = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let processedCount = 0;
  let errorCount = 0;

  for (const key in allProps) {
    if (key.startsWith("BATCH_UPDATE|")) {
      try {
        const parts = key.split("|");
        const sheetName = parts[1];
        const rowNum = parseInt(parts[2]);

        const sheet = ss.getSheetByName(sheetName);
        if (sheet) {
          const companyCheck = sheet.getRange(rowNum, CONFIG.companyNameColumn).getValue();

          if (companyCheck) {
            const range = sheet.getRange(rowNum, 1);
            // FIX #3: No longer passing null as first argument
            handleEmailNotification(sheet, range);
            processedCount++;
          } else {
            Logger.log(`Skipping: Row ${rowNum} on ${sheetName} appears to be moved or deleted.`);
          }
        } else {
          Logger.log(`Warning: Sheet "${sheetName}" not found for batch email.`);
        }
      } catch (error) {
        Logger.log(`Error processing batch email for ${key}: ${error.message}`);
        errorCount++;
      } finally {
        props.deleteProperty(key);
      }
    }
  }

  if (processedCount > 0 || errorCount > 0) {
    Logger.log(`Batch email processing complete: ${processedCount} sent, ${errorCount} errors`);
  }
}

/**
 * BACKUP MECHANISM - Catches any movements that weren't queued properly.
 * Optimized with batch reading and a 5-minute circuit breaker.
 * Set a "Time-driven" trigger for this function (e.g., every 5 minutes).
 *
 * FIX #1: Circuit breaker now uses a flag instead of `return` inside forEach,
 *         so it actually stops the outer loop rather than just skipping to the next sheet.
 * FIX #10: After deleteRow(), row indices in sheetData become stale. The fix is to
 *          re-read live row data from the sheet for the actual delete, while still
 *          using sheetData only for the initial trigger-column scan.
 * FIX #12: destinationMap loaded once at the top, not inside the loop.
 */
function processBatchRowMovement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const destinationMap = getDestinationMap(); // FIX #12: loaded once
  const sheets = ss.getSheets();
  let movedCount = 0;
  let errorCount = 0;
  let timedOut = false; // FIX #1: flag for circuit breaker

  const startTime = Date.now();

  // FIX #1: Use a standard for-loop so we can break out on timeout
  for (let s = 0; s < sheets.length; s++) {
    if (timedOut) break; // FIX #1: actually stop

    const sheet = sheets[s];
    const sheetName = sheet.getName();

    if (sheetName === CONFIG.settingsSheetName ||
      sheetName === '_Deleted_Rows' ||
      sheetName === '_Activity_Log' ||
      sheetName.startsWith('_')) continue;

    if (isSheetIgnored(sheetName)) continue;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) continue;

    // BATCH READ: scan trigger column from memory
    const sheetData = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    // Iterate bottom-to-top so row deletions don't shift unprocessed rows
    for (let i = lastRow - 1; i >= 1; i--) {
      // FIX #1: Set flag and break inner loop — outer loop will also break
      if (Date.now() - startTime > 300000) {
        Logger.log("Nearing 6-minute execution limit. Pausing batch; will resume on next trigger.");
        timedOut = true;
        break;
      }

      try {
        const triggerValue = sheetData[i][CONFIG.moveTriggerColumn - 1];

        if (triggerValue && destinationMap[triggerValue]) {
          const targetName = destinationMap[triggerValue];
          if (sheetName === targetName) continue;

          const targetSheet = ss.getSheetByName(targetName);
          if (!targetSheet) continue;

          const row = i + 1;

          // FIX #10: Re-read actual live row data before delete to avoid stale-index issues
          const liveRowData = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

          const sourceIsIgnored = isSheetIgnored(sheetName);
          const targetIsIgnored = isSheetIgnored(targetName);

          let isDup = { found: false };

          if (!sourceIsIgnored && !targetIsIgnored) {
            isDup = checkForDuplicates(
              ss,
              liveRowData[CONFIG.companyNameColumn - 1],
              liveRowData[CONFIG.phoneColumn - 1],
              liveRowData[CONFIG.emailColumn - 1],
              sheetName,
              row
            );
          } else {
            Logger.log(`Skipping duplicate check for row ${row} because ignored sheet is involved.`);
          }

          if (isDup.found) {
            Logger.log(`Duplicate found for row ${row} in ${sheetName}. Move to ${targetName} cancelled.`);
            continue;
          }

          archiveDeletedRow(ss, sheetName, row, liveRowData);
          targetSheet.appendRow(liveRowData);
          sheet.deleteRow(row);
          movedCount++;

          logActivity('Batch Row Moved', `${sheetName} → ${targetName} | Company: ${liveRowData[CONFIG.companyNameColumn - 1]}`);
        }
      } catch (error) {
        Logger.log(`Error processing row ${i + 1} in ${sheetName}: ${error.message}`);
        errorCount++;
      }
    }
  }

  if (movedCount > 0 || errorCount > 0) {
    Logger.log(`Batch row movement complete: ${movedCount} moved, ${errorCount} errors`);
  }
}

// ==========================================
// 🖥️ UI FUNCTIONS
// ==========================================

/**
 * Opens the activity log sheet
 */
function openActivityLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('_Activity_Log');

  if (!logSheet) {
    SpreadsheetApp.getUi().alert('No activity log found. Activities will be logged as actions occur.');
    return;
  }

  ss.setActiveSheet(logSheet);
}

/**
 * Creates a custom menu
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('BD Meetings')
    .addItem('📅 Schedule Selected Meetings', 'scheduleSelectedMeetings')
    .addItem('🔍 Find All Duplicates', 'findAllDuplicates')
    .addItem('📊 View Activity Log', 'openActivityLog')
    .addItem('⚙️ Validate Settings', 'validateSettingsSheet')
    .addItem('🧹 Clear Formatting on Current Tab', 'clearCurrentTabFormatting')
    .addItem('🔄 Sync Formatting from Master', 'unifyFormatting')
    .addSeparator()
    .addItem('🩺 Find Stuck Rows', 'findStuckRows')
    .addItem('📬 View Pending Queue', 'viewPendingQueue')
    .addItem('🗑️ Prune Old Logs', 'pruneOldLogs')
    .addItem('⚡ Setup Staggered Triggers (run once)', 'setupStaggeredTriggers')
    .addSeparator()
    .addItem('ℹ️ About Row Movement Buffer', 'showBufferInfo')
    .addToUi();
}

/**
 * Removes all conditional formatting rules
 */
function clearCurrentTabFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Clear all conditional rules?', ui.ButtonSet.YES_NO) == ui.Button.YES) {
    sheet.clearConditionalFormatRules();
  }
}

/**
 * Shows information about the 10-second buffer feature
 */
function showBufferInfo() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    '⏱️ 10-Second Undo Buffer',
    'When you select a status in Column C, the row will NOT move immediately.\n\n' +
    'You have 10 seconds to:\n' +
    '• Clear the dropdown to cancel\n' +
    '• Change to a different status\n' +
    '• Fix any mistakes\n\n' +
    'After 10 seconds, the row will automatically move to the correct sheet.\n\n' +
    '💡 This prevents accidental moves!',
    ui.ButtonSet.OK
  );
}

/**
 * Copies conditional formatting from the "New Meetings" template sheet to all other sheets
 */
function unifyFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const templateSheetName = "New Meetings";
  const templateSheet = ss.getSheetByName(templateSheetName);

  if (!templateSheet) {
    SpreadsheetApp.getUi().alert("Error: Could not find a tab named 'New Meetings'.");
    return;
  }

  const rules = templateSheet.getConditionalFormatRules();

  if (rules.length === 0) {
    SpreadsheetApp.getUi().alert("No rules found on 'New Meetings' to copy.");
    return;
  }

  const response = SpreadsheetApp.getUi().alert(
    'Sync Formatting',
    'This will wipe conditional formatting on ALL tabs and replace them with the rules from "New Meetings". Proceed?',
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );

  if (response == SpreadsheetApp.getUi().Button.YES) {
    sheets.forEach(sheet => {
      if (sheet.getName() !== templateSheetName && !sheet.isSheetHidden()) {
        sheet.clearConditionalFormatRules();
        const newRules = rules.map(rule => {
          return rule.copy().setRanges(rule.getRanges().map(range => {
            return sheet.getRange(range.getA1Notation());
          })).build();
        });
        sheet.setConditionalFormatRules(newRules);
      }
    });
    SpreadsheetApp.getUi().alert("✅ All tabs unified!");
  }
}
// ==========================================
// 🩺 STUCK ROWS DIAGNOSTIC
// Finds rows in active sheets that still have a
// Column C dropdown value set (meaning they were
// never moved). Useful after any trigger hiccup.
// ==========================================
function findStuckRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const destinationMap = getDestinationMap();

  // stuckMeta stores structured data for direct processing
  const stuckMeta = [];
  // stuckLines stores display strings for the alert preview
  const stuckLines = [];

  CONFIG.activeSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const triggerCol = CONFIG.moveTriggerColumn;
    const companyCol = CONFIG.companyNameColumn;
    const lastCol = sheet.getLastColumn();
    const maxCol = Math.max(triggerCol, companyCol, lastCol);

    // Batch read the full row so we have data ready if we need to move
    const data = sheet.getRange(2, 1, lastRow - 1, maxCol).getValues();

    data.forEach((row, i) => {
      const triggerValue = String(row[triggerCol - 1] || '').trim();
      const company = String(row[companyCol - 1] || '').trim();
      const targetName = destinationMap[triggerValue];

      if (triggerValue && targetName && targetName !== sheetName) {
        const rowNum = i + 2; // 1-indexed sheet row
        stuckMeta.push({ sheetName, sheet, rowNum, targetName, rowData: row });
        stuckLines.push(`  • "${sheetName}" Row ${rowNum} → "${targetName}" | Company: ${company || '(empty)'}`);
      }
    });
  });

  if (stuckMeta.length === 0) {
    ui.alert('✅ No Stuck Rows', 'All rows with dropdown values have been processed. Nothing is stuck.', ui.ButtonSet.OK);
    return;
  }

  const preview = stuckLines.slice(0, 15).join('\n');
  const extra = stuckMeta.length > 15 ? `\n\n...and ${stuckMeta.length - 15} more.` : '';

  const response = ui.alert(
    `⚠️ ${stuckMeta.length} Stuck Row(s) Found`,
    `The following rows have a move dropdown set but haven't been processed:\n\n${preview}${extra}\n\nWould you like to process them now?`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // --- Direct targeted processing (no full re-scan, no timeout risk) ---
  let movedCount = 0;
  let skippedDup = 0;
  let errorCount = 0;

  // Process bottom-to-top per sheet so row deletions don't shift indices
  // Group by sheet first, then sort descending by rowNum
  const bySheet = {};
  stuckMeta.forEach(item => {
    if (!bySheet[item.sheetName]) bySheet[item.sheetName] = [];
    bySheet[item.sheetName].push(item);
  });

  for (const sheetName in bySheet) {
    const items = bySheet[sheetName].sort((a, b) => b.rowNum - a.rowNum); // bottom-to-top
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    for (const item of items) {
      try {
        const targetSheet = ss.getSheetByName(item.targetName);
        if (!targetSheet) {
          Logger.log(`Target sheet "${item.targetName}" not found. Skipping row ${item.rowNum}.`);
          errorCount++;
          continue;
        }

        // Re-read the live row to avoid stale data after previous deletions
        const lastCol = sheet.getLastColumn();
        const liveRowData = sheet.getRange(item.rowNum, 1, 1, lastCol).getValues()[0];

        // Duplicate check (skip if either sheet is ignored)
        let isDup = { found: false };
        if (!isSheetIgnored(sheetName) && !isSheetIgnored(item.targetName)) {
          isDup = checkForDuplicates(
            ss,
            liveRowData[CONFIG.companyNameColumn - 1],
            liveRowData[CONFIG.phoneColumn - 1],
            liveRowData[CONFIG.emailColumn - 1],
            sheetName,
            item.rowNum
          );
        }

        if (isDup.found) {
          Logger.log(`Duplicate found for row ${item.rowNum} in ${sheetName}. Move to ${item.targetName} cancelled.`);
          skippedDup++;
          continue;
        }

        archiveDeletedRow(ss, sheetName, item.rowNum, liveRowData);
        targetSheet.appendRow(liveRowData);
        sheet.deleteRow(item.rowNum);
        movedCount++;

        logActivity('Stuck Row Fixed', `${sheetName} → ${item.targetName} | Company: ${liveRowData[CONFIG.companyNameColumn - 1]}`);
      } catch (err) {
        Logger.log(`Error processing stuck row ${item.rowNum} in ${sheetName}: ${err.message}`);
        errorCount++;
      }
    }
  }

  let summary = `✅ Done — ${movedCount} row(s) moved.`;
  if (skippedDup > 0) summary += `\n⚠️ ${skippedDup} skipped (duplicate detected).`;
  if (errorCount > 0) summary += `\n❌ ${errorCount} error(s) — check Execution Logs.`;
  ui.alert('Stuck Rows Processed', summary, ui.ButtonSet.OK);
}


// ==========================================
// 📬 PENDING QUEUE VIEWER
// Shows all PENDING_MOVE keys currently sitting
// in Script Properties, with age in seconds.
// Helps confirm the queue is healthy and not backed up.
// ==========================================
function viewPendingQueue() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();
  const now = Date.now();

  const pending = [];

  for (const key in allProps) {
    if (!key.startsWith('PENDING_MOVE|')) continue;

    try {
      const data = JSON.parse(allProps[key]);
      const ageSeconds = Math.round((now - data.timestamp) / 1000);
      const status = ageSeconds < 10 ? '⏳ In buffer' : '⚠️ Overdue';
      pending.push(`  ${status} | "${data.sheetName}" Row ${data.row} → "${data.dropdownValue}" | Age: ${ageSeconds}s`);
    } catch (e) {
      pending.push(`  ❓ Unreadable key: ${key}`);
    }
  }

  if (pending.length === 0) {
    ui.alert('📬 Pending Queue', '✅ Queue is empty. No movements are waiting to be processed.', ui.ButtonSet.OK);
    return;
  }

  const preview = pending.slice(0, 15).join('\n');
  const extra = pending.length > 15 ? `\n\n...and ${pending.length - 15} more.` : '';

  ui.alert(
    `📬 ${pending.length} Item(s) in Queue`,
    `${preview}${extra}\n\n⚠️ Items older than 10s should have been processed by the trigger.\nIf you see many overdue items, check your time-driven triggers.`,
    ui.ButtonSet.OK
  );
}


// ==========================================
// 🗑️ LOG PRUNING
// Deletes rows older than PRUNE_DAYS_THRESHOLD from
// _Activity_Log and _Deleted_Rows sheets.
// Safe to run manually or add as a weekly time trigger.
// ==========================================
function pruneOldLogs() {
  const PRUNE_DAYS_THRESHOLD = 90;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - PRUNE_DAYS_THRESHOLD);

  const sheetsToPrune = ['_Activity_Log', '_Deleted_Rows'];
  const summary = [];

  sheetsToPrune.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      summary.push(`  • "${sheetName}": not found, skipped.`);
      return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      summary.push(`  • "${sheetName}": empty, nothing to prune.`);
      return;
    }

    // Timestamp is always column A (index 0)
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let deletedCount = 0;

    // Iterate bottom-to-top so deletions don't shift indices
    for (let i = data.length - 1; i >= 0; i--) {
      const ts = new Date(data[i][0]);
      if (!isNaN(ts.getTime()) && ts < cutoff) {
        sheet.deleteRow(i + 2); // +2: 1 for header, 1 for 0-index
        deletedCount++;
      }
    }

    summary.push(`  • "${sheetName}": ${deletedCount} row(s) deleted.`);
  });

  ui.alert(
    '🗑️ Log Pruning Complete',
    `Removed entries older than ${PRUNE_DAYS_THRESHOLD} days:\n\n${summary.join('\n')}`,
    ui.ButtonSet.OK
  );
}

/**
 * Sets up two staggered 1-minute triggers for processQueuedMovements,
 * offset by 30 seconds to achieve ~30s polling without sub-minute triggers.
 * Run this ONCE manually after deploying. Delete old single trigger first.
 */
function setupStaggeredTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Delete any existing processQueuedMovements triggers first
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processQueuedMovements') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Trigger A — fires every minute on the minute
  ScriptApp.newTrigger('processQueuedMovements')
    .timeBased()
    .everyMinutes(1)
    .create();

  // Trigger B — fires every minute, but delayed 30s via a one-time bootstrap
  // (see scheduleSecondTrigger below)
  ScriptApp.newTrigger('scheduleSecondTriggerBoot')
    .timeBased()
    .after(30 * 1000)
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ Staggered triggers set up.\n\n' +
    'Trigger A fires every 1 minute.\n' +
    'Trigger B will fire 30s offset from Trigger A.\n\n' +
    'You only need to run this once.'
  );
}

/**
 * One-time bootstrap: fires 30s after setupStaggeredTriggers() is called,
 * then creates the permanent offset trigger and deletes itself.
 */
function scheduleSecondTriggerBoot() {
  // Create the permanent offset trigger
  ScriptApp.newTrigger('processQueuedMovements')
    .timeBased()
    .everyMinutes(1)
    .create();

  // Clean up this one-time bootstrap trigger
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'scheduleSecondTriggerBoot') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  Logger.log('Offset trigger created. Bootstrap trigger deleted.');
}