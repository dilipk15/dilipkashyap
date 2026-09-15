/**
 * AI Meeting Notes
 * ----------------------------------------------------------------------------
 * Free Otter.ai alternative for Google Workspace users. Watches your
 * "Google Meet" Drive folder for new transcripts, summarises them with
 * Gemini, drops a formatted Doc in "Meeting Notes (AI)/", and emails
 * the summary to attendees.
 *
 * REQUIRED: Google Workspace Business Standard or higher (for Meet transcripts).
 * REQUIRED: Free Gemini API key from https://aistudio.google.com/apikey
 *
 * How to install:
 *   1. Open a new Google Sheet (or any existing one).
 *   2. Extensions -> Apps Script. Delete the placeholder code.
 *   3. Paste this entire file. Save (Ctrl+S).
 *   4. Reload the Sheet tab.
 *   5. Menu -> "AI Meeting Notes" -> "One-Click Setup".
 *
 * Author: Dilip Kashyap. Free and open source. No warranty.
 */

// ============================================================================
// CONFIG
// ============================================================================

const CFG = {
  MEET_FOLDER_NAME: 'Google Meet',
  OUTPUT_FOLDER_NAME: 'Meeting Notes (AI)',
  SETTINGS_TAB: 'Settings',
  LOG_TAB: 'Log',
  TRIGGER_MINUTES: 15,
  MODEL: 'gemini-3.6-flash',
  API_BASE: 'https://generativelanguage.googleapis.com/v1beta/models',
  MAX_CHARS: 200000,
  MENU: 'AI Meeting Notes',
  PROCESSED_LIMIT: 500 // keep the last N processed file IDs to avoid re-processing
};

const P = {
  API_KEY: 'GEMINI_API_KEY',
  OUTPUT_FOLDER_ID: 'OUTPUT_FOLDER_ID',
  MEET_FOLDER_ID: 'MEET_FOLDER_ID',
  PROCESSED_IDS: 'PROCESSED_IDS',
  TRIGGER_ID: 'TRIGGER_ID',
  INSTALLED: 'INSTALLED',
  LAST_RUN: 'LAST_RUN'
};

const DEFAULT_SETTINGS = [
  ['Setting',                     'Value',    'Notes'],
  ['Email attendees',             'Yes',      'Yes = send summary email to everyone on the Calendar invite. No = only email yourself.'],
  ['Send HTML email',             'Yes',      'Yes = rich HTML. No = plain text.'],
  ['Include full transcript link','Yes',      'Yes = link to the original transcript Doc in the email footer.'],
  ['Meeting-type prompt add-on',  '',         'Optional extra instruction for Gemini (e.g. "Focus on client objections and next steps.").'],
  ['Skip meetings shorter than (minutes)', '3', 'Ignore any transcript representing a meeting shorter than this. Zero to disable.']
];

const LOG_HEADERS = ['Timestamp', 'Meeting Name', 'Transcript File', 'Notes Doc', 'Attendees', 'Status', 'Notes'];

// ============================================================================
// MENU
// ============================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(CFG.MENU)
    .addItem('One-Click Setup (First Time)', 'install')
    .addSeparator()
    .addItem('Set Gemini API Key', 'setApiKey')
    .addItem('Test Gemini Connection', 'testGemini')
    .addSeparator()
    .addItem('Process New Meetings Now', 'processNowManual')
    .addItem('Open Output Folder', 'openOutputFolder')
    .addSeparator()
    .addItem('Debug: Show What I See', 'debugScan')
    .addItem('Debug: Clear Processed History', 'debugClearProcessed')
    .addSeparator()
    .addItem('Uninstall (Remove Trigger)', 'uninstall')
    .addToUi();
}

// ============================================================================
// ONE-CLICK SETUP
// ============================================================================

function install() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  try {
    const confirm = ui.alert(
      'One-Click Setup',
      'This will:\n' +
      '  1. Create "Settings" and "Log" tabs in this sheet.\n' +
      '  2. Create the "Meeting Notes (AI)" folder in your Drive.\n' +
      '  3. Locate your existing "Google Meet" folder in Drive.\n' +
      '  4. Install an automatic trigger (runs every ' + CFG.TRIGGER_MINUTES + ' minutes).\n' +
      '  5. Ask for your free Gemini API key.\n\n' +
      'Ready?',
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;

    // Sheets
    ensureSheet_(CFG.SETTINGS_TAB);
    ensureSheet_(CFG.LOG_TAB);
    initSettingsTab_();
    initLogTab_();

    // Meet folder
    const meetFolder = findMeetFolder_();
    if (!meetFolder) {
      ui.alert(
        'Google Meet folder not found',
        'I could not find a folder named "Google Meet" in your Drive.\n\n' +
        'Google auto-creates this folder after your first Workspace Meet call with ' +
        'transcripts enabled (Business Standard or higher).\n\n' +
        'Run one short Meet call with transcripts on, wait 2 minutes, then re-run Setup.',
        ui.ButtonSet.OK
      );
      return;
    }
    props.setProperty(P.MEET_FOLDER_ID, meetFolder.getId());

    // Output folder
    const outputFolder = ensureFolder_(CFG.OUTPUT_FOLDER_NAME, null);
    props.setProperty(P.OUTPUT_FOLDER_ID, outputFolder.getId());

    // Trigger
    setupTrigger_();

    // API key
    props.setProperty(P.INSTALLED, 'true');
    const askKey = ui.alert(
      'Almost done',
      'Setup complete. Set your Gemini API key now?\n\n' +
      'Free at https://aistudio.google.com/apikey (30 seconds, no credit card).',
      ui.ButtonSet.YES_NO
    );
    if (askKey === ui.Button.YES) setApiKey();

    ui.alert(
      'You\'re all set',
      'AI Meeting Notes is now watching your Google Meet folder.\n\n' +
      'Every ' + CFG.TRIGGER_MINUTES + ' minutes it checks for new transcripts. When it finds one, it will:\n' +
      '  - Create a summary Doc in "Meeting Notes (AI)"\n' +
      '  - Email the summary to attendees\n' +
      '  - Log the row in this sheet\n\n' +
      'To test now: Menu -> "Process New Meetings Now".',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Setup failed', 'Error: ' + e.message, ui.ButtonSet.OK);
    console.error(e);
  }
}

function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const existing = props.getProperty(P.API_KEY);
  const label = existing ? 'A key is already set. Paste a new key to replace it (or Cancel to keep).' : 'Paste your Gemini API key:';
  const res = ui.prompt('Gemini API Key', label + '\n\n(Free at https://aistudio.google.com/apikey)', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const key = String(res.getResponseText() || '').trim();
  if (!key) return;
  if (!/^AIza[a-zA-Z0-9_-]{20,}$/.test(key)) {
    ui.alert('That does not look like a valid Gemini API key. Keys start with "AIza". Please try again.');
    return;
  }
  props.setProperty(P.API_KEY, key);
  ui.alert('Gemini API key saved.');
}

function testGemini() {
  const ui = SpreadsheetApp.getUi();
  const key = PropertiesService.getScriptProperties().getProperty(P.API_KEY);
  if (!key) {
    ui.alert('Set your Gemini API key first (Menu -> Set Gemini API Key).');
    return;
  }
  try {
    const out = callGemini_('Say the single word "connected" and nothing else.', {});
    const text = out && out.summary ? out.summary : JSON.stringify(out);
    ui.alert('Gemini responded:\n\n' + text.substring(0, 500));
  } catch (e) {
    ui.alert('Gemini test failed:\n\n' + e.message);
  }
}

function openOutputFolder() {
  const id = PropertiesService.getScriptProperties().getProperty(P.OUTPUT_FOLDER_ID);
  if (!id) return SpreadsheetApp.getUi().alert('Run Setup first.');
  const url = 'https://drive.google.com/drive/folders/' + id;
  const html = HtmlService.createHtmlOutput('<a href="' + url + '" target="_blank">Open Meeting Notes folder</a>').setHeight(60).setWidth(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Output Folder');
}

function uninstall() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert('Uninstall', 'This removes the automatic trigger. Your Sheet and Drive folder stay. Continue?', ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;
  removeAllTriggers_();
  PropertiesService.getScriptProperties().deleteProperty(P.TRIGGER_ID);
  ui.alert('Trigger removed. Re-run Setup any time to re-enable.');
}

// ============================================================================
// SHEET SETUP
// ============================================================================

function ensureSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function initSettingsTab_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SETTINGS_TAB);
  if (sh.getLastRow() > 0) return; // already initialised
  sh.getRange(1, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#0F766E').setFontColor('#ffffff');
  sh.setColumnWidth(1, 260);
  sh.setColumnWidth(2, 120);
  sh.setColumnWidth(3, 520);
  sh.setFrozenRows(1);
}

function initLogTab_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.LOG_TAB);
  if (sh.getLastRow() > 0) return;
  sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
  sh.getRange(1, 1, 1, LOG_HEADERS.length).setFontWeight('bold').setBackground('#0F766E').setFontColor('#ffffff');
  sh.setColumnWidths(1, LOG_HEADERS.length, 160);
  sh.setFrozenRows(1);
}

function getSetting_(name, fallback) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SETTINGS_TAB);
  if (!sh) return fallback;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toLowerCase() === String(name).trim().toLowerCase()) {
      const v = vals[i][1];
      return (v === '' || v === null) ? fallback : v;
    }
  }
  return fallback;
}

// ============================================================================
// TRIGGER
// ============================================================================

function setupTrigger_() {
  removeAllTriggers_();
  const t = ScriptApp.newTrigger('processNewMeetings')
    .timeBased()
    .everyMinutes(CFG.TRIGGER_MINUTES)
    .create();
  PropertiesService.getScriptProperties().setProperty(P.TRIGGER_ID, t.getUniqueId());
}

function removeAllTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'processNewMeetings') ScriptApp.deleteTrigger(t);
  });
}

// ============================================================================
// FOLDERS
// ============================================================================

function findMeetFolder_() {
  const it = DriveApp.getRootFolder().getFoldersByName(CFG.MEET_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  // Fallback: search entire Drive
  const q = DriveApp.getFoldersByName(CFG.MEET_FOLDER_NAME);
  return q.hasNext() ? q.next() : null;
}

function ensureFolder_(name, parent) {
  const root = parent || DriveApp.getRootFolder();
  const it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

// ============================================================================
// PROCESSING LOOP
// ============================================================================

function processNowManual() {
  const ui = SpreadsheetApp.getUi();
  try {
    const n = processNewMeetings();
    ui.alert('Done. Processed ' + n + ' new meeting(s). Check the Log tab.');
  } catch (e) {
    ui.alert('Error: ' + e.message);
    console.error(e);
  }
}

function processNewMeetings() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    console.log('Another run is in progress. Skipping.');
    return 0;
  }
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty(P.INSTALLED) !== 'true') throw new Error('Not installed. Run Setup first.');
    const key = props.getProperty(P.API_KEY);
    if (!key) throw new Error('Gemini API key missing. Menu -> Set Gemini API Key.');

    const meetFolderId = props.getProperty(P.MEET_FOLDER_ID);
    const outputFolderId = props.getProperty(P.OUTPUT_FOLDER_ID);
    if (!meetFolderId || !outputFolderId) throw new Error('Folders missing. Re-run Setup.');

    const meetFolder = DriveApp.getFolderById(meetFolderId);
    const outputFolder = DriveApp.getFolderById(outputFolderId);

    const processed = getProcessedIds_();
    const transcripts = collectNewTranscripts_(meetFolder, processed);

    let count = 0;
    for (let i = 0; i < transcripts.length; i++) {
      try {
        processOneTranscript_(transcripts[i], outputFolder);
        markProcessed_(transcripts[i].file.getId());
        count++;
      } catch (e) {
        console.error('Failed to process ' + transcripts[i].file.getName() + ': ' + e.message);
        const retryable = /\[RETRYABLE\]/.test(e.message);
        logRow_({
          meetingName: transcripts[i].meetingName || transcripts[i].file.getName(),
          transcriptUrl: transcripts[i].file.getUrl(),
          docUrl: '',
          attendees: '',
          status: retryable ? 'RETRY' : 'ERROR',
          notes: e.message
        });
        // Retryable = transient (Gemini 5xx/429): leave unprocessed, the next 15-min run will try again.
        // Non-retryable = permanent (bad transcript, quota, key): mark processed so it doesn't loop forever.
        if (!retryable) markProcessed_(transcripts[i].file.getId());
      }
    }

    props.setProperty(P.LAST_RUN, new Date().toISOString());
    return count;
  } finally {
    lock.releaseLock();
  }
}

function collectNewTranscripts_(meetFolder, processedSet) {
  const results = [];
  // Meet organises transcripts per-meeting in subfolders
  const subs = meetFolder.getFolders();
  while (subs.hasNext()) {
    const sub = subs.next();
    const files = sub.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      const name = f.getName();
      if (!/transcript/i.test(name)) continue;
      if (f.getMimeType() !== MimeType.GOOGLE_DOCS) continue;
      if (processedSet[f.getId()]) continue;
      results.push({
        file: f,
        parentFolder: sub,
        meetingName: cleanMeetingName_(sub.getName()),
        createdAt: f.getDateCreated()
      });
    }
  }
  // Also check root of Meet folder in case the auto-organisation isn't in place
  const rootFiles = meetFolder.getFiles();
  while (rootFiles.hasNext()) {
    const f = rootFiles.next();
    if (!/transcript/i.test(f.getName())) continue;
    if (f.getMimeType() !== MimeType.GOOGLE_DOCS) continue;
    if (processedSet[f.getId()]) continue;
    results.push({
      file: f,
      parentFolder: meetFolder,
      meetingName: cleanMeetingName_(f.getName()),
      createdAt: f.getDateCreated()
    });
  }
  return results;
}

function cleanMeetingName_(raw) {
  return String(raw || '')
    .replace(/^Transcript\s*[:\-]?\s*/i, '')
    .replace(/\(\d{4}-\d{2}-\d{2}[^)]*\)/g, '')
    .replace(/\d{4}[/-]\d{2}[/-]\d{2}.*$/, '')
    .replace(/\s{2,}/g, ' ')
    .trim() || 'Meeting';
}

function processOneTranscript_(t, outputFolder) {
  const text = readDocText_(t.file.getId());
  if (!text || text.length < 200) throw new Error('Transcript is empty or too short.');

  const skipMin = Number(getSetting_('Skip meetings shorter than (minutes)', 3)) || 0;
  const minChars = skipMin * 400; // rough proxy: 400 chars per spoken minute
  if (skipMin > 0 && text.length < minChars) {
    logRow_({
      meetingName: t.meetingName,
      transcriptUrl: t.file.getUrl(),
      docUrl: '',
      attendees: '',
      status: 'SKIPPED',
      notes: 'Under ' + skipMin + ' min threshold'
    });
    return;
  }

  const context = {
    meetingName: t.meetingName,
    date: t.createdAt,
    calendarEvent: findCalendarEvent_(t.createdAt, t.meetingName)
  };

  const transcript = text.substring(0, CFG.MAX_CHARS);
  const summary = callGemini_(transcript, context);

  const doc = createNotesDoc_(summary, context, t.file.getUrl(), outputFolder);
  const attendees = extractAttendees_(context.calendarEvent);

  const emailSetting = String(getSetting_('Email attendees', 'Yes')).trim().toLowerCase();
  const recipients = (emailSetting === 'yes' && attendees.length) ? attendees : [Session.getActiveUser().getEmail()];
  emailSummary_(recipients, summary, context, doc, t.file.getUrl());

  logRow_({
    meetingName: t.meetingName,
    transcriptUrl: t.file.getUrl(),
    docUrl: doc.getUrl(),
    attendees: recipients.join(', '),
    status: 'OK',
    notes: 'Summary: ' + (summary.summary || '').substring(0, 80) + '...'
  });
}

// ============================================================================
// TRANSCRIPT READING
// ============================================================================

function readDocText_(fileId) {
  const doc = DocumentApp.openById(fileId);
  return doc.getBody().getText();
}

// ============================================================================
// CALENDAR MATCHING
// ============================================================================

function findCalendarEvent_(when, meetingName) {
  try {
    const day = new Date(when);
    const start = new Date(day.getFullYear(), day.getMonth(), day.getDate(), 0, 0, 0);
    const end = new Date(day.getFullYear(), day.getMonth(), day.getDate(), 23, 59, 59);
    const cal = CalendarApp.getDefaultCalendar();
    const events = cal.getEvents(start, end);
    if (!events.length) return null;

    // Prefer events with a Meet link and a matching title
    let best = null;
    let bestScore = -1;
    for (let i = 0; i < events.length; i++) {
      const ev = events[i];
      const title = String(ev.getTitle() || '').toLowerCase();
      const target = String(meetingName || '').toLowerCase();
      let score = 0;
      if (title && target && (title.indexOf(target) >= 0 || target.indexOf(title) >= 0)) score += 5;
      const hasMeet = /meet\.google\.com/.test(ev.getDescription() || '') || /meet\.google\.com/.test(ev.getLocation() || '');
      if (hasMeet) score += 3;
      const timeDiff = Math.abs(ev.getEndTime().getTime() - when.getTime()) / 60000; // minutes
      if (timeDiff < 120) score += 3;
      else if (timeDiff < 360) score += 1;
      if (score > bestScore) { bestScore = score; best = ev; }
    }
    return bestScore > 3 ? best : null;
  } catch (e) {
    console.warn('Calendar lookup failed: ' + e.message);
    return null;
  }
}

function extractAttendees_(event) {
  if (!event) return [];
  try {
    const guests = event.getGuestList();
    const out = [];
    for (let i = 0; i < guests.length; i++) {
      const em = guests[i].getEmail();
      if (em && out.indexOf(em) < 0) out.push(em);
    }
    // include organizer
    try {
      const owner = event.getCreators();
      if (owner && owner[0] && out.indexOf(owner[0]) < 0) out.push(owner[0]);
    } catch (_) {}
    return out;
  } catch (e) {
    return [];
  }
}

// ============================================================================
// GEMINI
// ============================================================================

function callGemini_(transcript, context) {
  const key = PropertiesService.getScriptProperties().getProperty(P.API_KEY);
  if (!key) throw new Error('Gemini API key missing.');

  const promptAddOn = getSetting_('Meeting-type prompt add-on', '');
  const meetingName = context && context.meetingName ? context.meetingName : 'this meeting';
  const dateStr = context && context.date ? Utilities.formatDate(context.date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '';

  const systemPrompt =
    'You produce structured summaries of business meeting transcripts. ' +
    'Return valid JSON with this schema:\n' +
    '{\n' +
    '  "summary": "3-4 sentence overview in plain language, no marketing tone",\n' +
    '  "action_items": [{"item": "...", "owner": "name or Unassigned", "due": "date if mentioned, else empty"}],\n' +
    '  "decisions": ["one line per decision made"],\n' +
    '  "follow_up_questions": ["open questions raised but not answered"],\n' +
    '  "highlights": ["3-6 short quotes worth remembering, with the speaker in brackets"]\n' +
    '}\n' +
    'Rules:\n' +
    '- Never invent action items or decisions that are not in the transcript.\n' +
    '- If a section is empty, return an empty array.\n' +
    '- Do not include commentary. Return the JSON object only.\n' +
    (promptAddOn ? '\nAdditional focus: ' + promptAddOn : '');

  const userPrompt =
    'Meeting: ' + meetingName + '\n' +
    (dateStr ? 'Date: ' + dateStr + '\n' : '') +
    '\nTranscript:\n' + transcript;

  const body = {
    contents: [{ role: 'user', parts: [{ text: userPrompt }] }],
    systemInstruction: { role: 'system', parts: [{ text: systemPrompt }] },
    generationConfig: {
      temperature: 0.2,
      responseMimeType: 'application/json'
    }
  };

  const url = CFG.API_BASE + '/' + CFG.MODEL + ':generateContent?key=' + encodeURIComponent(key);
  const RETRYABLE = { 429: true, 500: true, 502: true, 503: true, 504: true };
  const MAX_ATTEMPTS = 4;
  let code, raw;
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    code = res.getResponseCode();
    raw = res.getContentText();
    if (code >= 200 && code < 300) break;
    if (!RETRYABLE[code] || attempt === MAX_ATTEMPTS) {
      const marker = RETRYABLE[code] ? '[RETRYABLE]' : '';
      throw new Error('Gemini API returned ' + code + ' ' + marker + ': ' + raw.substring(0, 400));
    }
    const wait = Math.pow(2, attempt) * 1000; // 2s, 4s, 8s
    console.log('Gemini ' + code + '. Retry ' + attempt + '/' + (MAX_ATTEMPTS - 1) + ' in ' + wait + 'ms.');
    Utilities.sleep(wait);
  }

  let parsed;
  try { parsed = JSON.parse(raw); } catch (e) { throw new Error('Invalid JSON from Gemini API: ' + raw.substring(0, 300)); }
  const candidateText = parsed && parsed.candidates && parsed.candidates[0] && parsed.candidates[0].content && parsed.candidates[0].content.parts && parsed.candidates[0].content.parts[0] && parsed.candidates[0].content.parts[0].text;
  if (!candidateText) throw new Error('Empty response from Gemini.');

  let summaryObj;
  try {
    summaryObj = JSON.parse(candidateText);
  } catch (e) {
    // sometimes the model wraps in ```json ... ```
    const stripped = candidateText.replace(/```json\s*|\s*```/g, '').trim();
    try { summaryObj = JSON.parse(stripped); }
    catch (_) { throw new Error('Gemini returned non-JSON summary: ' + candidateText.substring(0, 300)); }
  }
  // normalise
  summaryObj.summary = summaryObj.summary || '';
  summaryObj.action_items = Array.isArray(summaryObj.action_items) ? summaryObj.action_items : [];
  summaryObj.decisions = Array.isArray(summaryObj.decisions) ? summaryObj.decisions : [];
  summaryObj.follow_up_questions = Array.isArray(summaryObj.follow_up_questions) ? summaryObj.follow_up_questions : [];
  summaryObj.highlights = Array.isArray(summaryObj.highlights) ? summaryObj.highlights : [];
  return summaryObj;
}

// ============================================================================
// OUTPUT DOC
// ============================================================================

function createNotesDoc_(summary, context, transcriptUrl, outputFolder) {
  const dateStr = context.date ? Utilities.formatDate(context.date, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
  const title = 'Notes - ' + (context.meetingName || 'Meeting') + (dateStr ? ' (' + dateStr + ')' : '');
  const doc = DocumentApp.create(title);
  const body = doc.getBody();

  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  if (dateStr) body.appendParagraph('Meeting date: ' + dateStr).setForegroundColor('#5b6b69');

  body.appendParagraph('Summary').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(summary.summary || '(no summary produced)');

  if (summary.action_items && summary.action_items.length) {
    body.appendParagraph('Action items').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    for (let i = 0; i < summary.action_items.length; i++) {
      const a = summary.action_items[i];
      const line = '- ' + (a.item || '') + '  [Owner: ' + (a.owner || 'Unassigned') + (a.due ? ', Due: ' + a.due : '') + ']';
      body.appendListItem(line).setGlyphType(DocumentApp.GlyphType.BULLET);
    }
  }

  if (summary.decisions && summary.decisions.length) {
    body.appendParagraph('Decisions').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    for (let i = 0; i < summary.decisions.length; i++) {
      body.appendListItem(summary.decisions[i]).setGlyphType(DocumentApp.GlyphType.BULLET);
    }
  }

  if (summary.follow_up_questions && summary.follow_up_questions.length) {
    body.appendParagraph('Open questions').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    for (let i = 0; i < summary.follow_up_questions.length; i++) {
      body.appendListItem(summary.follow_up_questions[i]).setGlyphType(DocumentApp.GlyphType.BULLET);
    }
  }

  if (summary.highlights && summary.highlights.length) {
    body.appendParagraph('Highlights').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    for (let i = 0; i < summary.highlights.length; i++) {
      body.appendListItem(summary.highlights[i]).setGlyphType(DocumentApp.GlyphType.BULLET);
    }
  }

  const includeTranscript = String(getSetting_('Include full transcript link', 'Yes')).trim().toLowerCase() === 'yes';
  if (includeTranscript && transcriptUrl) {
    body.appendParagraph('');
    const p = body.appendParagraph('Original transcript');
    p.setLinkUrl(transcriptUrl);
    p.setForegroundColor('#0F766E');
  }

  body.appendParagraph('');
  const footer = body.appendParagraph('Generated by AI Meeting Notes for Google Workspace.');
  footer.setForegroundColor('#9aa8a6').setItalic(true).setFontSize(9);

  doc.saveAndClose();
  // move to output folder
  const file = DriveApp.getFileById(doc.getId());
  outputFolder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  return doc;
}

// ============================================================================
// EMAIL
// ============================================================================

function emailSummary_(recipients, summary, context, doc, transcriptUrl) {
  if (!recipients || !recipients.length) return;
  const subject = 'Notes: ' + (context.meetingName || 'Meeting');
  const html = String(getSetting_('Send HTML email', 'Yes')).trim().toLowerCase() === 'yes';

  const plain = buildPlainEmail_(summary, context, doc, transcriptUrl);
  const opts = { name: 'AI Meeting Notes' };
  if (html) opts.htmlBody = buildHtmlEmail_(summary, context, doc, transcriptUrl);

  GmailApp.sendEmail(recipients.join(','), subject, plain, opts);
}

function buildPlainEmail_(summary, context, doc, transcriptUrl) {
  const lines = [];
  lines.push('Notes for: ' + (context.meetingName || 'Meeting'));
  if (context.date) lines.push('Date: ' + Utilities.formatDate(context.date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
  lines.push('');
  lines.push('SUMMARY');
  lines.push(summary.summary || '(no summary)');
  lines.push('');
  if (summary.action_items && summary.action_items.length) {
    lines.push('ACTION ITEMS');
    summary.action_items.forEach(function (a) {
      lines.push('- ' + (a.item || '') + '  [Owner: ' + (a.owner || 'Unassigned') + (a.due ? ', Due: ' + a.due : '') + ']');
    });
    lines.push('');
  }
  if (summary.decisions && summary.decisions.length) {
    lines.push('DECISIONS');
    summary.decisions.forEach(function (d) { lines.push('- ' + d); });
    lines.push('');
  }
  if (summary.follow_up_questions && summary.follow_up_questions.length) {
    lines.push('OPEN QUESTIONS');
    summary.follow_up_questions.forEach(function (q) { lines.push('- ' + q); });
    lines.push('');
  }
  lines.push('Full notes: ' + doc.getUrl());
  if (transcriptUrl) lines.push('Original transcript: ' + transcriptUrl);
  lines.push('');
  lines.push('-- ');
  lines.push('Generated by AI Meeting Notes (Apps Script + Gemini).');
  return lines.join('\n');
}

function buildHtmlEmail_(summary, context, doc, transcriptUrl) {
  const esc = function (s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  const date = context.date ? Utilities.formatDate(context.date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '';
  let html = '';
  html += '<div style="font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#1a2b29;line-height:1.55;max-width:640px">';
  html += '<h2 style="margin:0 0 4px;color:#0F766E">' + esc(context.meetingName || 'Meeting') + '</h2>';
  if (date) html += '<div style="color:#5b6b69;font-size:13px;margin-bottom:20px">' + esc(date) + '</div>';

  html += '<h3 style="margin:22px 0 6px;font-size:15px;color:#0F766E">Summary</h3>';
  html += '<p style="margin:0 0 16px">' + esc(summary.summary || '(no summary)') + '</p>';

  if (summary.action_items && summary.action_items.length) {
    html += '<h3 style="margin:22px 0 6px;font-size:15px;color:#0F766E">Action items</h3><ul style="margin:0 0 16px;padding-left:20px">';
    summary.action_items.forEach(function (a) {
      html += '<li style="margin-bottom:6px">' + esc(a.item || '') +
        ' <span style="color:#5b6b69;font-size:13px">[Owner: ' + esc(a.owner || 'Unassigned') + (a.due ? ', Due: ' + esc(a.due) : '') + ']</span></li>';
    });
    html += '</ul>';
  }
  if (summary.decisions && summary.decisions.length) {
    html += '<h3 style="margin:22px 0 6px;font-size:15px;color:#0F766E">Decisions</h3><ul style="margin:0 0 16px;padding-left:20px">';
    summary.decisions.forEach(function (d) { html += '<li style="margin-bottom:6px">' + esc(d) + '</li>'; });
    html += '</ul>';
  }
  if (summary.follow_up_questions && summary.follow_up_questions.length) {
    html += '<h3 style="margin:22px 0 6px;font-size:15px;color:#0F766E">Open questions</h3><ul style="margin:0 0 16px;padding-left:20px">';
    summary.follow_up_questions.forEach(function (q) { html += '<li style="margin-bottom:6px">' + esc(q) + '</li>'; });
    html += '</ul>';
  }
  html += '<div style="margin-top:24px;padding-top:16px;border-top:1px solid #e3e8e7;font-size:13px;color:#5b6b69">';
  html += '<a href="' + esc(doc.getUrl()) + '" style="color:#0F766E;text-decoration:none">Open full notes</a>';
  if (transcriptUrl) html += '  •  <a href="' + esc(transcriptUrl) + '" style="color:#0F766E;text-decoration:none">Original transcript</a>';
  html += '</div>';
  html += '<div style="margin-top:8px;color:#9aa8a6;font-size:11px;font-style:italic">Generated by AI Meeting Notes (Apps Script + Gemini).</div>';
  html += '</div>';
  return html;
}

// ============================================================================
// LOG + PROCESSED IDS
// ============================================================================

function logRow_(row) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.LOG_TAB);
  if (!sh) return;
  sh.appendRow([
    new Date(),
    row.meetingName || '',
    row.transcriptUrl || '',
    row.docUrl || '',
    row.attendees || '',
    row.status || '',
    row.notes || ''
  ]);
}

function getProcessedIds_() {
  const raw = PropertiesService.getScriptProperties().getProperty(P.PROCESSED_IDS) || '';
  const arr = raw ? raw.split(',') : [];
  const map = {};
  for (let i = 0; i < arr.length; i++) map[arr[i]] = true;
  return map;
}

function markProcessed_(fileId) {
  const raw = PropertiesService.getScriptProperties().getProperty(P.PROCESSED_IDS) || '';
  const arr = raw ? raw.split(',') : [];
  arr.push(fileId);
  const trimmed = arr.length > CFG.PROCESSED_LIMIT ? arr.slice(arr.length - CFG.PROCESSED_LIMIT) : arr;
  PropertiesService.getScriptProperties().setProperty(P.PROCESSED_IDS, trimmed.join(','));
}

// ============================================================================
// DEBUG
// ============================================================================

function debugScan() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const meetFolderId = props.getProperty(P.MEET_FOLDER_ID);
  const lines = [];

  if (!meetFolderId) {
    ui.alert('Meet folder ID not stored. Run Setup first.');
    return;
  }

  let meetFolder;
  try {
    meetFolder = DriveApp.getFolderById(meetFolderId);
  } catch (e) {
    ui.alert('Stored Meet folder ID is invalid: ' + e.message);
    return;
  }

  lines.push('Meet folder: "' + meetFolder.getName() + '"');
  lines.push('Folder URL: https://drive.google.com/drive/folders/' + meetFolderId);
  lines.push('');

  // Files directly in the Meet folder
  lines.push('== Files directly in "' + meetFolder.getName() + '" ==');
  const rootFiles = meetFolder.getFiles();
  let rootCount = 0;
  while (rootFiles.hasNext()) {
    const f = rootFiles.next();
    const isDoc = f.getMimeType() === MimeType.GOOGLE_DOCS;
    const hasTranscript = /transcript/i.test(f.getName());
    lines.push('  - "' + f.getName() + '"');
    lines.push('     mimeType=' + f.getMimeType() + ' | isGoogleDoc=' + isDoc + ' | nameHasTranscript=' + hasTranscript);
    rootCount++;
  }
  if (rootCount === 0) lines.push('  (none)');
  lines.push('');

  // Subfolders and their files
  lines.push('== Subfolders and their files ==');
  const subs = meetFolder.getFolders();
  let subCount = 0;
  while (subs.hasNext()) {
    const sub = subs.next();
    subCount++;
    lines.push('  [Subfolder] "' + sub.getName() + '"');
    const files = sub.getFiles();
    let fCount = 0;
    while (files.hasNext()) {
      const f = files.next();
      const isDoc = f.getMimeType() === MimeType.GOOGLE_DOCS;
      const hasTranscript = /transcript/i.test(f.getName());
      lines.push('     - "' + f.getName() + '"');
      lines.push('        mimeType=' + f.getMimeType() + ' | isGoogleDoc=' + isDoc + ' | nameHasTranscript=' + hasTranscript);
      fCount++;
    }
    if (fCount === 0) lines.push('     (no files)');
  }
  if (subCount === 0) lines.push('  (no subfolders)');
  lines.push('');

  // Processed history
  const processed = props.getProperty(P.PROCESSED_IDS) || '';
  const processedCount = processed ? processed.split(',').length : 0;
  lines.push('Processed file IDs on record: ' + processedCount);
  if (processedCount > 0) {
    lines.push('(These files will be skipped even if they match. Use "Debug: Clear Processed History" to reset.)');
  }
  lines.push('');
  lines.push('== TO BE PROCESSED ==');
  lines.push('A file is processed only if ALL of: isGoogleDoc=true AND nameHasTranscript=true AND not in processed history.');

  const out = lines.join('\n');
  console.log(out);
  const html = HtmlService.createHtmlOutput(
    '<pre style="font-family:monospace;font-size:12px;white-space:pre-wrap;padding:12px">' +
    out.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') +
    '</pre>'
  ).setWidth(760).setHeight(560);
  ui.showModalDialog(html, 'Debug: What the script sees');
}

function debugClearProcessed() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'Clear processed history?',
    'This resets the list of transcripts the script has already handled. Next run will re-process everything it finds in your Meet folder.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;
  PropertiesService.getScriptProperties().deleteProperty(P.PROCESSED_IDS);
  ui.alert('Cleared. Now run "Process New Meetings Now".');
}
