/**
 * ACE Fitness – Upholstery Check-In / Check-Out Tracker
 * -----------------------------------------------------
 * Paste this whole file into Extensions > Apps Script of a NEW, empty Google Sheet,
 * save, then run setup() once. See docs/MAINTAINING.md in the GitHub repo for the full guide.
 *
 * What setup() builds:
 *   - A Google Form (no sign-in for techs) with branching:
 *       Status -> (new job: type client name) or (existing job: pick from list) -> details
 *   - Tabs in this spreadsheet:
 *       Current Status  one row per job, latest status, days in status, stuck-job alerts
 *       Log             one row per submission, full history (protected, never edited)
 *       Settings        tech names, alert recipients, "stuck" thresholds, form link
 *       Raw form responses (hidden backup Google keeps automatically)
 *   - A trigger that logs each submission and emails the Director on check-in and "Ready for reinstallation"
 *   - A trigger that refreshes the form's dropdowns when you edit the Settings tab
 */

// ===== Defaults used only the first time setup() runs. After that, edit the Settings tab. =====
const CONFIG = {
  TITLE: 'ACE Fitness – Upholstery Check-In',
  TIME_ZONE: 'America/Phoenix',
  ALERT_EMAILS: ['director@example.com'],  // placeholder: the real list lives in Settings, column C
  TECHS: ['Jesse', 'Tech name 2', 'Tech name 3', 'Tech name 4'],
  STUCK_DAYS_CHECKED_IN: 21,  // "Checked in at shop" and not back from Tucson yet
  STUCK_DAYS_READY: 3,        // "Ready for reinstallation" but not scheduled
};

// Tech check-in page (hosted on GitHub Pages): goes in Settings F6 and the QR code.
// It loads lists from, and posts entries to, this script's web app (doGet ?action=data / doPost).
const CHECKIN_URL = 'https://acefitnessequipment.github.io/ace-upholstery/';
// Director's dashboard (PIN in Settings F10): outstanding jobs + "Mark scheduled".
const DASHBOARD_URL = CHECKIN_URL + 'admin.html';
const DASHBOARD_TECH = 'Director (dashboard)';

// Three steps: the tech checks it in from a client, the tech marks it ready when it is back
// from Tucson, and the Director marks it scheduled (the final step, which closes the job).
const STATUS = {
  IN: 'Checked in at shop',
  BACK: 'Ready for reinstallation',
  SCHED: 'Scheduled',
};
const STATUS_LIST = [STATUS.IN, STATUS.BACK, STATUS.SCHED];
const FINAL_STATUS = STATUS.SCHED;
// Set by the dashboard's Delete button: hides the job everywhere. The Log keeps the record.
const DELETED = 'Deleted';
function isClosed_(status) { return status === FINAL_STATUS || status === DELETED; }
const STATUS_HELP = 'Checked in at shop = you just brought it back from a client.\n' +
  'Ready for reinstallation = it is back from Tucson and in the shop.\n' +
  'Scheduled = the reinstall is on the calendar (Director).';

const SHEETS = { STATUS: 'Current Status', LOG: 'Log', SETTINGS: 'Settings', RAW: 'Raw form responses' };
const LOG_HEADERS = ['Timestamp', 'Job', 'Client', 'Status', 'Pieces', 'Description', 'Technician', 'Notes', 'Response ID'];
const STATUS_HEADERS = ['Job', 'Client', 'Status', 'In status since', 'Days in status', 'Last tech',
  'Pieces', 'Description', 'First logged', 'Latest notes', 'Alert'];

const NOT_LISTED = 'Not in list (type the name below)';
const LABEL_SEP = ' · ';
const TEST_PREFIX = 'ZZ TEST';
const TEST_CLIENT = TEST_PREFIX + ' – Sample Gym';

// Settings tab layout (row/col are 1-based)
const SET = {
  TECH_COL: 1,          // A
  EMAIL_COL: 3,         // C
  DAYS_IN: 'F2',
  DAYS_SENT: 'F3',
  DAYS_BACK: 'F4',
  FORM_URL: 'F6',
  FORM_EDIT_URL: 'F7',
  ADMIN_PIN: 'F10',
  DASHBOARD_URL: 'F11',
};

// ======================================================================================
// Menu
// ======================================================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Upholstery Tracker')
    .addItem('Set up (run once)', 'setup')
    .addItem('Show check-in page link', 'showFormLink')
    .addItem('Write links to Settings', 'writeLinks')
    .addItem('Apply 3-step setup', 'applyThreeSteps')
    .addItem('Set up dashboard (PIN + link)', 'setupDashboard')
    .addItem('Refresh form dropdowns', 'refreshFormLists')
    .addItem('Rebuild Current Status', 'rebuildStatusNow')
    .addSeparator()
    .addItem('Run test job (sends 2 emails)', 'runTestJob')
    .addItem('Remove test job', 'removeTestJob')
    .addToUi();
}

// ======================================================================================
// Setup
// ======================================================================================
function setup() {
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Run this from Extensions > Apps Script inside the Google Sheet (a container-bound script).');
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('formId')) {
    notify('Already set up.\n\nForm link (for the QR code):\n' + props.getProperty('formUrl'));
    return;
  }

  ss.rename('ACE Upholstery Tracker');
  ss.setSpreadsheetTimeZone(CONFIG.TIME_ZONE);
  try { ss.setRecalculationInterval(SpreadsheetApp.RecalculationInterval.HOUR); } catch (e) { /* optional */ }

  // ---- Tabs ----
  const statusSh = getOrCreateSheet_(ss, SHEETS.STATUS, 0);
  const logSh = getOrCreateSheet_(ss, SHEETS.LOG, 1);
  const setSh = getOrCreateSheet_(ss, SHEETS.SETTINGS, 2);
  ss.getSheets().forEach(sh => {
    if ([SHEETS.STATUS, SHEETS.LOG, SHEETS.SETTINGS].indexOf(sh.getName()) === -1 && sh.getLastRow() === 0) {
      ss.deleteSheet(sh); // remove the empty default "Sheet1"
    }
  });

  setupLogSheet_(logSh);
  setupStatusSheet_(statusSh);
  setupSettingsSheet_(setSh);

  // ---- Form ----
  const built = buildForm_();
  const form = built.form;
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();
  ss.getSheets().forEach(sh => {
    if (sh.getFormUrl && sh.getFormUrl()) {
      sh.setName(SHEETS.RAW);
      sh.hideSheet();
    }
  });

  let formUrl = form.getPublishedUrl();
  try { formUrl = form.shortenFormUrl(formUrl); } catch (e) { /* long URL still works */ }

  props.setProperties({
    ssId: ss.getId(),
    formId: form.getId(),
    formUrl: formUrl,
    statusId: String(built.ids.status),
    clientTextId: String(built.ids.clientText),
    jobListId: String(built.ids.jobList),
    otherNameId: String(built.ids.otherName),
    piecesId: String(built.ids.pieces),
    descId: String(built.ids.desc),
    techId: String(built.ids.tech),
    notesId: String(built.ids.notes),
  });

  setSh.getRange(SET.FORM_URL).setValue(formUrl);
  setSh.getRange(SET.FORM_EDIT_URL).setValue(form.getEditUrl());

  // ---- Triggers ----
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('handleFormSubmit').forForm(form).onFormSubmit().create();
  ScriptApp.newTrigger('handleSettingsEdit').forSpreadsheet(ss).onEdit().create();

  refreshFormLists();
  ss.setActiveSheet(statusSh);

  notify('Setup complete.\n\nForm link (put this in the QR code):\n' + formUrl +
    '\n\nNext: open the form once to check it says "Published / Anyone with the link", ' +
    'edit the tech names in the Settings tab, then run "Run test job" from the Upholstery Tracker menu.');
}

function buildForm_() {
  const form = FormApp.create(CONFIG.TITLE);
  form.setDescription('One submission per client, every time upholstery changes hands. ' +
    'Client name must match the orange tape. Takes under a minute.');
  try { form.setCollectEmail(false); } catch (e) { }
  try { form.setRequireLogin(false); } catch (e) { }          // Workspace: don't restrict to org
  try { form.setLimitOneResponsePerUser(false); } catch (e) { }
  form.setAllowResponseEdits(false);
  form.setShowLinkToRespondAgain(true);
  form.setProgressBar(false);
  form.setConfirmationMessage('Logged. Thanks! Make sure every piece still has orange tape with the client name.');
  try { if (form.setPublished) form.setPublished(true); } catch (e) { }

  // Section 1: status
  const status = form.addMultipleChoiceItem()
    .setTitle('What are you logging?')
    .setHelpText(STATUS_HELP)
    .setRequired(true);

  // Section 2: new job
  const newPage = form.addPageBreakItem().setTitle('New job');
  const clientText = form.addTextItem()
    .setTitle('Client name (exactly as on the orange tape)')
    .setRequired(true);

  // Section 3: existing job
  const existingPage = form.addPageBreakItem().setTitle('Which job?');
  const jobList = form.addListItem()
    .setTitle('Client / job')
    .setHelpText('Jobs are listed as "Client · date checked in".')
    .setRequired(true);
  jobList.setChoiceValues([NOT_LISTED]);
  const otherName = form.addTextItem()
    .setTitle('Client name – only if you picked "Not in list"')
    .setRequired(false);

  // Section 4: details
  const detailsPage = form.addPageBreakItem().setTitle('Details');
  const pieces = form.addTextItem()
    .setTitle('Number of pieces')
    .setHelpText('Count them. e.g. 4')
    .setRequired(true)
    .setValidation(FormApp.createTextValidation().setHelpText('Enter a whole number, e.g. 4').requireWholeNumber().build());
  const desc = form.addTextItem()
    .setTitle('Piece description')
    .setHelpText('e.g. bench pad, seat, back pad')
    .setRequired(false);
  const tech = form.addListItem().setTitle('Your name').setRequired(true);
  tech.setChoiceValues(CONFIG.TECHS);
  const notes = form.addParagraphTextItem()
    .setTitle('Notes (optional)')
    .setHelpText('Damage, color, missing pieces, special instructions')
    .setRequired(false);

  // Branching
  setStatusChoices_(status, newPage, existingPage);
  // After the "New job" section, skip over "Which job?" straight to Details.
  // (setGoToPage on a page break controls where the section BEFORE it goes.)
  existingPage.setGoToPage(detailsPage);

  return {
    form: form,
    ids: {
      status: status.getId(), clientText: clientText.getId(), jobList: jobList.getId(),
      otherName: otherName.getId(), pieces: pieces.getId(), desc: desc.getId(),
      tech: tech.getId(), notes: notes.getId(),
    },
  };
}

/** New job -> "New job" section; every other status -> pick an existing job. */
function setStatusChoices_(item, newPage, existingPage) {
  item.setChoices(STATUS_LIST.map(s => item.createChoice(s, s === STATUS.IN ? newPage : existingPage)));
}

function setupLogSheet_(sh) {
  sh.clear();
  sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]).setFontWeight('bold').setBackground('#f3f3f3');
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('M/d/yyyy h:mm am/pm');
  sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 220); sh.setColumnWidth(3, 180); sh.setColumnWidth(4, 140);
  sh.setColumnWidth(6, 220); sh.setColumnWidth(8, 260);
  sh.hideColumns(LOG_HEADERS.length); // Response ID (used to avoid double-logging)
  const p = sh.protect().setDescription('Log is history – do not edit');
  p.setWarningOnly(true);
}

function setupStatusSheet_(sh) {
  sh.clear();
  sh.getRange(1, 1, 1, STATUS_HEADERS.length).setValues([STATUS_HEADERS]).setFontWeight('bold').setBackground('#f3f3f3');
  sh.setFrozenRows(1);
  sh.getRange('D:D').setNumberFormat('M/d/yyyy h:mm am/pm');
  sh.getRange('I:I').setNumberFormat('M/d/yyyy');
  sh.getRange('E:E').setNumberFormat('0').setHorizontalAlignment('center');
  sh.getRange('G:G').setHorizontalAlignment('center');
  [220, 170, 140, 150, 100, 120, 60, 220, 100, 240, 260].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  applyStatusFormatting_(sh);
}

/** Red = stuck, yellow = ready for reinstallation, grey = scheduled (closed). */
function applyStatusFormatting_(sh) {
  const range = sh.getRange('A2:K');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$K2<>""')
      .setBackground('#f8d7da').setFontColor('#842029').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$C2="' + STATUS.BACK + '"')
      .setBackground('#fff3cd').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$C2="' + FINAL_STATUS + '"')
      .setFontColor('#9e9e9e').setRanges([range]).build(),
  ]);
}

function setupSettingsSheet_(sh) {
  sh.clear();
  sh.getRange('A1').setValue('Technicians (form dropdown)');
  sh.getRange(2, SET.TECH_COL, CONFIG.TECHS.length, 1).setValues(CONFIG.TECHS.map(t => [t]));
  sh.getRange('C1').setValue('Director emails go to');
  sh.getRange(2, SET.EMAIL_COL, CONFIG.ALERT_EMAILS.length, 1).setValues(CONFIG.ALERT_EMAILS.map(e => [e]));
  sh.getRange('E1:F1').setValues([['Setting', 'Value']]);
  writeThresholds_(sh);
  sh.getRange('E6').setValue('Form link (this goes in the QR code)');
  sh.getRange('E7').setValue('Form edit link (Director only)');
  sh.getRange('A1:F1').setFontWeight('bold').setBackground('#f3f3f3');
  sh.getRange('E6:E7').setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 200); sh.setColumnWidth(3, 240); sh.setColumnWidth(5, 300); sh.setColumnWidth(6, 320);
  sh.getRange('A12').setValue('Edit columns A and C any time – the form dropdown updates automatically.')
    .setFontStyle('italic').setFontColor('#666666');
}

function writeThresholds_(sh) {
  sh.getRange('E2:F4').setValues([
    ['Flag "' + STATUS.IN + '" (not back from Tucson) after (days)', CONFIG.STUCK_DAYS_CHECKED_IN],
    ['Flag "' + STATUS.BACK + '" (not scheduled) after (days)', CONFIG.STUCK_DAYS_READY],
    ['', ''],
  ]);
}

// ======================================================================================
// Form submissions
// ======================================================================================
function handleFormSubmit(e) {
  processResponse_(e.response);
}

/** Logs one form response. Safe to call twice for the same response (deduped by response ID). */
function processResponse_(resp) {
  const p = getProps_();
  const a = {};
  resp.getItemResponses().forEach(ir => { a[String(ir.getItem().getId())] = ir.getResponse(); });
  return logEntry_(p, resp.getId(), resp.getTimestamp() || new Date(), a, true);
}

/**
 * Logs one entry. `a` holds answers keyed by the form's item IDs (p.statusId, p.clientTextId, ...).
 * Safe to call twice with the same rid (deduped). refreshForm also updates the backup Google Form's dropdowns.
 */
function logEntry_(p, rid, ts, a, refreshForm) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let emailError = null;
  let entry;
  try {
    const ss = SpreadsheetApp.openById(p.ssId);
    const logSh = ss.getSheetByName(SHEETS.LOG);
    const existingJob = loggedJobForResponse_(logSh, rid);
    if (existingJob) return { job: existingJob, duplicate: true };

    const status = String(a[p.statusId] || '');
    const jobs = loadJobs_(ss);
    let job, client;

    if (status === STATUS.IN) {
      client = cleanName_(a[p.clientTextId]) || '(name not given)';
      job = newJobLabel_(client, ts, jobs);
    } else {
      const picked = String(a[p.jobListId] || '');
      if (picked && picked !== NOT_LISTED) {
        job = picked;
        client = jobs[job] ? jobs[job].client : picked.split(LABEL_SEP)[0];
      } else {
        client = cleanName_(a[p.otherNameId]) || '(name not given)';
        job = findOpenJobByClient_(client, jobs) || newJobLabel_(client, ts, jobs);
        if (jobs[job]) client = jobs[job].client; // use the name as first logged
      }
    }

    entry = {
      ts: ts, job: job, client: client, status: status,
      pieces: toNumber_(a[p.piecesId]),
      desc: String(a[p.descId] || '').trim(),
      tech: String(a[p.techId] || ''),
      notes: String(a[p.notesId] || '').trim(),
      prev: jobs[job] || null,
    };

    logSh.appendRow([entry.ts, entry.job, entry.client, entry.status, entry.pieces,
      entry.desc, entry.tech, entry.notes, rid]);
    SpreadsheetApp.flush();

    if (status === STATUS.BACK || status === STATUS.IN) {
      try {
        if (status === STATUS.BACK) sendBackInShopEmail_(ss, entry);
        else sendCheckedInEmail_(ss, entry);
      } catch (err) { emailError = err; console.error(err); }
    }

    rebuildStatus_(ss);
    if (refreshForm) refreshFormLists_(ss, p);
  } finally {
    lock.releaseLock();
  }
  // Re-throw so Google emails the script owner a failure notice if the alert didn't go out.
  if (emailError) throw new Error('Logged, but the email to the Director failed: ' + emailError.message);
  return entry;
}

function sendBackInShopEmail_(ss, e) {
  const when = Utilities.formatDate(e.ts, CONFIG.TIME_ZONE, "EEE, MMM d 'at' h:mm a");
  const pieceWord = e.pieces === 1 ? 'piece' : 'pieces';
  const subject = 'Ready for reinstallation: ' + e.client + ' – ' + e.pieces + ' ' + pieceWord + ' back from Tucson';

  const list = splitPieces_(e.desc);
  const rows = [
    ['Client', e.client],
    ['Job', e.job],
    ['Pieces', String(e.pieces)],
  ];
  list.forEach((name, i) => rows.push(['Item ' + (i + 1), name]));
  rows.push(['Checked back in by', e.tech], ['When', when]);
  const prev = e.prev;
  if (prev && prev.firstTs) {
    const days = Math.floor((e.ts - prev.firstTs) / 86400000);
    rows.push(['Out for upholstery', days + ' day' + (days === 1 ? '' : 's') + ' (checked in ' +
      Utilities.formatDate(prev.firstTs, CONFIG.TIME_ZONE, 'MMM d') + ')']);
  }
  const back = list.map(x => x.toLowerCase());
  const missing = prev ? prev.pieceList.filter(x => back.indexOf(x.toLowerCase()) === -1) : [];
  if (missing.length) {
    rows.push(['⚠ Not back yet', missing.join(', ') + ' – please verify']);
  } else if (prev && prev.firstPieces !== '' && prev.firstPieces !== undefined && Number(prev.firstPieces) !== Number(e.pieces)) {
    rows.push(['⚠ Piece count', 'Checked in with ' + prev.firstPieces + ', back with ' + e.pieces + ' – please verify']);
  }
  if (e.notes) rows.push(['Notes', e.notes]);

  sendTrackerEmail_(ss, subject, 'Upholstery is back in the shop from Tucson and ready to schedule for reinstall.', rows,
    'When it is scheduled, log "' + STATUS.SCHED + '" on the check-in page.');
}

/** Sent when a tech brings upholstery back to the shop from a client. */
function sendCheckedInEmail_(ss, e) {
  const when = Utilities.formatDate(e.ts, CONFIG.TIME_ZONE, "EEE, MMM d 'at' h:mm a");
  const pieceWord = e.pieces === 1 ? 'piece' : 'pieces';
  const subject = 'Checked in at shop: ' + e.client + ' – ' + e.pieces + ' ' + pieceWord + ' removed from client';
  const rows = [['Client', e.client], ['Job', e.job], ['Pieces', String(e.pieces)]];
  splitPieces_(e.desc).forEach((name, i) => rows.push(['Item ' + (i + 1), name]));
  rows.push(['Checked in by', e.tech], ['When', when]);
  if (e.notes) rows.push(['Notes', e.notes]);
  sendTrackerEmail_(ss, subject, 'Upholstery was removed at a client and is now in the shop, waiting to go to Tucson.', rows,
    'Next: when it is back from Tucson, the tech logs "' + STATUS.BACK + '".');
}

function sendTrackerEmail_(ss, subject, intro, rows, footer) {
  const to = readList_(ss.getSheetByName(SHEETS.SETTINGS), SET.EMAIL_COL).filter(x => /.+@.+\..+/.test(x));
  if (!to.length) throw new Error('No email addresses in Settings column C.');
  const html =
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#222">' +
    '<p><b>' + esc_(intro) + '</b></p>' +
    '<table cellpadding="6" style="border-collapse:collapse">' +
    rows.map(r => '<tr><td style="color:#666;border-bottom:1px solid #eee">' + esc_(r[0]) +
      '</td><td style="border-bottom:1px solid #eee"><b>' + esc_(r[1]) + '</b></td></tr>').join('') +
    '</table>' +
    '<p><a href="' + ss.getUrl() + '">Open the tracker</a> &nbsp;·&nbsp; ' + esc_(footer) + '</p></div>';
  const text = intro + '\n\n' + rows.map(r => r[0] + ': ' + r[1]).join('\n') + '\n\n' + footer + '\nTracker: ' + ss.getUrl();
  MailApp.sendEmail({ to: to.join(','), subject: subject, htmlBody: html, body: text, name: 'ACE Upholstery Tracker' });
}

// ======================================================================================
// Current Status + form dropdowns
// ======================================================================================
function rebuildStatusNow() {
  rebuildStatus_(SpreadsheetApp.openById(getProps_().ssId));
}

function rebuildStatus_(ss) {
  const sh = ss.getSheetByName(SHEETS.STATUS);
  const all = loadJobs_(ss);
  const jobs = Object.keys(all).map(k => all[k]).filter(j => j.status !== DELETED);
  const order = {};
  [STATUS.BACK, STATUS.IN, FINAL_STATUS].forEach((s, i) => { order[s] = i; });
  const rank = s => (s in order ? order[s] : 9);
  jobs.sort((a, b) => rank(a.status) - rank(b.status) ||
    (a.status === FINAL_STATUS ? b.since - a.since : a.since - b.since));

  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, STATUS_HEADERS.length).clearContent();
  if (!jobs.length) return;

  const need = jobs.length + 1;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());

  const rows = jobs.map((j, i) => {
    const r = i + 2;
    return [j.job, j.client, j.status, j.since, '=IF(D' + r + '="","",TODAY()-INT(D' + r + '))', j.tech,
      j.pieces === undefined ? '' : j.pieces, j.desc || '', j.firstTs, j.notes || '', alertFormula_(r)];
  });
  sh.getRange(2, 1, rows.length, STATUS_HEADERS.length).setValues(rows);
}

function alertFormula_(r) {
  const C = 'C' + r, E = 'E' + r, S = SHEETS.SETTINGS + '!$F$';
  return '=IF(AND(' + C + '="' + STATUS.IN + '",' + E + '>' + S + '2),"Out "&' + E + '&" days – not back from Tucson",' +
    'IF(AND(' + C + '="' + STATUS.BACK + '",' + E + '>' + S + '3),"Ready "&' + E + '&" days – not scheduled",""))';
}

function refreshFormLists() {
  const p = getProps_();
  refreshFormLists_(SpreadsheetApp.openById(p.ssId), p);
}

function openJobLabels_(jobs) {
  return Object.keys(jobs).filter(k => !isClosed_(jobs[k].status))
    .sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));
}

function refreshFormLists_(ss, p) {
  const form = FormApp.openById(p.formId);
  const open = openJobLabels_(loadJobs_(ss));
  form.getItemById(Number(p.jobListId)).asListItem().setChoiceValues(open.concat([NOT_LISTED]));

  const techs = readList_(ss.getSheetByName(SHEETS.SETTINGS), SET.TECH_COL);
  form.getItemById(Number(p.techId)).asListItem()
    .setChoiceValues(techs.length ? techs : ['(add tech names in the Settings tab)']);
}

/** Installable onEdit: keep form dropdowns in sync with the Settings tab. */
function handleSettingsEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() === SHEETS.SETTINGS && e.range.getColumn() <= SET.EMAIL_COL) refreshFormLists();
}

// ======================================================================================
// Tech check-in page (web app). Same questions as the Google Form, but no Google
// account bar. Each entry is submitted to the Google Form behind the scenes, so the
// raw-responses backup and everything else keep working.
// Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone.
// ======================================================================================
function doGet(e) {
  // JSON feed for the hosted check-in page: ?action=data
  if (e && e.parameter && e.parameter.action === 'data') return json_({ ok: true, data: pageData_() });
  const t = HtmlService.createTemplateFromFile('CheckIn');
  t.data = JSON.stringify(pageData_()).replace(/</g, '\\u003c');
  return t.evaluate()
    .setTitle('Upholstery Check-In')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function pageData_() {
  const p = getProps_();
  const ss = SpreadsheetApp.openById(p.ssId);
  const jobs = loadJobs_(ss);
  return {
    statuses: STATUS_LIST,
    newStatus: STATUS.IN,
    backStatus: STATUS.BACK,
    finalStatus: FINAL_STATUS,
    notListed: NOT_LISTED,
    openJobs: openJobLabels_(jobs).map(k => ({ label: k, pieces: jobs[k].pieceList })),
    techs: readList_(ss.getSheetByName(SHEETS.SETTINGS), SET.TECH_COL),
  };
}

/** The hosted check-in page posts entries here as JSON (sent as text/plain to avoid a CORS preflight). */
function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.action === 'jobs') return json_(Object.assign({ ok: true }, adminJobs_(d.key)));
    if (d.action === 'schedule') return json_(Object.assign({ ok: true }, adminSchedule_(d)));
    if (d.action === 'delete') return json_(Object.assign({ ok: true }, adminDelete_(d)));
    return json_(Object.assign({ ok: true }, submitEntry(d)));
  } catch (err) {
    return json_({ ok: false, error: (err && err.message) || String(err) });
  }
}

// ---------- Director's dashboard (admin.html) ----------
/** Checks the dashboard PIN (Settings F10). Locks out for 10 minutes after 20 wrong tries. */
function checkAdmin_(key) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('adminFails') || 0);
  if (fails >= 20) throw new Error('Too many wrong PINs. Try again in 10 minutes.');
  const ss = SpreadsheetApp.openById(getProps_().ssId);
  const pin = String(ss.getSheetByName(SHEETS.SETTINGS).getRange(SET.ADMIN_PIN).getValue()).trim();
  if (!pin) throw new Error('The dashboard is not set up yet. In the sheet: Upholstery Tracker > Set up dashboard.');
  if (String(key || '').trim() !== pin) {
    cache.put('adminFails', String(fails + 1), 600);
    throw new Error('Wrong PIN.');
  }
  return ss;
}

/** Every job that is not scheduled yet, ready-for-reinstallation first, oldest first. */
function adminJobs_(key) {
  const ss = checkAdmin_(key);
  const setSh = ss.getSheetByName(SHEETS.SETTINGS);
  const limitIn = Number(setSh.getRange('F2').getValue()) || CONFIG.STUCK_DAYS_CHECKED_IN;
  const limitReady = Number(setSh.getRange('F3').getValue()) || CONFIG.STUCK_DAYS_READY;
  const jobs = loadJobs_(ss);
  const now = new Date();
  const list = Object.keys(jobs).map(k => jobs[k]).filter(j => !isClosed_(j.status)).map(j => {
    const days = Math.floor((now - j.since) / 86400000);
    return {
      job: j.job, client: j.client, status: j.status, days: days,
      since: j.since.toISOString(), checkedIn: j.firstTs.toISOString(),
      pieces: j.pieceList, pieceCount: j.pieces === undefined ? j.pieceList.length : j.pieces,
      tech: String(j.tech || ''), notes: String(j.notes || ''),
      stuck: days > (j.status === STATUS.BACK ? limitReady : limitIn),
    };
  });
  const rank = s => (s === STATUS.BACK ? 0 : 1);
  list.sort((a, b) => rank(a.status) - rank(b.status) || b.days - a.days);
  return {
    jobs: list,
    limits: { checkedIn: limitIn, ready: limitReady },
    statuses: { checkedIn: STATUS.IN, ready: STATUS.BACK },
    sheetUrl: ss.getUrl(),
  };
}

/** Marks one job Scheduled (closes it). Optional reinstall date goes in the notes. */
function adminSchedule_(d) {
  const ss = checkAdmin_(d.key);
  const j = loadJobs_(ss)[String(d.job || '')];
  if (!j) throw new Error('That job was not found. Refresh the list.');
  if (!isClosed_(j.status)) {
    const date = String(d.date || '').trim().slice(0, 40);
    submitEntry({
      status: FINAL_STATUS, job: j.job, tech: DASHBOARD_TECH, rid: d.rid,
      pieces: j.pieceList.length ? j.pieceList : [j.pieces ? j.pieces + ' pieces' : 'Pieces not listed'],
      notes: date ? 'Reinstall scheduled for ' + date : '',
    });
  }
  return adminJobs_(d.key);
}

/** Deletes one job: hides it from the dashboard, the check-in list and Current Status. The Log keeps a "Deleted" line. */
function adminDelete_(d) {
  const ss = checkAdmin_(d.key);
  const p = getProps_();
  const j = loadJobs_(ss)[String(d.job || '')];
  if (!j) throw new Error('That job was not found. Refresh the list.');
  if (j.status !== DELETED) {
    const a = {};
    a[p.statusId] = DELETED;
    a[p.jobListId] = j.job;
    a[p.piecesId] = String(j.pieces === undefined ? j.pieceList.length : j.pieces);
    a[p.descId] = j.pieceList.join('; ');
    a[p.techId] = DASHBOARD_TECH;
    a[p.notesId] = 'Deleted from the dashboard (was "' + j.status + '")';
    const rid = String(d.rid || '').replace(/[^\w-]/g, '').slice(0, 60) || ('del-' + Utilities.getUuid());
    logEntry_(p, rid, new Date(), a, false);
  }
  return adminJobs_(d.key);
}

/** Creates the dashboard PIN (if there isn't one) and writes the PIN + link into Settings. */
function setupDashboard() {
  const p = getProps_();
  const sh = SpreadsheetApp.openById(p.ssId).getSheetByName(SHEETS.SETTINGS);
  let pin = String(sh.getRange(SET.ADMIN_PIN).getValue()).trim();
  if (!pin) pin = String(100000 + Math.floor(Math.random() * 900000));
  sh.getRange(SET.ADMIN_PIN).setNumberFormat('@');
  sh.getRange('E10:F11').setValues([
    ['Dashboard PIN (change it any time)', pin],
    ['Dashboard link (Director only)', DASHBOARD_URL],
  ]);
  sh.getRange('E10:E11').setFontWeight('bold');
  notify('Dashboard ready: ' + DASHBOARD_URL + '\nThe PIN is in Settings F10.');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Called from the check-in page. Returns fresh lists for the next entry. */
function submitEntry(d) {
  d = d || {};
  const p = getProps_();
  const status = String(d.status || '');
  const pieceList = (Array.isArray(d.pieces) ? d.pieces : [])
    .map(x => String(x || '').replace(/;/g, ',').replace(/\s+/g, ' ').trim()).filter(Boolean);
  const tech = String(d.tech || '').trim();
  const notes = String(d.notes || '').trim();

  if (STATUS_LIST.indexOf(status) === -1) throw new Error('Pick what you are logging.');
  if (!pieceList.length) throw new Error('Add at least one piece.');
  if (!tech) throw new Error('Pick your name.');

  const a = {};
  a[p.statusId] = status;
  if (status === STATUS.IN) {
    const name = cleanName_(d.client);
    if (!name) throw new Error('Enter the client name from the orange tape.');
    a[p.clientTextId] = name;
  } else {
    const job = String(d.job || '');
    if (!job) throw new Error('Pick the client / job.');
    a[p.jobListId] = job;
    if (job === NOT_LISTED) {
      const name = cleanName_(d.other);
      if (!name) throw new Error('Type the client name from the orange tape.');
      a[p.otherNameId] = name;
    }
  }
  a[p.piecesId] = String(pieceList.length);
  a[p.descId] = pieceList.join('; ');
  a[p.techId] = tech;
  a[p.notesId] = notes;

  // Logged straight to the sheet (no Google Form round trip), so saving takes a couple of seconds.
  // The page sends the same rid if a tech taps Submit again, so a retry never logs twice.
  const rid = String(d.rid || '').replace(/[^\w-]/g, '').slice(0, 60) || ('web-' + Utilities.getUuid());
  const entry = logEntry_(p, rid, new Date(), a, false);
  return { job: entry.job, client: entry.client, status: status, data: pageData_() };
}

// ======================================================================================
// Test run
// ======================================================================================
/** Pushes one fake job through every status. Sends two emails (check-in and ready). */
function runTestJob() {
  const p = getProps_();
  const ss = SpreadsheetApp.openById(p.ssId);
  const form = FormApp.openById(p.formId);
  const item = id => form.getItemById(Number(id));
  const techs = readList_(ss.getSheetByName(SHEETS.SETTINGS), SET.TECH_COL);
  const tech = techs[0] || 'Test';

  let job = null;
  STATUS_LIST.forEach(status => {
    const r = form.createResponse();
    r.withItemResponse(item(p.statusId).asMultipleChoiceItem().createResponse(status));
    if (status === STATUS.IN) {
      r.withItemResponse(item(p.clientTextId).asTextItem().createResponse(TEST_CLIENT));
    } else {
      r.withItemResponse(item(p.jobListId).asListItem().createResponse(job));
    }
    r.withItemResponse(item(p.piecesId).asTextItem().createResponse('4'));
    r.withItemResponse(item(p.descId).asTextItem().createResponse('Bench pad; Seat pad; Back pad; Roller pad'));
    r.withItemResponse(item(p.techId).asListItem().createResponse(tech));
    r.withItemResponse(item(p.notesId).asParagraphTextItem().createResponse('Test entry – safe to delete'));
    const submitted = r.submit();
    const result = processResponse_(submitted); // the trigger may also fire; duplicates are skipped
    job = result.job;
  });

  notify('Test job "' + job + '" went through all 5 statuses.\n\n' +
    'Check the Log and Current Status tabs, and check ' + readList_(ss.getSheetByName(SHEETS.SETTINGS), SET.EMAIL_COL).join(', ') +
    ' for the two emails (check Junk the first time).\n\nWhen done, run "Remove test job".');
}

function removeTestJob() {
  const p = getProps_();
  const ss = SpreadsheetApp.openById(p.ssId);
  const isTest = v => String(v).indexOf(TEST_PREFIX) === 0;

  deleteRowsWhere_(ss.getSheetByName(SHEETS.LOG), row => isTest(row[2]));
  const raw = ss.getSheetByName(SHEETS.RAW);
  if (raw) deleteRowsWhere_(raw, row => row.some(isTest));

  const form = FormApp.openById(p.formId);
  form.getResponses().forEach(resp => {
    if (resp.getItemResponses().some(ir => isTest(ir.getResponse()))) form.deleteResponse(resp.getId());
  });

  rebuildStatus_(ss);
  refreshFormLists_(ss, p);
  notify('Test job removed.');
}

// ======================================================================================
// Helpers
// ======================================================================================
function showFormLink() {
  notify('Check-in page link (for the QR code):\n' + CHECKIN_URL);
}

/** Puts the current links into the Settings tab (E6:F8). */
function writeLinks() {
  const p = getProps_();
  const sh = SpreadsheetApp.openById(p.ssId).getSheetByName(SHEETS.SETTINGS);
  sh.getRange('E6:F8').setValues([
    ['Check-in page link (this goes in the QR code)', CHECKIN_URL],
    ['Form edit link (Director only)', FormApp.openById(p.formId).getEditUrl()],
    ['Old Google Form link (backup only, techs do not use it)', p.formUrl],
  ]);
  sh.getRange('E6:E8').setFontWeight('bold');
  notify('Links written to the Settings tab.');
}

/** Brings an existing tracker to the 3-step process: form choices, Settings thresholds, colors. Safe to re-run. */
function applyThreeSteps() {
  const p = getProps_();
  const ss = SpreadsheetApp.openById(p.ssId);
  const form = FormApp.openById(p.formId);
  const pages = form.getItems(FormApp.ItemType.PAGE_BREAK).map(i => i.asPageBreakItem());
  const page = title => pages.filter(x => x.getTitle() === title)[0];
  const status = form.getItemById(Number(p.statusId)).asMultipleChoiceItem();
  status.setHelpText(STATUS_HELP);
  setStatusChoices_(status, page('New job'), page('Which job?'));
  writeThresholds_(ss.getSheetByName(SHEETS.SETTINGS));
  ss.getSheetByName(SHEETS.SETTINGS).getRange('C1').setValue('Director emails go to');
  applyStatusFormatting_(ss.getSheetByName(SHEETS.STATUS));
  rebuildStatus_(ss);
  refreshFormLists_(ss, p);
  notify('3-step setup applied: ' + STATUS_LIST.join(' → ') + '.');
}

function getProps_() {
  const p = PropertiesService.getScriptProperties().getProperties();
  if (!p.formId) throw new Error('Not set up yet – run setup() first.');
  return p;
}

/** Replays the Log into one object per job, keyed by job label. */
function loadJobs_(ss) {
  const sh = ss.getSheetByName(SHEETS.LOG);
  const jobs = {};
  const n = sh.getLastRow() - 1;
  if (n >= 1) {
    const rows = sh.getRange(2, 1, n, LOG_HEADERS.length).getValues()
      .filter(r => r[1] !== '' && Object.prototype.toString.call(r[0]) === '[object Date]')
      .sort((a, b) => a[0] - b[0]);
    rows.forEach(r => {
      const ts = r[0], job = String(r[1]), client = String(r[2]), status = String(r[3]);
      let j = jobs[job];
      if (!j) {
        j = jobs[job] = { job: job, client: client, firstTs: ts, firstPieces: r[4], status: null, since: ts, pieceList: [] };
      }
      if (status !== j.status) { j.status = status; j.since = ts; }
      j.tech = r[6];
      j.lastTs = ts;
      if (r[4] !== '') j.pieces = r[4];
      if (r[5] !== '') {
        j.desc = r[5];
        if (!j.pieceList.length) j.pieceList = splitPieces_(r[5]); // the pieces as first logged
      }
      if (r[7] !== '') j.notes = r[7];
    });
  }
  return jobs;
}

function loggedJobForResponse_(logSh, rid) {
  const n = logSh.getLastRow() - 1;
  if (n < 1) return null;
  const vals = logSh.getRange(2, 1, n, LOG_HEADERS.length).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][LOG_HEADERS.length - 1] === rid) return String(vals[i][1]);
  }
  return null;
}

function newJobLabel_(client, ts, jobs) {
  const base = client + LABEL_SEP + Utilities.formatDate(ts, CONFIG.TIME_ZONE, 'M/d');
  if (!jobs[base]) return base;
  let i = 2;
  while (jobs[base + ' (' + i + ')']) i++;
  return base + ' (' + i + ')';
}

/** Most recent open job for this client name (case/spacing-insensitive), or null. */
function findOpenJobByClient_(client, jobs) {
  const key = client.toLowerCase();
  let best = null;
  Object.keys(jobs).forEach(k => {
    const j = jobs[k];
    if (!isClosed_(j.status) && j.client.toLowerCase() === key && (!best || j.firstTs > best.firstTs)) best = j;
  });
  return best ? best.job : null;
}

function readList_(sheet, col) {
  const n = Math.max(sheet.getLastRow() - 1, 0);
  if (!n) return [];
  return sheet.getRange(2, col, n, 1).getValues()
    .map(r => String(r[0]).trim())
    .filter(v => v && v.indexOf('Edit columns') !== 0);
}

function deleteRowsWhere_(sh, test) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const vals = sh.getRange(2, 1, n, sh.getLastColumn()).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (test(vals[i])) sh.deleteRow(i + 2);
  }
}

function getOrCreateSheet_(ss, name, index) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name, index);
  return sh;
}

/** "Head pad; Back pad" -> ["Head pad", "Back pad"] */
function splitPieces_(s) { return String(s || '').split(';').map(x => x.trim()).filter(Boolean); }
function cleanName_(v) { return String(v || '').replace(/\s+/g, ' ').trim(); }
function toNumber_(v) { const n = parseInt(String(v || '').trim(), 10); return isNaN(n) ? '' : n; }
function esc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** Non-blocking notice (a toast in the sheet). A modal alert would hold the script open until clicked. */
function notify(msg) {
  console.log(msg);
  try { SpreadsheetApp.getActive().toast(msg, 'Upholstery Tracker', 20); } catch (e) { /* run from the editor: see Execution log */ }
}
