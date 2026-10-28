/**
 * EXPERIMENTATION REPOSITORY AUTOMATIONS  (Audience Development tab)
 * -------------------------------------------------------------------
 * Runs once a week (Tuesday morning by default).
 *
 * 1. Lead emails (one per lead, only when there is something to say)
 *    - "Launching within 3 weeks": sent ONCE per test, in the first weekly email
 *      where the Live Date is 21 days away or less. A new reminder goes out only
 *      if the Live Date changes.
 *    - "Live Date has passed. Did these launch?": listed EVERY week while the Live
 *      Date has passed and Status is still Planning / Design / In Development.
 *    Status is never changed automatically. Overdue rows are also highlighted.
 *
 * 2. Weekly calendar summary (for the "Experimentation Calendar" slide)
 *    One email to the people in SUMMARY_RECIPIENTS, laid out like the slide:
 *    Recently Ended / Recently Launched / Upcoming / Delays, with fiscal weeks.
 *    Each cell can be copied straight into the slide table.
 *
 * 3. "Launched Last 7 Days" tab: a live formula, updates on its own.
 *
 * 4. Jira (optional, read-only): shows each test's Jira ticket status in the lead
 *    emails, using the free "Jira Cloud for Sheets" add-on.
 *
 * Lead emails are read from the name chips (@-mentions) in the Marketing Lead column.
 * Settings you may want to change are all in CONFIG below.
 */

const CONFIG = {
  // ---- Where the data lives ----
  TAB_NAME: 'Audience Development',
  HEADER_ROW: 2,       // row with column names
  FIRST_DATA_ROW: 3,   // first row of tests

  // Column names (matched by name, so moving columns won't break anything)
  COLUMNS: {
    id: 'Identifier',
    name: 'Experiment Name',
    lead: 'Marketing Lead',
    status: 'Status',
    liveDate: 'Live Date',
    // optional
    endDate: 'End Date',
    team: 'Marketing Team',
    site: 'Site of Origin',
    description: 'Description',
    result: 'Result',
    conversion: 'Conversion Impact',
    registration: 'Registration Impact',
    plan: 'Experiment Plan',
    tracker: 'Test Tracker',
    jira: 'Jira Ticket',
  },

  // ---- Schedule ----
  SEND_DAY: 'TUESDAY', // MONDAY, TUESDAY, WEDNESDAY, ...
  SEND_HOUR: 8,        // 8 = between 8 and 9am, in the spreadsheet's time zone
  PAUSED: false,       // set to true to stop all emails (anyone with edit access can do this)

  // ---- 3-week reminder ----
  REMINDER_DAYS_BEFORE: 21,
  REMINDER_STATUSES: ['New', 'Backlog', 'Planning', 'Design', 'In Development'],

  // ---- "Did this launch?" check (Live Date passed, Status still one of these) ----
  LAUNCH_CHECK_STATUSES: ['Planning', 'Design', 'In Development'],
  HIGHLIGHT_COLOR: '#FCE8B2',

  // Leads to leave out for now (exact match on the Marketing Lead cell)
  SKIP_LEADS: ['GAIN'],

  // ---- Email settings ----
  SENDER_NAME: 'Experimentation Repo Reminders',
  CC: '',              // optional, e.g. 'you@company.com'
  TEST_RECIPIENT: '',  // blank = whoever clicks "Send this week's emails to me"

  // ---- Weekly calendar summary (slide-ready) ----
  SUMMARY_RECIPIENTS: [],      // e.g. ['manager@company.com']
  SUMMARY_INCLUDE_OWNER: true, // also send it to whoever turned the weekly emails on
  FISCAL_YEAR_START: '2026-06-29', // Monday that starts FW1 (FY27). Update when FY28 begins.
  DEFAULT_TEAM: 'Audience',    // used when the Marketing Team cell is blank

  // ---- Jira (Jira Cloud for Sheets add-on) ----
  JIRA_ENABLED: true,
  JIRA_PROJECT_KEYS: ['CIRC'], // ticket keys to pick up from the Experiment Plan text
  JIRA_DONE_STATUSES: ['Done', 'Released', 'Closed', 'Resolved'],
  JIRA_REFRESH_HOUR: 6,        // refreshes the Jira tab two hours before the emails

  // ---- Tabs this script creates ----
  LAUNCH_TAB: 'Launched Last 7 Days',
  LOG_TAB: 'Automation Log',
  JIRA_TAB: 'Jira Data',
  LEADS_TAB: 'Leads',  // optional backup: Name in column A, Email in column B
};

const TYPE_REMINDER_ = 'Reminder';
const TYPE_CHECK_ = 'Launch check';
const TYPE_SUMMARY_ = 'Summary';
const EMAIL_RE_ = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const EMAIL_FIND_RE_ = /[^\s@<>(),;:"']+@[^\s@<>(),;:"']+\.[a-z]{2,}/gi;
const JIRA_KEY_RE_ = /\b[A-Z][A-Z0-9_]+-\d+\b/g;
const LOG_HEADERS_ = ['Timestamp', 'Type', 'Identifier', 'Experiment', 'Live Date', 'Status', 'Sent To', 'Result', 'Key'];
const JIRA_FORMULA_CELL_ = 'A4';
const SECTION_STYLE_ = {
  ended: { title: 'Recently Ended', color: '#0E2638' },
  launched: { title: 'Recently Launched', color: '#3197FE' },
  upcoming: { title: 'Upcoming', color: '#93C47D' },
  delays: { title: 'Delays/Major Changes', color: '#FF9900' },
};
let CHIP_READ_ERROR_ = '';

/* =========================  MENU  ========================= */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Repo Automations')
    .addItem('1. Check lead emails', 'checkLeadEmails')
    .addItem('2. Set up highlighting + launch tab', 'setUpSheet')
    .addItem('3. Set up Jira tab', 'setUpJira')
    .addItem('4. Check Jira data', 'checkJiraData')
    .addItem("5. Preview this week's emails", 'previewThisWeek')
    .addItem("6. Send this week's emails to me (test)", 'sendTestEmailsToMe')
    .addSeparator()
    .addItem('Turn ON weekly emails', 'turnOnWeeklyEmails')
    .addItem('Turn OFF weekly emails', 'turnOffWeeklyEmails')
    .addToUi();
}

/** Lists every Marketing Lead value and the email(s) the script found for it. */
function checkLeadEmails() {
  const ctx = loadContext_();
  const seen = {};
  const lines = [];
  ctx.rows.forEach(function (r) {
    if (!r.leadText || seen[r.leadText]) return;
    seen[r.leadText] = true;
    if (r.skipped) lines.push(r.leadText + '  →  skipped for now (on the skip list)');
    else if (r.emails.length) lines.push(r.leadText + '  →  ' + r.emails.join(', ') + '   [from ' + r.emailSources.join(', ') + ']');
    else lines.push(r.leadText + '  →  NO EMAIL FOUND');
  });
  let msg = lines.join('\n');
  if (lines.some(function (l) { return l.indexOf('NO EMAIL FOUND') !== -1; })) {
    msg += '\n\nFor any "NO EMAIL FOUND": re-enter that name as an @-mention chip, ' +
      'or add a "' + CONFIG.LEADS_TAB + '" tab with Name in column A and Email in column B.';
  }
  if (typeof Sheets === 'undefined') {
    msg += '\n\nNote: the Google Sheets API service is not added yet, so name chips could not be read. See setup step 3.';
  } else if (CHIP_READ_ERROR_) {
    msg += '\n\nNote: the name chips could not be read (' + CHIP_READ_ERROR_ + ').\n' +
      'Fix: in Apps Script, open the file with your existing macros and delete the first line, /** @OnlyCurrentDoc */. ' +
      'Save, run this check again, and approve the new permission prompt.';
  }
  showMessage_('Lead emails found', msg);
}

/** Adds the highlight rule, the "Launched Last 7 Days" tab and the log tab. Safe to run again. */
function setUpSheet() {
  const ctx = loadContext_();
  const hadLaunch = findSheet_(ctx.ss, CONFIG.LAUNCH_TAB);
  const hadLog = findSheet_(ctx.ss, CONFIG.LOG_TAB);
  addHighlightRule_(ctx);
  const launch = buildLaunchTab_(ctx);
  const log = ensureLogSheet_(ctx.ss);
  showMessage_('Setup complete',
    '• Highlight rule: set (rows whose Live Date has passed while Status is still ' + CONFIG.LAUNCH_CHECK_STATUSES.join(' / ') + ')\n' +
    '• "' + launch.getName() + '" tab: ' + (hadLaunch ? 'already existed, formula refreshed' : 'created') + '\n' +
    '• "' + log.getName() + '" tab: ' + (hadLog ? 'already existed, left as is' : 'created') + '\n\n' +
    'Running setup again is safe. It reuses these tabs instead of adding new ones.');
}

/** Shows what would be emailed this week, without sending anything. */
function previewThisWeek() {
  const ctx = loadContext_();
  const keys = readLogKeys_(ctx.ss);
  const plan = buildPlan_(ctx, keys);
  const summary = buildSummary_(ctx);
  showMessage_("Preview: this week's emails", describePlan_(ctx, plan) + '\n' + describeSummary_(ctx, summary, keys));
}

/** Sends this week's emails (lead emails + calendar summary) to you. Nothing is recorded as sent. */
function sendTestEmailsToMe() {
  const ctx = loadContext_();
  const plan = buildPlan_(ctx, readLogKeys_(ctx.ss));
  const n = sendDigests_(ctx, plan, 'test', {});
  const s = sendSummary_(ctx, buildSummary_(ctx), 'test', {});
  const me = CONFIG.TEST_RECIPIENT || currentUserEmail_();
  showMessage_('Test emails sent', (n + s) + ' test email(s) sent to ' + me + ' (' + n + ' lead email(s), ' + s + ' calendar summary). Nobody else was emailed.');
}

function turnOnWeeklyEmails() {
  removeScheduleTriggers_();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const day = ScriptApp.WeekDay[CONFIG.SEND_DAY];
  ScriptApp.newTrigger('weeklyRun').timeBased().onWeekDay(day).atHour(CONFIG.SEND_HOUR).inTimezone(tz).create();
  let jiraNote = '';
  if (CONFIG.JIRA_ENABLED && findSheet_(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.JIRA_TAB)) {
    ScriptApp.newTrigger('refreshJiraQuery').timeBased().onWeekDay(day).atHour(CONFIG.JIRA_REFRESH_HOUR).inTimezone(tz).create();
    jiraNote = '\nThe Jira tab refreshes at ' + CONFIG.JIRA_REFRESH_HOUR + ':00 the same morning.';
  }
  showMessage_('Weekly emails are ON',
    'Emails go out every ' + titleCase_(CONFIG.SEND_DAY) + ' between ' + CONFIG.SEND_HOUR + ':00 and ' + (CONFIG.SEND_HOUR + 1) + ':00 (' + tz + '), ' +
    'sent from ' + currentUserEmail_() + '.' + jiraNote + '\n\n' +
    'Leads only get an email in weeks when they have something on the list.\nOnly one person should turn this on.');
}

function turnOffWeeklyEmails() {
  const n = removeScheduleTriggers_();
  showMessage_('Weekly emails are OFF', n ? 'Your weekly schedule has been removed.' : 'You had no schedule turned on.');
}

/** Runs on the weekly schedule. */
function weeklyRun() {
  if (CONFIG.PAUSED) { Logger.log('Paused (CONFIG.PAUSED is true). No emails sent.'); return; }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { Logger.log('Another run is in progress. Skipping.'); return; }
  try {
    const ctx = loadContext_();
    const keys = readLogKeys_(ctx.ss);
    const n = sendDigests_(ctx, buildPlan_(ctx, keys), 'live', keys);
    const s = sendSummary_(ctx, buildSummary_(ctx), 'live', keys);
    Logger.log('Weekly run complete. Lead emails: ' + n + ', summaries: ' + s);
  } finally {
    lock.releaseLock();
  }
}

/** Left over from the earlier daily schedule. Does nothing; "Turn ON weekly emails" removes it. */
function dailyRun() {
  Logger.log('The daily schedule has been replaced. Click Repo Automations → Turn ON weekly emails.');
}

/* =========================  READING THE SHEET  ========================= */

function loadContext_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = findSheet_(ss, CONFIG.TAB_NAME);
  if (!sheet) throw new Error('Could not find a tab named "' + CONFIG.TAB_NAME + '". Update CONFIG.TAB_NAME.');
  const tz = ss.getSpreadsheetTimeZone();
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0];
  const cols = findColumns_(headers);
  const numRows = Math.max(0, lastRow - CONFIG.FIRST_DATA_ROW + 1);
  const today = today_(tz);
  const rows = [];

  if (numRows > 0) {
    const range = sheet.getRange(CONFIG.FIRST_DATA_ROW, 1, numRows, lastCol);
    const values = range.getValues();
    const display = range.getDisplayValues();
    const leadEmails = readLeadEmails_(ss, sheet, cols.lead, numRows, display);
    const trackerLinks = readCellLinks_(ss, sheet, cols.tracker, numRows);
    const planLinks = readCellLinks_(ss, sheet, cols.plan, numRows);
    const cell = function (i, key) { return cols[key] ? String(display[i][cols[key] - 1]).trim() : ''; };

    for (let i = 0; i < numRows; i++) {
      const id = cell(i, 'id');
      if (!id) continue;
      const leadText = cell(i, 'lead');
      rows.push({
        rowNumber: CONFIG.FIRST_DATA_ROW + i,
        id: id,
        name: cell(i, 'name'),
        status: cell(i, 'status'),
        leadText: leadText,
        team: cell(i, 'team'),
        description: cell(i, 'description'),
        result: cell(i, 'result'),
        conversion: cell(i, 'conversion'),
        registration: cell(i, 'registration'),
        liveYmd: toYmd_(values[i][cols.liveDate - 1], tz),
        endYmd: cols.endDate ? toYmd_(values[i][cols.endDate - 1], tz) : null,
        link: trackerLinks[i][0] || planLinks[i][0] || '',
        jiraKeys: findJiraKeys_(cell(i, 'jira'), cell(i, 'plan'), planLinks[i]),
        emails: leadEmails[i].emails,
        emailSources: leadEmails[i].sources,
        skipped: isSkippedLead_(leadText),
      });
    }
  }
  const ctx = { ss: ss, sheet: sheet, tz: tz, cols: cols, rows: rows, lastCol: lastCol, today: today, weekStart: weekStart_(today) };
  ctx.jira = CONFIG.JIRA_ENABLED ? readJiraData_(ss) : { ok: false, missingTab: true, byKey: {} };
  return ctx;
}

function findColumns_(headers) {
  const norm = function (s) { return String(s).split('\n')[0].trim().toLowerCase(); };
  const index = {};
  headers.forEach(function (h, i) {
    const k = norm(h);
    if (k && !(k in index)) index[k] = i + 1;
  });
  const required = ['id', 'name', 'lead', 'status', 'liveDate'];
  const cols = {};
  Object.keys(CONFIG.COLUMNS).forEach(function (key) {
    const want = norm(CONFIG.COLUMNS[key]);
    let col = index[want];
    if (!col) {
      Object.keys(index).forEach(function (k) { if (!col && k.indexOf(want) === 0) col = index[k]; });
    }
    if (!col && required.indexOf(key) !== -1) {
      throw new Error('Could not find a column called "' + CONFIG.COLUMNS[key] + '" in row ' + CONFIG.HEADER_ROW +
        ' of "' + CONFIG.TAB_NAME + '". If it was renamed, update CONFIG.COLUMNS.');
    }
    cols[key] = col || null;
  });
  return cols;
}

/** Reads cells through the Google Sheets API (needed for name chips and file chips). Returns null if unavailable. */
function fetchApiCells_(ss, sheet, col, numRows) {
  if (typeof Sheets === 'undefined' || !col || numRows < 1) return null;
  try {
    const letter = colLetter_(col);
    const first = CONFIG.FIRST_DATA_ROW;
    const a1 = quoteSheet_(sheet.getName()) + '!' + letter + first + ':' + letter + (first + numRows - 1);
    const resp = Sheets.Spreadsheets.get(ss.getId(), {
      ranges: [a1],
      fields: 'sheets.data.rowData.values(chipRuns,hyperlink,textFormatRuns)',
    });
    const grid = resp && resp.sheets && resp.sheets[0] && resp.sheets[0].data && resp.sheets[0].data[0];
    const rowData = (grid && grid.rowData) || [];
    const out = [];
    for (let i = 0; i < numRows; i++) out.push((rowData[i] && rowData[i].values && rowData[i].values[0]) || {});
    return out;
  } catch (e) {
    CHIP_READ_ERROR_ = String(e && e.message ? e.message : e);
    Logger.log('Could not read cells through the Sheets API: ' + CHIP_READ_ERROR_);
    return null;
  }
}

/**
 * Finds the email(s) behind each Marketing Lead cell. Checks, in order:
 * name chips, mailto links, cell notes, emails typed in the cell, and the optional Leads tab.
 */
function readLeadEmails_(ss, sheet, col, numRows, display) {
  const out = [];
  for (let i = 0; i < numRows; i++) out.push({ emails: [], sources: [] });
  const add = function (i, email, source) {
    if (!email || !out[i]) return;
    email = String(email).replace(/^mailto:/i, '').split('?')[0].trim().toLowerCase();
    if (!EMAIL_RE_.test(email)) return;
    if (out[i].emails.indexOf(email) === -1) out[i].emails.push(email);
    if (out[i].sources.indexOf(source) === -1) out[i].sources.push(source);
  };
  const range = sheet.getRange(CONFIG.FIRST_DATA_ROW, col, numRows, 1);

  // 1. Name chips (needs the "Google Sheets API" service added in Apps Script)
  const api = fetchApiCells_(ss, sheet, col, numRows);
  if (api) {
    api.forEach(function (cell, i) {
      (cell.chipRuns || []).forEach(function (run) {
        const person = run && run.chip && run.chip.personProperties;
        if (person && person.email) add(i, person.email, 'name chip');
      });
      if (cell.hyperlink && /^mailto:/i.test(cell.hyperlink)) add(i, cell.hyperlink, 'link');
    });
  }

  // 2. mailto links
  try {
    range.getRichTextValues().forEach(function (row, i) {
      const rtv = row[0];
      if (!rtv) return;
      rtv.getRuns().forEach(function (run) {
        const url = run.getLinkUrl();
        if (url && /^mailto:/i.test(url)) add(i, url, 'link');
      });
    });
  } catch (e) {
    Logger.log('Could not read links: ' + e);
  }

  // 3. Cell notes
  range.getNotes().forEach(function (row, i) {
    (String(row[0] || '').match(EMAIL_FIND_RE_) || []).forEach(function (m) { add(i, m, 'note'); });
  });

  // 4. Email addresses typed straight into the cell
  display.forEach(function (row, i) {
    (String(row[col - 1] || '').match(EMAIL_FIND_RE_) || []).forEach(function (m) { add(i, m, 'typed'); });
  });

  // 5. Optional Leads tab (Name | Email)
  const leads = readLeadsTab_(ss);
  if (leads.length) {
    display.forEach(function (row, i) {
      const text = String(row[col - 1] || '').toLowerCase();
      leads.forEach(function (l) { if (text.indexOf(l.name) !== -1) add(i, l.email, CONFIG.LEADS_TAB + ' tab'); });
    });
  }
  return out;
}

/** Returns every web link in each cell of a column (file chips, hyperlinks, linked text). */
function readCellLinks_(ss, sheet, col, numRows) {
  const out = [];
  for (let i = 0; i < numRows; i++) out.push([]);
  if (!col || numRows < 1) return out;
  const add = function (i, url) {
    if (url && /^https?:/i.test(url) && out[i].indexOf(url) === -1) out[i].push(url);
  };
  const api = fetchApiCells_(ss, sheet, col, numRows);
  if (api) {
    api.forEach(function (cell, i) {
      (cell.chipRuns || []).forEach(function (run) {
        const rl = run && run.chip && run.chip.richLinkProperties;
        if (rl && rl.uri) add(i, rl.uri);
      });
      add(i, cell.hyperlink);
      (cell.textFormatRuns || []).forEach(function (run) {
        if (run && run.format && run.format.link) add(i, run.format.link.uri);
      });
    });
  }
  try {
    sheet.getRange(CONFIG.FIRST_DATA_ROW, col, numRows, 1).getRichTextValues().forEach(function (row, i) {
      const rtv = row[0];
      if (!rtv) return;
      rtv.getRuns().forEach(function (run) { add(i, run.getLinkUrl()); });
    });
  } catch (e) {
    Logger.log('Could not read links: ' + e);
  }
  return out;
}

function readLeadsTab_(ss) {
  const sh = findSheet_(ss, CONFIG.LEADS_TAB);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues()
    .map(function (r) { return { name: String(r[0]).trim().toLowerCase(), email: String(r[1]).trim() }; })
    .filter(function (r) { return r.name && r.email; });
}

function isSkippedLead_(leadText) {
  const t = String(leadText).trim().toLowerCase();
  return CONFIG.SKIP_LEADS.some(function (s) { return String(s).trim().toLowerCase() === t; });
}

/* =========================  LEAD EMAILS: WHAT TO SEND  ========================= */

function buildPlan_(ctx, logKeys) {
  const reminderStatuses = lowerList_(CONFIG.REMINDER_STATUSES);
  const checkStatuses = lowerList_(CONFIG.LAUNCH_CHECK_STATUSES);
  const plan = { byPerson: {}, skipped: [], noEmail: [], alreadySent: 0 };

  ctx.rows.forEach(function (r) {
    if (!r.liveYmd) return;
    const status = r.status.toLowerCase();
    const days = daysBetween_(ctx.today, r.liveYmd);
    let type = null;
    if (reminderStatuses.indexOf(status) !== -1 && days >= 0 && days <= CONFIG.REMINDER_DAYS_BEFORE) type = TYPE_REMINDER_;
    else if (checkStatuses.indexOf(status) !== -1 && days < 0) type = TYPE_CHECK_;
    if (!type) return;

    // Reminders: once per test per Live Date. Launch checks: once per week, every week.
    const key = type === TYPE_REMINDER_
      ? type + '|' + r.id + '|' + r.liveYmd
      : type + '|' + r.id + '|' + r.liveYmd + '|week ' + ctx.weekStart;
    const item = { type: type, row: r, days: days, key: key };
    if (r.skipped) { plan.skipped.push(item); return; }
    if (!r.emails.length) { plan.noEmail.push(item); return; }

    let anyNew = false;
    r.emails.forEach(function (email) {
      if (logKeys[key + '|' + email]) return;
      anyNew = true;
      (plan.byPerson[email] = plan.byPerson[email] || []).push(item);
    });
    if (!anyNew) plan.alreadySent++;
  });
  return plan;
}

function describePlan_(ctx, plan) {
  const lines = ['Week of ' + fmtYmd_(ctx.weekStart) + ' (' + fwLabel_(ctx.today) + ')', '', 'LEAD EMAILS'];
  const people = Object.keys(plan.byPerson);
  if (!people.length) lines.push('No lead emails this week.');
  people.forEach(function (email) {
    lines.push(email + '  (1 email)');
    sortItems_(plan.byPerson[email]).forEach(function (it) { lines.push('   • ' + itemLine_(it)); });
  });
  if (plan.noEmail.length) {
    lines.push('Not sent, no email found for the lead:');
    plan.noEmail.forEach(function (it) { lines.push('   • ' + it.row.id + ' (' + (it.row.leadText || 'no lead') + ')'); });
  }
  if (plan.skipped.length) {
    lines.push('Skipped, lead is on the skip list:');
    plan.skipped.forEach(function (it) { lines.push('   • ' + it.row.id + ' (' + it.row.leadText + ')'); });
  }
  if (plan.alreadySent) lines.push(plan.alreadySent + ' item(s) were already sent and won\'t be repeated.');
  return lines.join('\n');
}

function itemLine_(it) {
  const name = it.row.name.length > 45 ? it.row.name.slice(0, 44) + '…' : it.row.name;
  const base = it.row.id + ' ' + name + ', live ' + fmtYmd_(it.row.liveYmd) + ' (' + whenLabel_(it.days) + ')';
  return it.type === TYPE_REMINDER_ ? 'Launching within ' + windowLabel_() + ': ' + base : 'Did this launch? ' + base + ', still ' + it.row.status;
}

/* =========================  LEAD EMAILS: SENDING  ========================= */

/** mode: 'live' sends to leads and records it; 'test' sends everything to you and records nothing. */
function sendDigests_(ctx, plan, mode, logKeys) {
  const me = CONFIG.TEST_RECIPIENT || currentUserEmail_();
  const logRows = [];
  let sent = 0;

  Object.keys(plan.byPerson).forEach(function (email) {
    const items = sortItems_(plan.byPerson[email]);
    const note = mode === 'test' ? 'TEST: in live mode this email would go to <b>' + escapeHtml_(email) + '</b>.' : '';
    const msg = composeEmail_(ctx, items, note);
    const opts = {
      to: mode === 'test' ? me : email,
      subject: (mode === 'test' ? '[TEST] ' : '') + msg.subject,
      htmlBody: msg.html,
      body: msg.text,
      name: CONFIG.SENDER_NAME,
    };
    if (mode === 'live' && CONFIG.CC) opts.cc = CONFIG.CC;
    MailApp.sendEmail(opts);
    sent++;
    if (mode === 'live') {
      items.forEach(function (it) {
        logRows.push([new Date(), it.type, it.row.id, it.row.name, fmtYmd_(it.row.liveYmd), it.row.status, email, 'Sent', it.key + '|' + email]);
      });
    }
  });

  if (mode === 'live') {
    plan.noEmail.forEach(function (it) {
      const k = it.key + '|NO_EMAIL';
      if (logKeys[k]) return;
      logRows.push([new Date(), it.type, it.row.id, it.row.name, fmtYmd_(it.row.liveYmd), it.row.status, '',
        'Not sent: no email found for "' + (it.row.leadText || 'blank') + '"', k]);
    });
    appendLog_(ctx.ss, logRows);
  }
  return sent;
}

function composeEmail_(ctx, items, topNote) {
  const reminders = items.filter(function (it) { return it.type === TYPE_REMINDER_; });
  const checks = items.filter(function (it) { return it.type === TYPE_CHECK_; });
  const sheetUrl = repoUrl_(ctx);
  const windowText = windowLabel_();
  const showJira = ctx.jira.ok && items.some(function (it) { return it.row.jiraKeys.length; });

  const parts = [];
  if (reminders.length) parts.push(reminders.length + (reminders.length === 1 ? ' test' : ' tests') + ' launching within ' + windowText);
  if (checks.length) parts.push(checks.length + (checks.length === 1 ? ' test needs' : ' tests need') + ' a launch update');
  const subject = 'Experiment repo weekly check: ' + parts.join('; ');

  const heads = ['Test', 'Experiment', 'Live Date', 'When', 'Status'].concat(showJira ? ['Dev ticket (Jira)'] : []);
  const rowCells = function (it) {
    const cells = [
      '<a href="' + rowUrl_(ctx, it.row) + '">' + escapeHtml_(it.row.id) + '</a>',
      escapeHtml_(it.row.name), fmtYmd_(it.row.liveYmd), whenLabel_(it.days), escapeHtml_(it.row.status),
    ];
    if (showJira) cells.push(jiraCellHtml_(ctx, it));
    return cells;
  };

  let html = '<div style="font-family:Arial,Helvetica,sans-serif;color:#222;max-width:760px;">';
  if (topNote) html += '<div style="background:#FFF4CC;padding:8px 12px;margin-bottom:12px;font-size:13px;">' + topNote + '</div>';
  html += '<p style="font-size:14px;">Here\'s your weekly update from the <a href="' + sheetUrl + '">' +
    escapeHtml_(CONFIG.TAB_NAME) + ' experimentation repository</a>.</p>';
  let text = "Here's your weekly update from the " + CONFIG.TAB_NAME + ' experimentation repository.\n' + sheetUrl + '\n\n';

  if (reminders.length) {
    html += '<h3 style="font-size:15px;margin:18px 0 4px;">Launching within ' + windowText + '</h3>' +
      '<p style="font-size:13px;margin:0 0 8px;">Please check that test plans, QA and trackers are on track for these dates.</p>' +
      htmlTable_(heads, reminders.map(rowCells), 1);
    text += 'LAUNCHING WITHIN ' + windowText.toUpperCase() + '\n' + reminders.map(function (it) { return '- ' + textLine_(ctx, it); }).join('\n') + '\n\n';
  }
  if (checks.length) {
    html += '<h3 style="font-size:15px;margin:18px 0 4px;">Live Date has passed. Did these launch?</h3>' +
      '<p style="font-size:13px;margin:0 0 8px;">These are still marked ' + escapeHtml_(CONFIG.LAUNCH_CHECK_STATUSES.join(' / ')) +
      '. If the test launched, change its Status to <b>In-Flight</b>. If it slipped, update the <b>Live Date</b>. ' +
      'This list repeats each week until the sheet is updated.</p>' +
      htmlTable_(heads, checks.map(rowCells), 1);
    text += 'LIVE DATE HAS PASSED: DID THESE LAUNCH?\nIf launched, set Status to In-Flight. If it slipped, update the Live Date.\n' +
      checks.map(function (it) { return '- ' + textLine_(ctx, it); }).join('\n') + '\n\n';
  }

  const footer = 'Launch reminders are sent once, about three weeks ahead. Tests whose Live Date has passed are listed every week until the sheet is updated.';
  html += '<p style="font-size:12px;color:#777;margin-top:20px;">' + footer + '</p></div>';
  text += footer;
  return { subject: subject, html: html, text: text };
}

function textLine_(ctx, it) {
  let s = it.row.id + ' ' + it.row.name + ' | Live ' + fmtYmd_(it.row.liveYmd) + ' (' + whenLabel_(it.days) + ') | ' + it.row.status;
  if (ctx.jira.ok && it.row.jiraKeys.length) s += ' | Jira: ' + jiraSummaryText_(ctx, it.row);
  return s;
}

/* =========================  WEEKLY CALENDAR SUMMARY  ========================= */

/**
 * Sorts tests into the four slide sections for the current fiscal week:
 *  Recently Ended    End Date in the previous fiscal week or this week so far
 *  Recently Launched Live Date in the previous fiscal week or this week so far, Status In-Flight/Complete
 *  Upcoming          Live Date from tomorrow to the end of next fiscal week
 *  Delays            Live Date has passed but Status is still Planning / Design / In Development
 */
function buildSummary_(ctx) {
  const pre = lowerList_(CONFIG.REMINDER_STATUSES.concat(CONFIG.LAUNCH_CHECK_STATUSES));
  const checkStatuses = lowerList_(CONFIG.LAUNCH_CHECK_STATUSES);
  const closed = ['cancelled', 'canceled', 'duplicate'];
  const prevStart = addDays_(ctx.weekStart, -7);
  const nextEnd = addDays_(ctx.weekStart, 13);
  const s = { ended: [], launched: [], upcoming: [], delays: [] };

  ctx.rows.forEach(function (r) {
    const status = r.status.toLowerCase();
    if (!status || closed.indexOf(status) !== -1) return;
    const isPre = pre.indexOf(status) !== -1;
    if (r.endYmd && !isPre && r.endYmd >= prevStart && r.endYmd <= ctx.today) {
      s.ended.push({ row: r, text: 'Concluded ' + fwLabel_(r.endYmd) + '.' + resultText_(r) + descText_(r) });
    }
    if (!r.liveYmd) return;
    if (!isPre && r.liveYmd >= prevStart && r.liveYmd <= ctx.today) {
      s.launched.push({ row: r, text: 'Launched in ' + fwLabel_(r.liveYmd) + '.' + descText_(r) });
    } else if (r.liveYmd > ctx.today && r.liveYmd <= nextEnd && status !== 'complete') {
      s.upcoming.push({ row: r, text: 'Launching ' + fwLabel_(r.liveYmd) + '.' + descText_(r) });
    } else if (checkStatuses.indexOf(status) !== -1 && r.liveYmd < ctx.today) {
      s.delays.push({ row: r, needsInput: true,
        text: 'Planned for ' + fwLabel_(r.liveYmd) + ' (' + fmtYmd_(r.liveYmd) + ') but still marked ' + r.status + '. [Add the reason and new launch date.]' });
    }
  });
  const byDate = function (k) { return function (a, b) { return String(a.row[k]).localeCompare(String(b.row[k])); }; };
  s.ended.sort(byDate('endYmd')); s.launched.sort(byDate('liveYmd')); s.upcoming.sort(byDate('liveYmd')); s.delays.sort(byDate('liveYmd'));
  return s;
}

function summaryRecipients_() {
  const list = CONFIG.SUMMARY_RECIPIENTS.slice();
  if (CONFIG.SUMMARY_INCLUDE_OWNER) list.push(currentUserEmail_());
  const out = [];
  list.forEach(function (e) { e = String(e || '').trim().toLowerCase(); if (EMAIL_RE_.test(e) && out.indexOf(e) === -1) out.push(e); });
  return out;
}

function describeSummary_(ctx, s, logKeys) {
  const lines = ['', 'CALENDAR SUMMARY (' + fwLabel_(ctx.today) + ') to: ' + (summaryRecipients_().join(', ') || 'nobody, add SUMMARY_RECIPIENTS')];
  Object.keys(SECTION_STYLE_).forEach(function (k) {
    lines.push('   ' + SECTION_STYLE_[k].title + ': ' + (s[k].length ? s[k].map(function (it) { return it.row.id; }).join(', ') : 'nothing this week'));
  });
  return lines.join('\n');
}

function sendSummary_(ctx, s, mode, logKeys) {
  const recipients = mode === 'test' ? [CONFIG.TEST_RECIPIENT || currentUserEmail_()] : summaryRecipients_();
  const msg = composeSummary_(ctx, s, mode === 'test' ? 'TEST: in live mode this goes to <b>' + escapeHtml_(summaryRecipients_().join(', ') || '(no one yet)') + '</b>.' : '');
  const logRows = [];
  let sent = 0;
  recipients.forEach(function (email) {
    const key = TYPE_SUMMARY_ + '|week ' + ctx.weekStart + '|' + email;
    if (mode === 'live' && logKeys[key]) return;
    MailApp.sendEmail({ to: email, subject: (mode === 'test' ? '[TEST] ' : '') + msg.subject, htmlBody: msg.html, body: msg.text, name: CONFIG.SENDER_NAME });
    sent++;
    if (mode === 'live') logRows.push([new Date(), TYPE_SUMMARY_, '', 'Calendar summary ' + fwLabel_(ctx.today), '', '', email, 'Sent', key]);
  });
  if (mode === 'live') appendLog_(ctx.ss, logRows);
  return sent;
}

function composeSummary_(ctx, s, topNote) {
  const font = "'Public Sans',Arial,Helvetica,sans-serif";
  const border = '1px solid #B7B7B7';
  const fw = fwLabel_(ctx.today);
  const subject = 'Experimentation Calendar draft: ' + CONFIG.DEFAULT_TEAM + ', ' + fw;

  let html = '<div style="font-family:' + font + ';color:#222;max-width:900px;">';
  if (topNote) html += '<div style="background:#FFF4CC;padding:8px 12px;margin-bottom:12px;font-size:13px;">' + topNote + '</div>';
  html += '<p style="font-size:14px;margin:0 0 6px;"><b>' + escapeHtml_(CONFIG.DEFAULT_TEAM) + ' rows for the Experimentation Calendar slide, ' + fw +
    ' (week of ' + fmtYmd_(ctx.weekStart) + ')</b></p>' +
    '<p style="font-size:12px;color:#555;margin:0 0 12px;">Drafted from the <a href="' + repoUrl_(ctx) + '">repository</a>. ' +
    'Copy a cell into the matching row of the slide table, or copy the whole table. The wording comes from each test\'s Description, so give it a quick edit. ' +
    'Anything in [brackets] needs your input.</p>';

  html += '<table style="border-collapse:collapse;width:100%;font-family:' + font + ';font-size:13px;">';
  let text = subject + '\n\n';
  Object.keys(SECTION_STYLE_).forEach(function (k) {
    const st = SECTION_STYLE_[k];
    html += '<tr><td colspan="2" style="background:' + st.color + ';color:#FFFFFF;font-weight:bold;padding:4px 8px;border:' + border + ';">' + st.title + '</td></tr>';
    text += st.title.toUpperCase() + '\n';
    if (!s[k].length) {
      html += '<tr><td style="padding:6px 8px;border:' + border + ';width:110px;"></td><td style="padding:6px 8px;border:' + border + ';color:#888;font-style:italic;">Nothing from the repository this week</td></tr>';
      text += '(nothing this week)\n\n';
      return;
    }
    s[k].forEach(function (it) {
      const team = it.row.team || CONFIG.DEFAULT_TEAM;
      const link = it.row.link || rowUrl_(ctx, it.row);
      const body = escapeHtml_(it.text).replace(/\[([^\]]+)\]/g, '<span style="background:#FFF2CC;">[$1]</span>');
      html += '<tr><td style="padding:6px 8px;border:' + border + ';width:110px;vertical-align:middle;">' + escapeHtml_(team) + '</td>' +
        '<td style="padding:6px 8px;border:' + border + ';line-height:1.4;"><a href="' + link + '" style="color:#3197FE;font-weight:bold;">' +
        escapeHtml_(it.row.name) + '</a><b>:</b> ' + body + '</td></tr>';
      text += team + ' | ' + it.row.name + ': ' + it.text + '\n';
    });
    text += '\n';
  });
  html += '</table><p style="font-size:12px;color:#777;margin-top:14px;">Covers the ' + escapeHtml_(CONFIG.TAB_NAME) +
    ' tab only. Other teams\' rows still need to be added from their own sources.</p></div>';
  return { subject: subject, html: html, text: text };
}

function resultText_(r) {
  const val = function (v) { v = String(v || '').trim(); return v && !/^(n\/?a|tbd|-)$/i.test(v) ? v : ''; };
  const bits = [];
  if (val(r.conversion)) bits.push('conversion impact ' + val(r.conversion));
  if (val(r.registration)) bits.push('registration impact ' + val(r.registration));
  const result = val(r.result);
  if (!result && !bits.length) return ' [Add the result.]';
  return ' Result: ' + (result || 'see tracker') + (bits.length ? ' (' + bits.join(', ') + ')' : '') + '.';
}

/** Cleans a Description for the slide: drops side notes ("*Pending…", "6/17: …"), keeps the first two sentences. */
function descText_(r) {
  const lines = String(r.description || '').split(/\n+/).map(function (l) { return l.trim(); }).filter(function (l) {
    return l && l.charAt(0) !== '*' && !/^\d{1,2}\/\d{1,2}(\/\d{2,4})?\s*:/.test(l) && !/^\(.*\)$/.test(l);
  }).map(function (l) { return l.replace(/^updated for q\d\s*:\s*/i, ''); });
  let text = lines.join(' ').replace(/\s+/g, ' ').trim();
  if (!text) return ' [Add a one-line description.]';
  const sentences = text.split(/(?<=[.!?])\s+(?=[A-Z0-9"“‘(])/);
  text = sentences.slice(0, 2).join(' ');
  if (!/[.!?]$/.test(text)) text += '.';
  return ' ' + text;
}

/* =========================  JIRA (Jira Cloud for Sheets add-on)  ========================= */

/** Creates the Jira Data tab (and a Jira Ticket column if missing), then writes this week's query. */
function setUpJira() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ctx = loadContext_();
  const notes = [];

  if (!ctx.cols.jira) {
    const col = ctx.sheet.getLastColumn() + 1;
    ctx.sheet.getRange(CONFIG.HEADER_ROW, col).setValue(CONFIG.COLUMNS.jira).setFontWeight('bold');
    const prefill = ctx.rows.map(function (r) { return [r.jiraKeys.join(', ')]; });
    const numRows = ctx.sheet.getLastRow() - CONFIG.FIRST_DATA_ROW + 1;
    const values = [];
    for (let i = 0; i < numRows; i++) values.push(['']);
    ctx.rows.forEach(function (r, j) { values[r.rowNumber - CONFIG.FIRST_DATA_ROW] = prefill[j]; });
    if (numRows > 0) ctx.sheet.getRange(CONFIG.FIRST_DATA_ROW, col, numRows, 1).setValues(values);
    const filled = prefill.filter(function (p) { return p[0]; }).length;
    notes.push('Added a "' + CONFIG.COLUMNS.jira + '" column (column ' + colLetter_(col) + ') and filled in ' + filled +
      ' ticket(s) found in the Experiment Plan text.');
    ctx = loadContext_();
  }

  let sh = findSheet_(ss, CONFIG.JIRA_TAB);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.JIRA_TAB);
    notes.push('Created the "' + CONFIG.JIRA_TAB + '" tab.');
  } else {
    notes.push('The "' + sh.getName() + '" tab already existed, so it was reused.');
  }
  sh.getRange('A1').setValue('Jira query (rewritten by the script every week):').setFontWeight('bold');
  sh.getRange('A2').setValue('Filled in by the Jira Cloud for Sheets add-on. Do not type below row 3.').setFontStyle('italic').setFontColor('#666666');
  sh.getRange(JIRA_FORMULA_CELL_).setFormula('=IF($B$1="","No Jira tickets listed in the repository yet",JIRA($B$1,"key,summary,status,assignee,duedate",0,1000))');
  const keys = writeJiraQuery_(ctx, sh);
  notes.push('Query covers ' + keys.length + ' ticket(s)' + (keys.length ? ': ' + keys.join(', ') : '') + '.');
  showMessage_('Jira tab ready', notes.join('\n') + '\n\nGive it a minute to load, then click "4. Check Jira data".\n' +
    'If the tab shows #NAME?, the =JIRA() function is not turned on yet (see the setup guide).');
}

/** Runs on the weekly schedule (before the emails) to make the add-on fetch fresh statuses. */
function refreshJiraQuery() {
  if (!CONFIG.JIRA_ENABLED) return;
  const ctx = loadContext_();
  const sh = findSheet_(ctx.ss, CONFIG.JIRA_TAB);
  if (sh) writeJiraQuery_(ctx, sh);
}

function writeJiraQuery_(ctx, sh) {
  const closed = ['cancelled', 'canceled', 'duplicate', 'complete'];
  const keys = [];
  ctx.rows.forEach(function (r) {
    if (closed.indexOf(r.status.toLowerCase()) !== -1) return;
    r.jiraKeys.forEach(function (k) { if (keys.indexOf(k) === -1) keys.push(k); });
  });
  keys.sort();
  // The "created <=" date changes every day. It filters nothing out, but makes the add-on re-run the query.
  const jql = keys.length ? 'key in (' + keys.join(', ') + ') AND created <= "' + addDays_(ctx.today, 1) + '" ORDER BY key' : '';
  sh.getRange('B1').setValue(jql);
  return keys;
}

/** Reads the Jira Data tab into { ok, error, byKey }. Never throws. */
function readJiraData_(ss) {
  const out = { ok: false, error: '', missingTab: false, byKey: {}, count: 0 };
  try {
    const sh = findSheet_(ss, CONFIG.JIRA_TAB);
    if (!sh) { out.missingTab = true; return out; }
    const lastRow = sh.getLastRow();
    const lastCol = Math.max(1, sh.getLastColumn());
    const startRow = 4;
    if (lastRow < startRow) { out.error = 'The Jira tab is empty.'; return out; }
    const vals = sh.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getDisplayValues();
    const first = String(vals[0][0] || '').trim();
    if (!first || /^#|^loading|^error/i.test(first)) { out.error = 'The Jira tab shows "' + (first || 'nothing') + '".'; return out; }
    if (/^no jira tickets/i.test(first)) { out.ok = true; return out; }

    const norm = function (h) { return String(h).toLowerCase().replace(/[^a-z]/g, ''); };
    const head = vals[0].map(norm);
    const find = function (names) { for (let i = 0; i < head.length; i++) if (names.indexOf(head[i]) !== -1) return i; return -1; };
    const kCol = find(['key', 'issuekey', 'workitemkey']);
    const sCol = find(['status']);
    const aCol = find(['assignee']);
    const dCol = find(['duedate', 'due']);
    if (kCol === -1 || sCol === -1) {
      out.error = 'Could not find "Key" and "Status" headings in row 4 of the Jira tab (found: ' + vals[0].filter(String).join(', ') + ').';
      return out;
    }
    vals.slice(1).forEach(function (r) {
      const key = String(r[kCol] || '').trim();
      if (!key) return;
      out.byKey[key] = { status: String(r[sCol] || '').trim(), assignee: aCol !== -1 ? String(r[aCol] || '').trim() : '',
        due: dCol !== -1 ? String(r[dCol] || '').trim() : '' };
      out.count++;
    });
    out.ok = true;
  } catch (e) {
    out.error = String(e && e.message ? e.message : e);
  }
  return out;
}

function checkJiraData() {
  const ctx = loadContext_();
  const lines = [];
  const withKeys = ctx.rows.filter(function (r) { return r.jiraKeys.length; });
  lines.push('Tests with a Jira ticket in the repository: ' + withKeys.length);
  if (!ctx.cols.jira) lines.push('There is no "' + CONFIG.COLUMNS.jira + '" column yet. Run "3. Set up Jira tab".');
  if (ctx.jira.missingTab) {
    lines.push('No "' + CONFIG.JIRA_TAB + '" tab yet. Run "3. Set up Jira tab".');
  } else if (!ctx.jira.ok) {
    lines.push('Jira data is not available: ' + ctx.jira.error);
    lines.push('If it says #NAME?, turn on the =JIRA() function (Extensions → Jira Cloud for Sheets → Open → Settings).');
    lines.push('If it names a ticket key, that key is probably mistyped in the repository.');
  } else {
    lines.push('Tickets loaded from Jira: ' + ctx.jira.count);
    withKeys.forEach(function (r) {
      lines.push('   • ' + r.id + ': ' + r.jiraKeys.map(function (k) {
        const j = ctx.jira.byKey[k];
        return k + ' → ' + (j ? (j.status || 'no status') : 'not loaded yet');
      }).join('; '));
    });
  }
  lines.push('', 'Emails still go out if Jira data is unavailable; they just leave out the Jira column.');
  showMessage_('Jira data', lines.join('\n'));
}

function findJiraKeys_(jiraText, planText, planLinks) {
  const keys = [];
  const add = function (k) { if (keys.indexOf(k) === -1) keys.push(k); };
  (String(jiraText || '').toUpperCase().match(JIRA_KEY_RE_) || []).forEach(add);
  const prefixes = CONFIG.JIRA_PROJECT_KEYS.map(function (p) { return String(p).toUpperCase() + '-'; });
  const allowed = function (k) { return prefixes.some(function (p) { return k.indexOf(p) === 0; }); };
  (String(planText || '').match(JIRA_KEY_RE_) || []).filter(allowed).forEach(add);
  (planLinks || []).forEach(function (u) {
    const m = String(u).match(/atlassian\.net\/browse\/([A-Z][A-Z0-9_]+-\d+)/);
    if (m) add(m[1]);
  });
  return keys;
}

function jiraSummaryText_(ctx, row) {
  return row.jiraKeys.map(function (k) {
    const j = ctx.jira.byKey[k];
    if (!j) return k + ' (not found)';
    return k + ' ' + [j.status, j.assignee, j.due ? 'due ' + j.due : ''].filter(String).join(', ');
  }).join('; ');
}

function jiraCellHtml_(ctx, it) {
  if (!it.row.jiraKeys.length) return '<span style="color:#999;">None listed</span>';
  const done = lowerList_(CONFIG.JIRA_DONE_STATUSES);
  return it.row.jiraKeys.map(function (k) {
    const j = ctx.jira.byKey[k];
    if (!j) return escapeHtml_(k) + '<br><span style="color:#999;font-size:12px;">not found in Jira tab</span>';
    const isDone = done.indexOf(j.status.toLowerCase()) !== -1;
    const extra = [j.assignee, j.due ? 'due ' + j.due : ''].filter(String).join(', ');
    let s = '<b>' + escapeHtml_(k) + '</b><br><span style="font-size:12px;color:' + (isDone ? '#1E7B34' : '#444') + ';">' +
      escapeHtml_(j.status || 'no status') + (extra ? ' · ' + escapeHtml_(extra) : '') + '</span>';
    if (isDone && it.type === TYPE_CHECK_) s += '<br><span style="font-size:12px;color:#1E7B34;">Dev work is done, so it may have launched</span>';
    return s;
  }).join('<br>');
}

/* =========================  SHEET SETUP  ========================= */

function addHighlightRule_(ctx) {
  const first = CONFIG.FIRST_DATA_ROW;
  const live = '$' + colLetter_(ctx.cols.liveDate) + first;
  const status = '$' + colLetter_(ctx.cols.status) + first;
  const statusTests = CONFIG.LAUNCH_CHECK_STATUSES.map(function (s) { return 'TRIM(' + status + ')="' + s + '"'; }).join(',');
  const formula = '=AND(IFERROR(IF(ISNUMBER(' + live + '),' + live + ',DATEVALUE(' + live + '))<TODAY(),FALSE),OR(' + statusTests + '))';

  const sheet = ctx.sheet;
  const range = sheet.getRange(first, 1, sheet.getMaxRows() - first + 1, ctx.lastCol);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setBackground(CONFIG.HIGHLIGHT_COLOR)
    .setRanges([range])
    .build();
  const others = sheet.getConditionalFormatRules().filter(function (r) {
    const c = r.getBooleanCondition();
    const vals = c ? c.getCriteriaValues() : [];
    return !(vals.length && String(vals[0]) === formula);
  });
  sheet.setConditionalFormatRules([rule].concat(others));
  return formula;
}

function buildLaunchTab_(ctx) {
  const ss = ctx.ss;
  let sh = findSheet_(ss, CONFIG.LAUNCH_TAB);
  if (!sh) sh = ss.insertSheet(CONFIG.LAUNCH_TAB);
  else sh.clear();

  const src = quoteSheet_(ctx.sheet.getName()) + '!';
  const first = CONFIG.FIRST_DATA_ROW;
  const rng = function (col) { const c = colLetter_(col); return src + c + first + ':' + c; };

  const parts = [];
  const heads = [];
  const addCol = function (label, expr) { heads.push(label); parts.push(expr); return parts.length; };
  addCol('Identifier', rng(ctx.cols.id));
  addCol('Experiment Name', rng(ctx.cols.name));
  addCol('Marketing Lead', rng(ctx.cols.lead));
  if (ctx.cols.site) addCol('Site of Origin', rng(ctx.cols.site));
  addCol('Status', 's');
  const dateIdx = addCol('Live Date', 'd');
  if (ctx.cols.tracker) addCol('Test Tracker', rng(ctx.cols.tracker));

  const L = rng(ctx.cols.liveDate);
  const formula =
    '=LET(' +
      'd,ARRAYFORMULA(IF(ISNUMBER(' + L + '),' + L + ',IFERROR(DATEVALUE(' + L + '),""))),' +
      's,' + rng(ctx.cols.status) + ',' +
      'out,FILTER(HSTACK(' + parts.join(',') + '),' +
        rng(ctx.cols.id) + '<>"",d<>"",d>=TODAY()-7,d<=TODAY(),s<>"Cancelled",s<>"Duplicate"),' +
      'IFERROR(SORT(out,' + dateIdx + ',FALSE),"No tests with a Live Date in the last 7 days"))';

  sh.getRange(1, 1).setValue('Tests with a Live Date in the last 7 days').setFontWeight('bold').setFontSize(12);
  sh.getRange(2, 1).setValue('Updates automatically. Status is shown so you can spot tests not yet confirmed as In-Flight.')
    .setFontStyle('italic').setFontColor('#666666');
  sh.getRange(3, 1, 1, heads.length).setValues([heads]).setFontWeight('bold').setBackground('#F1F3F4');
  sh.getRange(4, 1).setFormula(formula);
  sh.getRange(4, dateIdx, Math.max(1, sh.getMaxRows() - 3), 1).setNumberFormat('M/d/yyyy');
  sh.setFrozenRows(3);
  sh.setColumnWidth(2, 320);
  return sh;
}

function ensureLogSheet_(ss) {
  let sh = findSheet_(ss, CONFIG.LOG_TAB);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.LOG_TAB);
    sh.getRange(1, 1, 1, LOG_HEADERS_.length).setValues([LOG_HEADERS_]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.hideColumns(LOG_HEADERS_.length); // the Key column is for the script only
  }
  return sh;
}

function readLogKeys_(ss) {
  const sh = findSheet_(ss, CONFIG.LOG_TAB);
  const keys = {};
  if (!sh || sh.getLastRow() < 2) return keys;
  sh.getRange(2, LOG_HEADERS_.length, sh.getLastRow() - 1, 1).getValues()
    .forEach(function (r) { if (r[0]) keys[String(r[0])] = true; });
  return keys;
}

function appendLog_(ss, rows) {
  if (!rows.length) return;
  const sh = ensureLogSheet_(ss);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, LOG_HEADERS_.length).setValues(rows);
}

function removeScheduleTriggers_() {
  const names = ['dailyRun', 'weeklyRun', 'refreshJiraQuery'];
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (names.indexOf(t.getHandlerFunction()) !== -1) { ScriptApp.deleteTrigger(t); n++; }
  });
  return n;
}

/* =========================  SMALL HELPERS  ========================= */

/** Finds a tab by name, ignoring capital letters and extra spaces ("Launched Last 7 days" matches). */
function findSheet_(ss, name) {
  const norm = function (s) { return String(s).replace(/\s+/g, ' ').trim().toLowerCase(); };
  const want = norm(name);
  const matches = ss.getSheets().filter(function (sh) { return norm(sh.getName()) === want; });
  return matches[0] || null;
}

function today_(tz) { return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'); }

/** Accepts real dates or text like 10/21/2026. Returns 'yyyy-MM-dd' or null (for NA, Q2, blank...). */
function toYmd_(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v.getTime()) ? null : Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  const m = String(v == null ? '' : v).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
  if (!m) return null;
  let y = Number(m[3]);
  if (y < 100) y += 2000;
  const mo = Number(m[1]);
  const d = Number(m[2]);
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  return y + '-' + pad2_(mo) + '-' + pad2_(d);
}

function daysBetween_(fromYmd, toYmd) {
  const a = fromYmd.split('-').map(Number);
  const b = toYmd.split('-').map(Number);
  return Math.round((Date.UTC(b[0], b[1] - 1, b[2]) - Date.UTC(a[0], a[1] - 1, a[2])) / 86400000);
}

function addDays_(ymd, n) {
  const p = ymd.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + n));
  return d.getUTCFullYear() + '-' + pad2_(d.getUTCMonth() + 1) + '-' + pad2_(d.getUTCDate());
}

/** Fiscal week number, counting from CONFIG.FISCAL_YEAR_START (FW1). Null if outside that fiscal year. */
function fiscalWeek_(ymd) {
  const days = daysBetween_(CONFIG.FISCAL_YEAR_START, ymd);
  if (days < 0) return null;
  const fw = Math.floor(days / 7) + 1;
  return fw <= 53 ? fw : null;
}

function fwLabel_(ymd) {
  const fw = fiscalWeek_(ymd);
  return fw ? 'FW' + fw : 'the week of ' + fmtYmd_(weekStart_(ymd));
}

/** First day of the fiscal week containing ymd. */
function weekStart_(ymd) {
  const days = daysBetween_(CONFIG.FISCAL_YEAR_START, ymd);
  return addDays_(ymd, -(((days % 7) + 7) % 7));
}

function fmtYmd_(ymd) {
  const p = ymd.split('-').map(Number);
  return p[1] + '/' + p[2] + '/' + p[0];
}

function whenLabel_(days) {
  if (days === 0) return 'Today';
  if (days === 1) return 'Tomorrow';
  if (days > 1) return 'In ' + days + ' days';
  if (days === -1) return '1 day ago';
  return (-days) + ' days ago';
}

function windowLabel_() {
  const d = CONFIG.REMINDER_DAYS_BEFORE;
  return d % 7 === 0 ? (d / 7) + (d === 7 ? ' week' : ' weeks') : d + ' days';
}

function sortItems_(items) {
  return items.slice().sort(function (a, b) {
    if (a.type !== b.type) return a.type === TYPE_REMINDER_ ? -1 : 1;
    return a.days - b.days;
  });
}

function htmlTable_(heads, rows, wrapCol) {
  const TH = 'text-align:left;padding:6px 10px;border-bottom:2px solid #ddd;font-size:13px;';
  const TD = 'padding:6px 10px;border-bottom:1px solid #eee;font-size:13px;vertical-align:top;';
  return '<table style="border-collapse:collapse;width:100%;margin-bottom:8px;"><tr>' +
    heads.map(function (h) { return '<th style="' + TH + '">' + h + '</th>'; }).join('') + '</tr>' +
    rows.map(function (cells) {
      return '<tr>' + cells.map(function (c, j) {
        return '<td style="' + TD + (j === wrapCol || j === 5 ? '' : 'white-space:nowrap;') + '">' + c + '</td>';
      }).join('') + '</tr>';
    }).join('') + '</table>';
}

function repoUrl_(ctx) { return ctx.ss.getUrl() + '#gid=' + ctx.sheet.getSheetId(); }
function rowUrl_(ctx, r) { return repoUrl_(ctx) + '&range=A' + r.rowNumber; }
function lowerList_(list) { return list.map(function (s) { return String(s).trim().toLowerCase(); }); }
function pad2_(n) { return (n < 10 ? '0' : '') + n; }
function quoteSheet_(name) { return "'" + String(name).replace(/'/g, "''") + "'"; }
function titleCase_(s) { s = String(s).toLowerCase(); return s.charAt(0).toUpperCase() + s.slice(1); }

function colLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function escapeHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function currentUserEmail_() {
  return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
}

function showMessage_(title, text) {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(title, text, ui.ButtonSet.OK);
  } catch (e) {
    Logger.log(title + '\n' + text); // when run from the schedule or the editor
  }
}
