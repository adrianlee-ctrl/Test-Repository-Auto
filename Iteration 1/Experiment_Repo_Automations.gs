/**
 * EXPERIMENTATION REPOSITORY AUTOMATIONS  (Audience Development tab)
 * -------------------------------------------------------------------
 * 1. 3-week reminder
 *    Emails the Marketing Lead when a test's Live Date is 21 days away or less.
 *    Sent once per test per Live Date. If the Live Date changes, a new reminder goes out.
 *
 * 2. Launch check (flag only, Status is never changed)
 *    Highlights any row whose Live Date has passed while Status is still
 *    Planning / Design / In Development, and emails the lead once to confirm.
 *
 * 3. "Launched Last 7 Days" tab
 *    A live formula listing tests with a Live Date in the past 7 days.
 *
 * Lead emails are read from the people chips (@-mentions) in the Marketing Lead
 * column. Cells with two people send to both. Each lead gets at most one combined
 * email per day, and only when there is something new.
 *
 * Settings you may want to change are all in CONFIG below.
 */

const CONFIG = {
  // Where the data lives
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
    site: 'Site of Origin',   // optional
    tracker: 'Test Tracker',  // optional
  },

  // 3-week reminder
  REMINDER_DAYS_BEFORE: 21,
  REMINDER_STATUSES: ['New', 'Backlog', 'Planning', 'Design', 'In Development'],

  // Launch check: flagged when the Live Date has passed and Status is still one of these
  LAUNCH_CHECK_STATUSES: ['Planning', 'Design', 'In Development'],
  HIGHLIGHT_COLOR: '#FCE8B2',

  // Leads to leave out for now (exact match on the Marketing Lead cell)
  SKIP_LEADS: ['GAIN'],

  // Email settings
  SENDER_NAME: 'Experimentation Repo Reminders',
  CC: '',              // optional, e.g. 'you@company.com'
  TEST_RECIPIENT: '',  // blank = whoever clicks "Send today's emails to me"
  DAILY_RUN_HOUR: 8,   // 8 = between 8 and 9am, in the spreadsheet's time zone
  PAUSED: false,       // set to true to stop all daily emails (anyone with edit access can do this)

  // Tabs this script creates
  LAUNCH_TAB: 'Launched Last 7 Days',
  LOG_TAB: 'Automation Log',
  LEADS_TAB: 'Leads',  // optional backup: Name in column A, Email in column B
};

const TYPE_REMINDER_ = 'Reminder';
const TYPE_CHECK_ = 'Launch check';
const EMAIL_RE_ = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const EMAIL_FIND_RE_ = /[^\s@<>(),;:"']+@[^\s@<>(),;:"']+\.[a-z]{2,}/gi;
const LOG_HEADERS_ = ['Timestamp', 'Type', 'Identifier', 'Experiment', 'Live Date', 'Status', 'Sent To', 'Result', 'Key'];
let CHIP_READ_ERROR_ = '';

/* =========================  MENU  ========================= */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Repo Automations')
    .addItem('1. Check lead emails', 'checkLeadEmails')
    .addItem('2. Set up highlighting + launch tab', 'setUpSheet')
    .addItem("3. Preview today's emails", 'previewToday')
    .addItem("4. Send today's emails to me (test)", 'sendTestEmailsToMe')
    .addSeparator()
    .addItem('Turn ON daily emails', 'turnOnDailyEmails')
    .addItem('Turn OFF daily emails', 'turnOffDailyEmails')
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
    if (r.skipped) {
      lines.push(r.leadText + '  →  skipped for now (on the skip list)');
    } else if (r.emails.length) {
      lines.push(r.leadText + '  →  ' + r.emails.join(', ') + '   [from ' + r.emailSources.join(', ') + ']');
    } else {
      lines.push(r.leadText + '  →  NO EMAIL FOUND');
    }
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

/** Shows what would be emailed today, without sending anything. */
function previewToday() {
  const ctx = loadContext_();
  const plan = buildPlan_(ctx, readLogKeys_(ctx.ss));
  showMessage_("Preview: today's emails", describePlan_(ctx, plan));
}

/** Sends today's emails to you instead of the leads. Nothing is recorded as sent. */
function sendTestEmailsToMe() {
  const ctx = loadContext_();
  const plan = buildPlan_(ctx, readLogKeys_(ctx.ss));
  const n = sendDigests_(ctx, plan, 'test', {});
  const me = CONFIG.TEST_RECIPIENT || currentUserEmail_();
  showMessage_('Test emails sent', n ? (n + ' test email(s) sent to ' + me + '. Leads were not emailed.') : 'Nothing to send today.');
}

function turnOnDailyEmails() {
  removeDailyTriggers_();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  ScriptApp.newTrigger('dailyRun').timeBased().everyDays(1).atHour(CONFIG.DAILY_RUN_HOUR).inTimezone(tz).create();
  showMessage_('Daily emails are ON',
    'The check runs every morning between ' + CONFIG.DAILY_RUN_HOUR + ':00 and ' + (CONFIG.DAILY_RUN_HOUR + 1) + ':00 (' + tz + ').\n' +
    'Emails are sent from ' + currentUserEmail_() + '.\n\n' +
    'Only one person should turn this on.');
}

function turnOffDailyEmails() {
  const n = removeDailyTriggers_();
  showMessage_('Daily emails are OFF', n ? 'Your daily schedule has been removed.' : 'You had no daily schedule turned on.');
}

/** Runs on the daily schedule. */
function dailyRun() {
  if (CONFIG.PAUSED) { Logger.log('Paused (CONFIG.PAUSED is true). No emails sent.'); return; }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { Logger.log('Another run is in progress. Skipping.'); return; }
  try {
    const ctx = loadContext_();
    const keys = readLogKeys_(ctx.ss);
    const plan = buildPlan_(ctx, keys);
    const n = sendDigests_(ctx, plan, 'live', keys);
    Logger.log('Daily run complete. Emails sent: ' + n);
  } finally {
    lock.releaseLock();
  }
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
  const rows = [];

  if (numRows > 0) {
    const range = sheet.getRange(CONFIG.FIRST_DATA_ROW, 1, numRows, lastCol);
    const values = range.getValues();
    const display = range.getDisplayValues();
    const leadEmails = readLeadEmails_(ss, sheet, cols.lead, numRows, display);

    for (let i = 0; i < numRows; i++) {
      const id = String(display[i][cols.id - 1]).trim();
      if (!id) continue;
      const leadText = String(display[i][cols.lead - 1]).trim();
      rows.push({
        rowNumber: CONFIG.FIRST_DATA_ROW + i,
        id: id,
        name: String(display[i][cols.name - 1]).trim(),
        status: String(display[i][cols.status - 1]).trim(),
        leadText: leadText,
        liveYmd: toYmd_(values[i][cols.liveDate - 1], tz),
        emails: leadEmails[i].emails,
        emailSources: leadEmails[i].sources,
        skipped: isSkippedLead_(leadText),
      });
    }
  }
  return { ss: ss, sheet: sheet, tz: tz, cols: cols, rows: rows, lastCol: lastCol, today: today_(tz) };
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

/**
 * Finds the email(s) behind each Marketing Lead cell. Checks, in order:
 * people chips, mailto links, cell notes, emails typed in the cell, and the optional Leads tab.
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
  const first = CONFIG.FIRST_DATA_ROW;
  const range = sheet.getRange(first, col, numRows, 1);

  // 1. People chips (needs the "Google Sheets API" service added in Apps Script)
  if (typeof Sheets !== 'undefined') {
    try {
      const letter = colLetter_(col);
      const a1 = quoteSheet_(sheet.getName()) + '!' + letter + first + ':' + letter + (first + numRows - 1);
      const resp = Sheets.Spreadsheets.get(ss.getId(), {
        ranges: [a1],
        fields: 'sheets.data.rowData.values(chipRuns,hyperlink)',
      });
      const grid = resp && resp.sheets && resp.sheets[0] && resp.sheets[0].data && resp.sheets[0].data[0];
      const rowData = (grid && grid.rowData) || [];
      rowData.forEach(function (r, i) {
        const cell = r && r.values && r.values[0];
        if (!cell) return;
        (cell.chipRuns || []).forEach(function (run) {
          const person = run && run.chip && run.chip.personProperties;
          if (person && person.email) add(i, person.email, 'name chip');
        });
        if (cell.hyperlink && /^mailto:/i.test(cell.hyperlink)) add(i, cell.hyperlink, 'link');
      });
    } catch (e) {
      CHIP_READ_ERROR_ = String(e && e.message ? e.message : e);
      Logger.log('Could not read name chips: ' + CHIP_READ_ERROR_);
    }
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

/* =========================  DECIDING WHAT TO SEND  ========================= */

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

    const item = { type: type, row: r, days: days, key: type + '|' + r.id + '|' + r.liveYmd };
    if (r.skipped) { plan.skipped.push(item); return; }
    if (!r.emails.length) { plan.noEmail.push(item); return; }

    let anyNew = false;
    r.emails.forEach(function (email) {
      if (logKeys[item.key + '|' + email]) return;
      anyNew = true;
      (plan.byPerson[email] = plan.byPerson[email] || []).push(item);
    });
    if (!anyNew) plan.alreadySent++;
  });
  return plan;
}

function describePlan_(ctx, plan) {
  const lines = ['Today: ' + fmtYmd_(ctx.today), ''];
  const people = Object.keys(plan.byPerson);
  if (!people.length) lines.push('No new emails to send today.');
  people.forEach(function (email) {
    lines.push(email + '  (1 email)');
    sortItems_(plan.byPerson[email]).forEach(function (it) { lines.push('   • ' + itemLine_(it)); });
    lines.push('');
  });
  if (plan.noEmail.length) {
    lines.push('Not sent, no email found for the lead:');
    plan.noEmail.forEach(function (it) { lines.push('   • ' + it.row.id + ' (' + (it.row.leadText || 'no lead') + ')'); });
    lines.push('');
  }
  if (plan.skipped.length) {
    lines.push('Skipped, lead is on the skip list:');
    plan.skipped.forEach(function (it) { lines.push('   • ' + it.row.id + ' (' + it.row.leadText + ')'); });
    lines.push('');
  }
  if (plan.alreadySent) lines.push(plan.alreadySent + ' test(s) were already emailed on an earlier day and won\'t be sent again.');
  return lines.join('\n');
}

function itemLine_(it) {
  const name = it.row.name.length > 45 ? it.row.name.slice(0, 44) + '…' : it.row.name;
  if (it.type === TYPE_REMINDER_) {
    return 'Reminder: ' + it.row.id + ' ' + name + ', live ' + fmtYmd_(it.row.liveYmd) + ' (' + whenLabel_(it.days) + ')';
  }
  return 'Launch check: ' + it.row.id + ' ' + name + ', live ' + fmtYmd_(it.row.liveYmd) +
    ' (' + whenLabel_(it.days) + '), still ' + it.row.status;
}

/* =========================  SENDING  ========================= */

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
  const sheetUrl = ctx.ss.getUrl() + '#gid=' + ctx.sheet.getSheetId();
  const rowUrl = function (r) { return sheetUrl + '&range=A' + r.rowNumber; };
  const windowText = windowLabel_();

  const parts = [];
  if (reminders.length) parts.push(reminders.length + (reminders.length === 1 ? ' test' : ' tests') + ' launching within ' + windowText);
  if (checks.length) parts.push(checks.length + (checks.length === 1 ? ' launch date' : ' launch dates') + ' passed, please confirm');
  const subject = 'Experiment repo: ' + parts.join('; ');

  const TH = 'style="text-align:left;padding:6px 10px;border-bottom:2px solid #ddd;font-size:13px;"';
  const TD = 'style="padding:6px 10px;border-bottom:1px solid #eee;font-size:13px;vertical-align:top;"';
  const table = function (heads, rows) {
    return '<table style="border-collapse:collapse;width:100%;margin-bottom:8px;"><tr>' +
      heads.map(function (h) { return '<th ' + TH + '>' + h + '</th>'; }).join('') + '</tr>' +
      rows.map(function (cells) {
        return '<tr>' + cells.map(function (c, j) {
          const td = j === 1 ? TD : TD.replace('vertical-align:top;', 'vertical-align:top;white-space:nowrap;');
          return '<td ' + td + '>' + c + '</td>';
        }).join('') + '</tr>';
      }).join('') +
      '</table>';
  };
  const idLink = function (r) { return '<a href="' + rowUrl(r) + '">' + escapeHtml_(r.id) + '</a>'; };

  let html = '<div style="font-family:Arial,Helvetica,sans-serif;color:#222;max-width:700px;">';
  if (topNote) html += '<div style="background:#FFF4CC;padding:8px 12px;margin-bottom:12px;font-size:13px;">' + topNote + '</div>';
  html += '<p style="font-size:14px;">Here\'s your update from the <a href="' + sheetUrl + '">' +
    escapeHtml_(CONFIG.TAB_NAME) + ' experimentation repository</a>.</p>';

  let text = "Here's your update from the " + CONFIG.TAB_NAME + ' experimentation repository.\n' + sheetUrl + '\n\n';

  if (reminders.length) {
    html += '<h3 style="font-size:15px;margin:18px 0 4px;">Launching within ' + windowText + '</h3>' +
      '<p style="font-size:13px;margin:0 0 8px;">Please check that test plans, QA and trackers are on track for these dates.</p>' +
      table(['Test', 'Experiment', 'Live Date', 'When', 'Status'], reminders.map(function (it) {
        return [idLink(it.row), escapeHtml_(it.row.name), fmtYmd_(it.row.liveYmd), whenLabel_(it.days), escapeHtml_(it.row.status)];
      }));
    text += 'LAUNCHING WITHIN ' + windowText.toUpperCase() + '\n' + reminders.map(function (it) {
      return '- ' + it.row.id + ' ' + it.row.name + ' | Live ' + fmtYmd_(it.row.liveYmd) + ' (' + whenLabel_(it.days) + ') | ' + it.row.status;
    }).join('\n') + '\n\n';
  }

  if (checks.length) {
    html += '<h3 style="font-size:15px;margin:18px 0 4px;">Live Date has passed. Did these launch?</h3>' +
      '<p style="font-size:13px;margin:0 0 8px;">These are still marked ' + escapeHtml_(CONFIG.LAUNCH_CHECK_STATUSES.join(' / ')) +
      '. If the test launched, change its Status to <b>In-Flight</b>. If it slipped, update the <b>Live Date</b>.</p>' +
      table(['Test', 'Experiment', 'Live Date', 'When', 'Status'], checks.map(function (it) {
        return [idLink(it.row), escapeHtml_(it.row.name), fmtYmd_(it.row.liveYmd), whenLabel_(it.days), escapeHtml_(it.row.status)];
      }));
    text += 'LIVE DATE HAS PASSED: DID THESE LAUNCH?\nIf launched, set Status to In-Flight. If it slipped, update the Live Date.\n' +
      checks.map(function (it) {
        return '- ' + it.row.id + ' ' + it.row.name + ' | Live ' + fmtYmd_(it.row.liveYmd) + ' (' + whenLabel_(it.days) + ') | ' + it.row.status;
      }).join('\n') + '\n\n';
  }

  const footer = "You get one email per test per Live Date. If a Live Date changes, you'll get a new reminder.";
  html += '<p style="font-size:12px;color:#777;margin-top:20px;">' + footer + '</p></div>';
  text += footer;
  return { subject: subject, html: html, text: text };
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

  // Replace an earlier copy of this rule if setup is run twice
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

function removeDailyTriggers_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyRun') { ScriptApp.deleteTrigger(t); n++; }
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

function lowerList_(list) { return list.map(function (s) { return String(s).trim().toLowerCase(); }); }
function pad2_(n) { return (n < 10 ? '0' : '') + n; }
function quoteSheet_(name) { return "'" + String(name).replace(/'/g, "''") + "'"; }

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
