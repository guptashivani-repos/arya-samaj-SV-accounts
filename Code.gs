/**
 * Arya Samaj Accounts API
 * Google Apps Script backend for the GitHub Pages PWA.
 *
 * REQUIRED Script Properties:
 * MAIN_SPREADSHEET_ID
 * MEMBER_SPREADSHEET_ID
 * APP_PIN_HASH  (SHA-256 hex of the PIN)
 */

const PROPS = PropertiesService.getScriptProperties();
const CONFIG = {
  MAIN_SPREADSHEET_ID: PROPS.getProperty('MAIN_SPREADSHEET_ID') || '',
  MEMBER_SPREADSHEET_ID: PROPS.getProperty('MEMBER_SPREADSHEET_ID') || '',
  SUBSCRIPTION_SHEET: PROPS.getProperty('SUBSCRIPTION_SHEET') || 'Annual Subscription',
  MEMBER_SHEET_NAME: PROPS.getProperty('MEMBER_SHEET_NAME') || '',
  TIMEZONE: 'Asia/Kolkata',
  MAX_LEDGER_ROWS: 200
};

/** JSONP API endpoint for GitHub Pages (avoids browser CORS/preflight issues). */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const callback = safeCallback_(p.callback || 'callback');
  try {
    validateConfig_();
    authorize_(p.auth);
    const action = clean_(p.action);
    const payload = p.payload ? JSON.parse(p.payload) : {};
    const result = routeApi_(action, payload);
    return jsonp_(callback, {ok: true, data: result});
  } catch (err) {
    return jsonp_(callback, {ok: false, error: err && err.message ? err.message : String(err)});
  }
}

function routeApi_(action, p) {
  switch (action) {
    case 'bootstrap': return getBootstrapData();
    case 'dashboard': return getDashboard(p);
    case 'ledger': return getLedger(p);
    case 'members': return searchMembers(p.query || '');
    case 'subscriptionStatus': return getSubscriptionStatus(p.financialYear || '', p.query || '');
    case 'addIncome': return addIncome(p);
    case 'addExpense': return addExpense(p);
    case 'addMember': return addMember(p);
    case 'addSubscription': return addSubscription(p);
    default: throw new Error('Unknown API action.');
  }
}

function authorize_(auth) {
  const expected = (PROPS.getProperty('APP_PIN_HASH') || '').toLowerCase();
  if (!expected) throw new Error('APP_PIN_HASH is not configured in Script Properties.');
  const supplied = String(auth || '').toLowerCase();
  if (!supplied || supplied !== expected) throw new Error('Incorrect PIN.');
}

function validateConfig_() {
  if (!CONFIG.MAIN_SPREADSHEET_ID) throw new Error('MAIN_SPREADSHEET_ID is not configured.');
  if (!CONFIG.MEMBER_SPREADSHEET_ID) throw new Error('MEMBER_SPREADSHEET_ID is not configured.');
}

function safeCallback_(name) {
  const value = String(name || 'callback');
  return /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(value) ? value : 'callback';
}

function jsonp_(callback, obj) {
  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(obj) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/** Run once from the Apps Script editor to set/replace the app PIN. */
function setAppPin(pin) {
  const value = String(pin || '').trim();
  if (value.length < 4) throw new Error('Use a PIN of at least 4 characters.');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8);
  const hex = digest.map(function(b) { return ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2); }).join('');
  PROPS.setProperty('APP_PIN_HASH', hex);
  return 'PIN configured.';
}

/** Optional helper: run once, then change the IDs if necessary. */
function configureAryaSamajApp() {
  PROPS.setProperties({
    MAIN_SPREADSHEET_ID: '1ZQcEyyqAHBB7YDMxwyWSOCQ5tasZKFIkkBtYNYL5vJA',
    MEMBER_SPREADSHEET_ID: '1K4ZskRGvAA4rwEJ0npiOShRgyGn5N9S-f2I7UP5SpM0',
    SUBSCRIPTION_SHEET: 'Annual Subscription'
  }, false);
  return 'Spreadsheet configuration saved.';
}

function getBootstrapData() {
  return {
    today: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    suggestedSunday: formatDate_(nextOrTodaySunday_(new Date())),
    financialYear: financialYear_(new Date()),
    paymentTypes: ['Cash', 'UPI', 'Cheque', 'Bank Transfer', 'Other'],
    incomeCategories: ['Donation', 'Hall Booking', 'Annual Function', 'Other'],
    expenseCategories: [
      'Donation', 'Honorarium to Mamta', 'Honorarium to Ramkishor Shastri',
      'Honorarium to Raju', 'Prashad', 'Ghee', 'Samagri / Samidha',
      'Electricity Bill', 'Gas Bill', 'Salary to Gardner', 'Repair Water Cooler',
      'Repair Harmonium', 'Miscellaneous', 'Other'
    ],
    dashboard: getDashboard({scope: 'fy', financialYear: financialYear_(new Date())})
  };
}

function addIncome(payload) {
  return withLock_(function () {
    validateTransaction_(payload, false);
    const date = parseDate_(payload.date);
    const sheet = ensureMonthlySheet_(date, 'Income');
    const receiptNo = clean_(payload.receiptNo);
    const description = clean_(payload.description);
    const amount = positiveNumber_(payload.amount, 'Amount');
    const payment = paymentLabel_(payload.paymentType, payload.paymentDetail);

    sheet.appendRow([date, receiptNo, description, amount, payment]);
    formatLastTransactionRow_(sheet, 5);
    clearAppCache_();
    return {ok: true, message: 'Income saved successfully', sheet: sheet.getName()};
  });
}

function addExpense(payload) {
  return withLock_(function () {
    validateTransaction_(payload, true);
    const date = parseDate_(payload.date);
    const sheet = ensureMonthlySheet_(date, 'Expense');
    const voucherNo = clean_(payload.receiptNo);
    const description = clean_(payload.description);
    const amount = positiveNumber_(payload.amount, 'Amount');
    const payment = paymentLabel_(payload.paymentType, payload.paymentDetail);
    const vendor = clean_(payload.vendor);

    sheet.appendRow([date, voucherNo, description, amount, payment, vendor]);
    formatLastTransactionRow_(sheet, 6);
    clearAppCache_();
    return {ok: true, message: 'Expense saved successfully', sheet: sheet.getName()};
  });
}



function addSubscription(payload) {
  return withLock_(function () {
    const memberName = clean_(payload.memberName);
    if (!memberName) throw new Error('Please select or enter the member name.');
    const date = parseDate_(payload.date);
    const receiptNo = clean_(payload.receiptNo);
    const amount = positiveNumber_(payload.amount, 'Subscription amount');
    const payment = paymentLabel_(payload.paymentType, payload.paymentDetail);
    const address = clean_(payload.address);
    const fy = clean_(payload.financialYear) || financialYear_(date);

    const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SUBSCRIPTION_SHEET);
    if (!sheet) sheet = ss.insertSheet(CONFIG.SUBSCRIPTION_SHEET);
    ensureSubscriptionHeaders_(sheet);

    const description = memberName + ' ' + fy;
    sheet.appendRow([date, receiptNo, description, amount, payment, address, memberName, fy]);
    const row = sheet.getLastRow();
    sheet.getRange(row, 1).setNumberFormat('dd-mmm-yyyy');
    sheet.getRange(row, 4).setNumberFormat('₹#,##0.00');
    clearAppCache_();
    return {ok: true, message: 'Annual subscription saved', financialYear: fy};
  });
}

function getDashboard(filter) {
  const f = normalizeFilter_(filter);
  const tx = getAllTransactions_();
  const filtered = tx.filter(r => inFilter_(r.dateObj, f));
  const regularIncome = sum_(filtered.filter(r => r.type === 'Income').map(r => r.amount));
  const subscriptions = sum_(filtered.filter(r => r.type === 'Subscription').map(r => r.amount));
  const expenses = sum_(filtered.filter(r => r.type === 'Expense').map(r => r.amount));
  const members = getMembers_('');

  const monthly = {};
  filtered.forEach(r => {
    const key = Utilities.formatDate(r.dateObj, CONFIG.TIMEZONE, 'MMM yyyy');
    if (!monthly[key]) monthly[key] = {income: 0, expense: 0, subscription: 0};
    if (r.type === 'Expense') monthly[key].expense += r.amount;
    else if (r.type === 'Subscription') monthly[key].subscription += r.amount;
    else monthly[key].income += r.amount;
  });

  const result = {
    regularIncome,
    subscriptions,
    totalIncome: regularIncome + subscriptions,
    expenses,
    balance: regularIncome + subscriptions - expenses,
    memberCount: members.length,
    recent: filtered.sort((a, b) => b.dateObj - a.dateObj).slice(0, 8).map(publicTx_),
    monthly: Object.keys(monthly).map(k => ({month: k, ...monthly[k]}))
  };
  return result;
}

function getLedger(filter) {
  const f = normalizeFilter_(filter);
  const query = normalize_(filter && filter.query);
  let rows = getAllTransactions_().filter(r => inFilter_(r.dateObj, f));
  if (query) {
    rows = rows.filter(r => normalize_([r.description, r.reference, r.payment, r.vendor, r.member].join(' ')).includes(query));
  }
  rows.sort((a, b) => b.dateObj - a.dateObj);
  return rows.slice(0, CONFIG.MAX_LEDGER_ROWS).map(publicTx_);
}

function searchMembers(query) {
  return getMembers_(query).slice(0, 80);
}

function getSubscriptionStatus(financialYear, query) {
  const fy = clean_(financialYear) || financialYear_(new Date());
  const q = normalize_(query);
  const members = getMembers_('');
  const subs = readSubscriptions_();
  const paidMap = {};
  subs.forEach(s => {
    if (s.financialYear === fy) paidMap[normalize_(s.member)] = (paidMap[normalize_(s.member)] || 0) + s.amount;
  });
  return members
    .filter(m => !q || normalize_([m.name, m.address].join(' ')).includes(q))
    .map(m => ({...m, paid: paidMap[normalize_(m.name)] || 0, financialYear: fy}))
    .slice(0, 250);
}

// ---------- Data readers ----------

function getAllTransactions_() {
  const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
  const out = [];

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (/\bIncome$/i.test(name) && name !== CONFIG.SUBSCRIPTION_SHEET) {
      readTransactionSheet_(sheet, 'Income', out);
    } else if (/\bExpense$/i.test(name)) {
      readTransactionSheet_(sheet, 'Expense', out);
    }
  });

  readSubscriptions_().forEach(s => out.push({
    type: 'Subscription', dateObj: s.dateObj, amount: s.amount,
    description: 'Annual Subscription', reference: s.receiptNo,
    payment: s.payment, vendor: '', member: s.member,
    source: CONFIG.SUBSCRIPTION_SHEET
  }));
  return out;
}

function readTransactionSheet_(sheet, type, out) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const cols = type === 'Expense' ? 6 : 5;
  const values = sheet.getRange(2, 1, lastRow - 1, cols).getValues();
  values.forEach(r => {
    const dateObj = coerceDate_(r[0]);
    const amount = Number(r[3]);
    if (!dateObj || !isFinite(amount) || amount <= 0) return;
    out.push({
      type, dateObj, reference: r[1] == null ? '' : String(r[1]),
      description: clean_(r[2]), amount,
      payment: clean_(r[4]), vendor: type === 'Expense' ? clean_(r[5]) : '',
      member: '', source: sheet.getName()
    });
  });
}

function readSubscriptions_() {
  const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SUBSCRIPTION_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  ensureSubscriptionHeaders_(sheet);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  return values.map(r => {
    const dateObj = coerceDate_(r[0]);
    const amount = Number(r[3]);
    if (!dateObj || !isFinite(amount) || amount <= 0) return null;
    const parsed = parseSubscriptionDescription_(clean_(r[2]));
    return {
      dateObj,
      receiptNo: r[1] == null ? '' : String(r[1]),
      description: clean_(r[2]), amount, payment: clean_(r[4]), address: clean_(r[5]),
      member: clean_(r[6]) || parsed.member,
      financialYear: clean_(r[7]) || parsed.financialYear || financialYear_(dateObj)
    };
  }).filter(Boolean);
}



// ---------- Sheet creation / preservation ----------

function ensureMonthlySheet_(date, type) {
  const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
  const candidates = monthlySheetCandidates_(date, type);
  let sheet = null;
  for (let i = 0; i < candidates.length; i++) {
    sheet = ss.getSheetByName(candidates[i]);
    if (sheet) break;
  }
  if (sheet) return sheet;

  const name = candidates[0];
  sheet = ss.insertSheet(name);
  if (type === 'Income') {
    sheet.getRange(1, 1, 1, 5).setValues([['Date', 'Receipt No.', 'Item description', 'Amount', 'Payment type']]);
    styleHeader_(sheet, 5);
    sheet.setColumnWidths(1, 5, 120);
    sheet.setColumnWidth(3, 260);
  } else {
    sheet.getRange(1, 1, 1, 6).setValues([['Date', 'Expense Voucher', 'Item Description', 'Amount', 'Payment type', 'Vendor']]);
    styleHeader_(sheet, 6);
    sheet.setColumnWidths(1, 6, 120);
    sheet.setColumnWidth(3, 260);
  }
  sheet.setFrozenRows(1);
  return sheet;
}

function monthlySheetCandidates_(date, type) {
  const month = Utilities.formatDate(date, CONFIG.TIMEZONE, 'MMMM');
  const year = Number(Utilities.formatDate(date, CONFIG.TIMEZONE, 'yyyy'));
  const legacy = [];

  // Existing 2025-26 workbook naming conventions.
  if (year === 2025) {
    legacy.push(month + ' ' + type);
    if (month === 'May' && type === 'Expense') legacy.push('may Expense');
  } else if (year === 2026 && ['January', 'February', 'March'].includes(month)) {
    legacy.push(month + ' ' + type);
  } else {
    legacy.push(month + String(year).slice(-2) + ' ' + type);
    legacy.push(month + ' ' + type);
  }
  return [...new Set(legacy)];
}

function ensureSubscriptionHeaders_(sheet) {
  const headers = ['Date', 'Receipt No.', 'Item description', 'Amount', 'Payment type', 'Address', 'Member Name', 'Financial Year'];
  const current = sheet.getRange(1, 1, 1, 8).getValues()[0];
  headers.forEach((h, i) => {
    if (!current[i]) sheet.getRange(1, i + 1).setValue(h);
  });
  sheet.setFrozenRows(1);
}

function ensureMemberHeaders_(sheet) {
  const headers = ['S.No.', 'Name', 'Address', 'Mobile'];
  const current = sheet.getRange(1, 1, 1, 4).getValues()[0];
  const looksLikeHeader = normalize_(current[1]).includes('name');
  if (!looksLikeHeader) {
    // The shared member list currently has title rows above the real header.
    const data = sheet.getDataRange().getValues();
    let headerRow = data.findIndex(r => normalize_(r[1]) === 'name' && normalize_(r[2]) === 'address') + 1;
    if (headerRow > 1) {
      // Keep the archive intact; use its actual header row rather than moving rows.
      PropertiesService.getScriptProperties().setProperty('MEMBER_HEADER_ROW', String(headerRow));
      if (!sheet.getRange(headerRow, 4).getValue()) sheet.getRange(headerRow, 4).setValue('Mobile');
      return;
    }
    sheet.getRange(1, 1, 1, 4).setValues([headers]);
  } else if (!current[3]) {
    sheet.getRange(1, 4).setValue('Mobile');
  }
}

function memberSheet_(ss) {
  return CONFIG.MEMBER_SHEET_NAME ? ss.getSheetByName(CONFIG.MEMBER_SHEET_NAME) : ss.getSheets()[0];
}

// getMembers_ / addMember need to respect an existing member header below title rows.
function memberHeaderRow_(sheet) {
  const remembered = Number(PropertiesService.getScriptProperties().getProperty('MEMBER_HEADER_ROW'));
  if (remembered > 0 && normalize_(sheet.getRange(remembered, 2).getValue()) === 'name') return remembered;
  const max = Math.min(sheet.getLastRow(), 20);
  if (max) {
    const rows = sheet.getRange(1, 1, max, Math.max(3, sheet.getLastColumn())).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (normalize_(rows[i][1]) === 'name' && normalize_(rows[i][2]) === 'address') {
        PropertiesService.getScriptProperties().setProperty('MEMBER_HEADER_ROW', String(i + 1));
        return i + 1;
      }
    }
  }
  return 1;
}

// Override member helpers to safely handle title rows in the existing member sheet.
function getMembers_(query) {
  const ss = SpreadsheetApp.openById(CONFIG.MEMBER_SPREADSHEET_ID);
  const sheet = memberSheet_(ss);
  ensureMemberHeaders_(sheet);
  const headerRow = memberHeaderRow_(sheet);
  if (sheet.getLastRow() <= headerRow) return [];
  const values = sheet.getRange(headerRow + 1, 1, sheet.getLastRow() - headerRow, Math.max(4, sheet.getLastColumn())).getValues();
  const q = normalize_(query);
  return values.map(r => ({
    serial: r[0] || '', name: clean_(r[1]), address: clean_(r[2]), mobile: clean_(r[3])
  })).filter(m => m.name && (!q || normalize_([m.name, m.address, m.mobile].join(' ')).includes(q)));
}

function addMember(payload) {
  return withLock_(function () {
    const name = clean_(payload.name);
    if (!name) throw new Error('Member name is required.');
    const address = clean_(payload.address);
    const mobile = clean_(payload.mobile);
    const ss = SpreadsheetApp.openById(CONFIG.MEMBER_SPREADSHEET_ID);
    const sheet = memberSheet_(ss);
    ensureMemberHeaders_(sheet);
    const headerRow = memberHeaderRow_(sheet);
    const members = getMembers_('');
    const duplicate = members.some(m => normalize_(m.name) === normalize_(name) && normalize_(m.address) === normalize_(address));
    if (duplicate) throw new Error('This member already appears in the member list.');
    const serials = members.map(m => Number(m.serial)).filter(Number.isFinite);
    const nextSerial = serials.length ? Math.max.apply(null, serials) + 1 : 1;
    const row = Math.max(sheet.getLastRow() + 1, headerRow + 1);
    sheet.getRange(row, 1, 1, 4).setValues([[nextSerial, name, address, mobile]]);
    clearAppCache_();
    return {ok: true, message: 'New member added', serial: nextSerial};
  });
}

// ---------- Utilities ----------

function normalizeFilter_(filter) {
  filter = filter || {};
  return {
    scope: filter.scope || 'fy',
    financialYear: clean_(filter.financialYear) || financialYear_(new Date()),
    month: clean_(filter.month)
  };
}

function inFilter_(date, f) {
  if (!date) return false;
  if (f.scope === 'all') return true;
  if (f.scope === 'month' && f.month) return Utilities.formatDate(date, CONFIG.TIMEZONE, 'yyyy-MM') === f.month;
  return financialYear_(date) === f.financialYear;
}

function financialYear_(date) {
  const y = Number(Utilities.formatDate(date, CONFIG.TIMEZONE, 'yyyy'));
  const m = Number(Utilities.formatDate(date, CONFIG.TIMEZONE, 'M'));
  const start = m >= 4 ? y : y - 1;
  return start + '-' + String(start + 1).slice(-2);
}

function parseSubscriptionDescription_(text) {
  const match = text.match(/\b(20\d{2})\s*[-–]\s*(\d{2,4})\b/);
  let fy = '';
  let member = text;
  if (match) {
    fy = match[1] + '-' + String(match[2]).slice(-2);
    member = text.slice(0, match.index).replace(/[\s,]+$/, '');
  }
  return {member: member || text, financialYear: fy};
}

function validateTransaction_(p, isExpense) {
  if (!p) throw new Error('No data received.');
  if (!clean_(p.date)) throw new Error('Date is required.');
  if (!clean_(p.description)) throw new Error('Item description is required.');
  positiveNumber_(p.amount, 'Amount');
}

function parseDate_(value) {
  const s = clean_(value);
  const d = new Date(s + 'T12:00:00');
  if (!s || isNaN(d.getTime())) throw new Error('Please enter a valid date.');
  return d;
}

function coerceDate_(value) {
  if (value instanceof Date && !isNaN(value)) return value;
  if (!value) return null;
  const d = new Date(value);
  return isNaN(d) ? null : d;
}

function formatDate_(date) {
  return Utilities.formatDate(date, CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function nextOrTodaySunday_(date) {
  const d = new Date(date);
  const day = Number(Utilities.formatDate(d, CONFIG.TIMEZONE, 'u')); // Mon=1..Sun=7
  d.setDate(d.getDate() + ((7 - day) % 7));
  return d;
}

function publicTx_(r) {
  return {
    type: r.type,
    date: Utilities.formatDate(r.dateObj, CONFIG.TIMEZONE, 'dd MMM yyyy'),
    isoDate: Utilities.formatDate(r.dateObj, CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    reference: r.reference || '', description: r.description || '',
    amount: r.amount || 0, payment: r.payment || '', vendor: r.vendor || '',
    member: r.member || '', source: r.source || ''
  };
}

function paymentLabel_(type, detail) {
  const t = clean_(type);
  const d = clean_(detail);
  if (!t) return d;
  if (t === 'Cheque' && d) return 'Cheque ' + d;
  if (t === 'Other' && d) return d;
  return d ? t + ' - ' + d : t;
}

function positiveNumber_(value, label) {
  const n = Number(value);
  if (!isFinite(n) || n <= 0) throw new Error(label + ' must be greater than zero.');
  return n;
}

function clean_(v) {
  return v == null ? '' : String(v).trim();
}

function normalize_(v) {
  return clean_(v).toLowerCase().replace(/\s+/g, ' ');
}

function sum_(arr) {
  return arr.reduce((a, b) => a + (Number(b) || 0), 0);
}

function nextSerial_(rows) {
  const nums = rows.map(r => Number(r[0])).filter(Number.isFinite);
  return nums.length ? Math.max.apply(null, nums) + 1 : 1;
}

function styleHeader_(sheet, cols) {
  sheet.getRange(1, 1, 1, cols)
    .setFontWeight('bold')
    .setBackground('#7b2d26')
    .setFontColor('#ffffff');
}

function formatLastTransactionRow_(sheet, cols) {
  const row = sheet.getLastRow();
  sheet.getRange(row, 1).setNumberFormat('dd-mmm-yyyy');
  sheet.getRange(row, 4).setNumberFormat('₹#,##0.00');
  sheet.getRange(row, 1, 1, cols).setFontSize(11);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function clearAppCache_() {
  // Dashboard is calculated live so newly saved entries appear immediately.
}
