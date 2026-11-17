/**
 * BILL TRACKER — Google Apps Script web app (multi-person)
 * Data is stored in the Google Sheet this script is attached to
 * (tabs: Bills, Payments, Paychecks, Settings). Run setup() once, then deploy as a web app.
 *
 * ===== FIRST LOGIN =====
 * In Project Settings (⚙️) → Script Properties add:
 *   APP_USERNAME = your username
 *   APP_PASSWORD = your password
 * Your first login creates your account (as the admin), stores the password as a
 * secure hash and deletes the readable APP_PASSWORD automatically.
 * Add other people (e.g. your wife) in the app: Settings → People.
 *
 * Forgot your password? Add APP_PASSWORD again in Script Properties and run resetLogin().
 */
const SESSION_DAYS = 90;        // how long "Keep me logged in" lasts

const BILLS = 'Bills';
const PAYMENTS = 'Payments';
const SETTINGS = 'Settings';
const PAYCHECKS = 'Paychecks';

const BILL_COLS = ['id', 'name', 'amount', 'dueDay', 'category', 'frequency', 'startMonth', 'endMonth',
  'variable', 'autopay', 'trackBalance', 'balance', 'originalBalance', 'minPayment', 'interestRate',
  'url', 'notes', 'active', 'owner', 'startDate', 'every', 'unit', 'createdAt'];
const PAY_COLS = ['id', 'billId', 'month', 'amount', 'paidDate', 'confirmation', 'note', 'status', 'paidBy', 'createdAt'];
const SET_COLS = ['key', 'value'];
const PC_COLS = ['id', 'payDate', 'amount', 'note', 'allocations', 'owner', 'createdAt'];

// Shared settings (everyone). Income, email and reminder on/off are per person.
const DEFAULT_SETTINGS = {
  currency: '$',
  reminderDays: '3',
  reminderHour: '8',
  skipAutopayReminders: 'false',
  categories: JSON.stringify(['Housing', 'Utilities', 'Phone & Internet', 'Insurance', 'Credit Card', 'Loan', 'Car',
    'Subscriptions', 'Medical', 'Childcare', 'Savings', 'Other'])
};
const GLOBAL_KEYS = Object.keys(DEFAULT_SETTINGS);

/* ======================================================================
   Web app entry
   ====================================================================== */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Bills')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * JSON API for the full-screen version of the app (hosted on GitHub Pages).
 * Same security as the built-in page: every action still needs a valid login token.
 */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    let result;
    if (req.action === 'login') result = login(req.username, req.password);
    else if (req.action === 'logout') result = logout(req.token);
    else result = api(req.token, req.fn, req.args || []);
    out = { ok: true, result: result };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** Run this once from the editor to create the tabs and approve permissions. */
function setup() {
  sheet_(BILLS, BILL_COLS);
  sheet_(PAYMENTS, PAY_COLS);
  sheet_(SETTINGS, SET_COLS);
  sheet_(PAYCHECKS, PC_COLS);
  const ss = ss_();
  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  Logger.log('Ready! Data lives here: ' + ss.getUrl());
}

/* ======================================================================
   API — everything the app does goes through api(), which checks the login
   and passes the logged-in person to each action.
   ====================================================================== */

const API = {
  getData: getData_, saveBill: saveBill_, deleteBill: deleteBill_, markPaid: markPaidOne_, markPaidBatch: markPaidBatch_,
  unmarkPaid: unmarkPaid_, unmarkPaidBatch: unmarkPaidBatch_, updateBalance: updateBalance_, saveSettings: saveSettings_, sendTestReminder: sendTestReminder_,
  savePaycheck: savePaycheck_, deletePaycheck: deletePaycheck_, payAndApply: payAndApply_, changeLogin: changeLogin_,
  addPerson: addPerson_, removePerson: removePerson_, resetPersonPassword: resetPersonPassword_
};

function api(token, fn, args) {
  const user = checkToken_(token);
  if (!user) throw new Error('AUTH');
  if (!Object.prototype.hasOwnProperty.call(API, fn)) throw new Error('Unknown action');
  return API[fn].apply(null, [user].concat(args || []));
}

/* ---------- Visibility rules ---------- */

// A bill is "shared" (household) or belongs to one person (owner = their id).
function canSee_(user, b) { return !b.owner || b.owner === 'shared' || b.owner === user.id; }
function pcOwner_(pc) { return pc.owner || adminId_(); }

function visibleBills_(user) {
  return read_(BILLS, BILL_COLS).filter(b => canSee_(user, b));
}
function findBill_(user, id) {
  const b = read_(BILLS, BILL_COLS).find(x => x.id === id);
  if (!b || !canSee_(user, b)) throw new Error('Bill not found');
  return b;
}

/* ---------- Data ---------- */

function getData_(user) {
  const bills = visibleBills_(user).map(billOut_);
  const ids = new Set(bills.map(b => b.id));
  const payments = read_(PAYMENTS, PAY_COLS).filter(p => ids.has(p.billId)).map(payOut_);
  // Your paychecks in full; other people's only show which shared bills they cover.
  const paychecks = read_(PAYCHECKS, PC_COLS).map(pcOut_).map(pc => {
    if (pc.owner === user.id) return pc;
    return {
      id: pc.id, payDate: pc.payDate, amount: 0, note: '', owner: pc.owner, readOnly: true,
      allocations: pc.allocations.filter(a => a.type === 'bill' && ids.has(a.billId))
        .map(a => ({ type: 'bill', billId: a.billId, month: a.month, amount: 0 }))
    };
  }).filter(pc => !pc.readOnly || pc.allocations.length);
  const settings = Object.assign(getSettings_(), {
    income: String(user.income || 0),
    email: user.email || '',
    emailReminders: String(!!user.emailReminders)
  });
  return {
    bills: bills, payments: payments, paychecks: paychecks, settings: settings,
    me: { id: user.id, name: user.name, username: user.username, admin: !!user.admin },
    people: users_().map(u => ({ id: u.id, name: u.name, username: user.admin ? u.username : undefined, admin: !!u.admin })),
    sheetUrl: user.admin ? ss_().getUrl() : ''
  };
}

/* ---------- Bills ---------- */

function saveBill_(user, b) {
  return withLock_(() => {
    if (!b || !String(b.name || '').trim()) throw new Error('Bill name is required');
    const sh = sheet_(BILLS, BILL_COLS);
    const owner = b.owner === 'me' ? user.id : 'shared';
    let rowNum = null;
    if (b.id) {
      const ex = findBill_(user, b.id);
      rowNum = ex._row;
      b = Object.assign({}, ex, b);
    } else {
      b.id = Utilities.getUuid();
      b.createdAt = new Date().toISOString();
      if (b.active === undefined) b.active = true;
    }
    b.owner = owner;
    writeRow_(sh, rowNum, b);
    addCategory_(b.category);
    return getData_(user);
  });
}

function deleteBill_(user, id) {
  return withLock_(() => {
    findBill_(user, id);
    deleteWhere_(PAYMENTS, PAY_COLS, r => r.billId === id);
    deleteWhere_(BILLS, BILL_COLS, r => r.id === id);
    return getData_(user);
  });
}

function updateBalance_(user, billId, balance) {
  return withLock_(() => {
    const b = findBill_(user, billId);
    b.balance = round_(num_(balance));
    if (!num_(b.originalBalance) || num_(b.balance) > num_(b.originalBalance)) b.originalBalance = b.balance;
    writeRow_(sheet_(BILLS, BILL_COLS), b._row, b);
    return getData_(user);
  });
}

/* ---------- Payments ---------- */

function markPaidOne_(user, p) {
  return withLock_(() => { markPaid_(user, p); return getData_(user); });
}

function markPaidBatch_(user, list) {
  return withLock_(() => { (list || []).forEach(p => markPaid_(user, p)); return getData_(user); });
}

function markPaid_(user, p) {
  findBill_(user, p.billId);
  const sh = sheet_(PAYMENTS, PAY_COLS);
  const existing = read_(PAYMENTS, PAY_COLS).find(x => x.billId === p.billId && x.month === p.month);
  const status = p.status === 'skipped' ? 'skipped' : 'paid';
  const amount = status === 'skipped' ? 0 : num_(p.amount);
  const oldAmount = existing && existing.status !== 'skipped' ? num_(existing.amount) : 0;
  const rec = {
    id: existing ? existing.id : Utilities.getUuid(),
    billId: p.billId, month: p.month, amount: amount,
    paidDate: p.paidDate || today_(),
    confirmation: p.confirmation || '', note: p.note || '',
    status: status, paidBy: user.id,
    createdAt: existing ? existing.createdAt : new Date().toISOString()
  };
  writeRow_(sh, existing ? existing._row : null, rec);
  adjustBalance_(p.billId, -(amount - oldAmount));
}

/** Mark several bills unpaid at once: list = [{billId, month}] */
function unmarkPaidBatch_(user, list) {
  return withLock_(() => { (list || []).forEach(x => unmark_(user, x.billId, x.month)); return getData_(user); });
}

function unmarkPaid_(user, billId, month) {
  return withLock_(() => { unmark_(user, billId, month); return getData_(user); });
}

function unmark_(user, billId, month) {
  {
    findBill_(user, billId);
    const sh = sheet_(PAYMENTS, PAY_COLS);
    const ex = read_(PAYMENTS, PAY_COLS).find(x => x.billId === billId && x.month === month);
    if (ex) {
      sh.deleteRow(ex._row);
      if (ex.status !== 'skipped') adjustBalance_(billId, num_(ex.amount));
    }
  }
}

/* ---------- Paychecks (each person has their own) ---------- */

function savePaycheck_(user, pc) {
  return withLock_(() => {
    if (!pc || !pc.payDate) throw new Error('Pay date is required');
    const sh = sheet_(PAYCHECKS, PC_COLS);
    const ex = pc.id ? read_(PAYCHECKS, PC_COLS).find(x => x.id === pc.id) : null;
    if (ex && pcOwner_(ex) !== user.id) throw new Error('That paycheck belongs to someone else');
    const ids = new Set(visibleBills_(user).map(b => b.id));
    const allocations = (pc.allocations || []).filter(a => a.type === 'custom' || ids.has(a.billId));
    const rec = {
      id: ex ? ex.id : Utilities.getUuid(),
      payDate: pc.payDate, amount: round_(num_(pc.amount)), note: pc.note || '',
      allocations: JSON.stringify(allocations), owner: user.id,
      createdAt: ex ? ex.createdAt : new Date().toISOString()
    };
    writeRow_(sh, ex ? ex._row : null, rec);
    return getData_(user);
  });
}

/**
 * Mark a bill paid AND take it out of one of your paychecks.
 * pcId = an existing paycheck of yours, or null with newPc = {payDate, amount} to create one.
 * The bill is removed from any other paycheck of yours for that month, so it's never counted twice.
 */
function payAndApply_(user, rec, pcId, newPc) {
  return withLock_(() => {
    const sh = sheet_(PAYCHECKS, PC_COLS);
    if (pcId && !read_(PAYCHECKS, PC_COLS).some(r => r.id === pcId && pcOwner_(r) === user.id)) throw new Error('Paycheck not found');
    markPaid_(user, rec);
    if (!pcId && newPc && newPc.payDate) {
      pcId = Utilities.getUuid();
      writeRow_(sh, null, { id: pcId, payDate: newPc.payDate, amount: round_(num_(newPc.amount)), note: '',
        allocations: '[]', owner: user.id, createdAt: new Date().toISOString() });
    }
    if (!pcId) return getData_(user);
    const mine = read_(PAYCHECKS, PC_COLS).filter(r => pcOwner_(r) === user.id);
    if (!mine.some(r => r.id === pcId)) throw new Error('Paycheck not found');
    mine.forEach(r => {
      let a = [];
      try { a = JSON.parse(r.allocations || '[]'); } catch (e) { a = []; }
      const before = a.length;
      a = a.filter(x => !(x.type === 'bill' && x.billId === rec.billId && x.month === rec.month));
      if (r.id === pcId) a.push({ type: 'bill', billId: rec.billId, month: rec.month, amount: round_(num_(rec.amount)) });
      else if (a.length === before) return;
      r.allocations = JSON.stringify(a);
      writeRow_(sh, r._row, r);
    });
    return getData_(user);
  });
}

function deletePaycheck_(user, id) {
  return withLock_(() => {
    deleteWhere_(PAYCHECKS, PC_COLS, r => r.id === id && pcOwner_(r) === user.id);
    return getData_(user);
  });
}

/* ---------- Settings ---------- */

function saveSettings_(user, s) {
  return withLock_(() => {
    s = s || {};
    // shared settings
    const sh = sheet_(SETTINGS, SET_COLS);
    const cur = getSettings_();
    GLOBAL_KEYS.forEach(k => { if (s[k] !== undefined) cur[k] = String(s[k]); });
    const rows = Object.keys(cur).map(k => [k, String(cur[k])]);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
    sh.getRange(2, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
    // personal settings
    updateUser_(user.id, u => {
      if (s.income !== undefined) u.income = num_(s.income);
      if (s.email !== undefined) u.email = String(s.email).trim();
      if (s.emailReminders !== undefined) u.emailReminders = bool_(s.emailReminders);
      if (s.name !== undefined && String(s.name).trim()) u.name = String(s.name).trim().slice(0, 40);
    });
    configureTrigger_(users_().some(u => u.emailReminders && u.email), parseInt(cur.reminderHour, 10) || 8);
    return getData_(checkUserById_(user.id));
  });
}

/* ======================================================================
   Email reminders (daily trigger, one email per person who turned them on)
   ====================================================================== */

function sendReminders() {
  users_().filter(u => u.emailReminders && u.email).forEach(u => {
    try { sendRemindersFor_(u, false); } catch (e) { console.error(e); }
  });
}

function sendTestReminder_(user) {
  const fresh = checkUserById_(user.id);
  if (!fresh.email) throw new Error('Add your email address in Settings first.');
  const any = sendRemindersFor_(fresh, true);
  return (any ? 'Sent to ' : 'Nothing due soon — test sent to ') + fresh.email;
}

function sendRemindersFor_(user, force) {
  const s = getSettings_();
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const ym = todayStr.slice(0, 7);
  const days = parseInt(s.reminderDays, 10) || 3;
  const skipAuto = bool_(s.skipAutopayReminders);
  const cur = s.currency || '$';
  const bills = visibleBills_(user).map(billOut_).filter(b => b.active);
  const pays = read_(PAYMENTS, PAY_COLS);
  const isDone = (id, m) => pays.some(p => p.billId === id && p.month === m);

  const items = [];
  [addMonths_(ym, -2), addMonths_(ym, -1), ym, addMonths_(ym, 1)].forEach(m => {
    bills.forEach(b => occurrences_(b, m).forEach(o => {
      if (isDone(b.id, o.key)) return;
      const due = o.due;
      const diff = daysBetween_(todayStr, due);
      if (skipAuto && b.autopay && diff >= 0) return;
      if (diff <= days) items.push({ b: b, due: due, diff: diff });
    }));
  });

  if (!items.length && force !== true) return false;
  items.sort((a, b) => a.diff - b.diff);

  const url = ScriptApp.getService().getUrl();
  const email = reminderEmail_(user, items, cur, days, url);
  MailApp.sendEmail({ to: user.email, subject: email.subject, htmlBody: email.html, name: 'Bill Tracker' });
  return items.length > 0;
}

/* ---------- Reminder email design (inline styles so it looks right in Gmail, Apple Mail, Outlook) ---------- */

function reminderEmail_(user, items, cur, days, url) {
  const DOW = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const nice = d => { const t = new Date(Date.parse(d)); return DOW[t.getUTCDay()] + ', ' + MON[t.getUTCMonth()] + ' ' + t.getUTCDate(); };
  const money = n => cur + Number(n).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const font = "font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;";
  const overdue = items.filter(i => i.diff < 0);
  const total = items.reduce((t, i) => t + i.b.amount, 0);

  let subject, headline;
  if (!items.length) { subject = '🧾 Bill Tracker test reminder'; headline = 'All clear 🎉'; }
  else if (items.length === 1) {
    const i = items[0];
    subject = i.diff < 0 ? '⚠️ ' + i.b.name + ' is overdue' : i.diff === 0 ? '🧾 ' + i.b.name + ' is due today' : '🧾 ' + i.b.name + ' is due ' + nice(i.due);
    headline = i.diff < 0 ? '1 overdue bill' : i.diff === 0 ? '1 bill due today' : '1 bill due soon';
  } else {
    subject = (overdue.length ? '⚠️ ' + overdue.length + ' overdue · ' : '🧾 ') + items.length + ' bills due soon · ' + money(total);
    headline = items.length + ' bills need attention';
  }

  const pill = i => {
    const [bg, fg, txt] = i.diff < 0 ? ['#fdecec', '#dc2626', 'Overdue ' + (-i.diff) + (i.diff === -1 ? ' day' : ' days')]
      : i.diff === 0 ? ['#fff4e0', '#c2410c', 'Due today']
      : i.diff === 1 ? ['#fff4e0', '#c2410c', 'Due tomorrow']
      : ['#eef0ff', '#4f46e5', 'In ' + i.diff + ' days'];
    return '<span style="display:inline-block;background:' + bg + ';color:' + fg + ';font-size:12px;font-weight:700;padding:4px 10px;border-radius:99px;">' + txt + '</span>';
  };
  const tag = t => '<span style="display:inline-block;background:#f1f3f8;color:#64748b;font-size:11px;font-weight:600;padding:2px 8px;border-radius:99px;margin:6px 4px 0 0;">' + t + '</span>';

  const rows = items.map((i, n) =>
    '<tr><td style="padding:16px 0;' + (n ? 'border-top:1px solid #eceef3;' : '') + '">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' +
        '<td style="' + font + 'vertical-align:top;">' +
          '<div style="font-size:16px;font-weight:700;color:#0f172a;">' + esc_(i.b.name) + '</div>' +
          '<div style="font-size:13px;color:#64748b;margin-top:3px;">' + (i.diff < 0 ? 'Was due ' : 'Due ') + nice(i.due) + '</div>' +
          '<div>' + (i.b.owner !== 'shared' ? tag('Yours') : '') + (i.b.autopay ? tag('Autopay') : '') + (i.b.variable ? tag('Estimate') : '') + '</div>' +
        '</td>' +
        '<td style="' + font + 'vertical-align:top;text-align:right;white-space:nowrap;">' +
          '<div style="font-size:17px;font-weight:800;color:#0f172a;">' + money(i.b.amount) + '</div>' +
          '<div style="margin-top:6px;">' + pill(i) + '</div>' +
        '</td>' +
      '</tr></table>' +
    '</td></tr>').join('');

  const body = items.length
    ? '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' + rows + '</table>'
    : '<p style="' + font + 'font-size:15px;color:#334155;margin:8px 0 0;">Nothing is due in the next ' + days + ' day' + (days === 1 ? '' : 's') + '. This was just a test.</p>';

  const html =
    '<!DOCTYPE html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="color-scheme" content="light"></head>' +
    '<body style="margin:0;padding:0;background:#f3f4f8;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f3f4f8;"><tr><td align="center" style="padding:24px 12px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:520px;">' +
      // header
      '<tr><td style="background:#6366f1;background-image:linear-gradient(135deg,#6366f1,#8b5cf6 60%,#ec4899);border-radius:20px 20px 0 0;padding:26px 24px;' + font + '">' +
        '<div style="font-size:13px;color:#e0e7ff;font-weight:600;">🧾 Bill Tracker</div>' +
        '<div style="font-size:24px;color:#ffffff;font-weight:800;margin-top:8px;">' + headline + '</div>' +
        (items.length ? '<div style="font-size:15px;color:#eef2ff;margin-top:4px;">' + money(total) + ' total' + (overdue.length ? ' · ' + overdue.length + ' overdue' : '') + '</div>' : '') +
      '</td></tr>' +
      // card
      '<tr><td style="background:#ffffff;border-radius:0 0 20px 20px;padding:8px 24px 24px;' + font + '">' +
        '<p style="font-size:15px;color:#334155;margin:16px 0 4px;">Hi ' + esc_(user.name) + ',' + (items.length ? ' here’s what’s coming up:' : '') + '</p>' +
        body +
        (url ? '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:18px;"><tr><td align="center" style="border-radius:14px;background:#6366f1;">' +
          '<a href="' + url + '" style="display:block;padding:14px 20px;' + font + 'font-size:15px;font-weight:700;color:#ffffff;text-decoration:none;">Open Bill Tracker</a></td></tr></table>' : '') +
      '</td></tr>' +
      '<tr><td style="padding:16px 8px;text-align:center;' + font + 'font-size:12px;color:#94a3b8;">You’re getting this because email reminders are on. Turn them off in Settings → Email reminders.</td></tr>' +
    '</table></td></tr></table></body></html>';

  return { subject: subject, html: html };
}

function configureTrigger_(on, hour) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendReminders')
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (on) ScriptApp.newTrigger('sendReminders').timeBased().everyDays(1).atHour(hour).create();
}

/* ======================================================================
   People & logins
   Accounts live in Script Properties (USERS) with salted SHA-256 password hashes.
   ====================================================================== */

function login(username, password) {
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();
  const lockUntil = Number(props.getProperty('LOCK_UNTIL') || 0);
  if (now < lockUntil) throw new Error('Too many wrong tries. Try again in ' + Math.ceil((lockUntil - now) / 60000) + ' min.');

  const uname = norm_(username);
  let users = users_();
  let user = null;

  if (!users.length) {
    // First ever login: create the admin account from Script Properties
    const pu = props.getProperty('APP_USERNAME'), pp = props.getProperty('APP_PASSWORD');
    if (!pu || !pp) throw new Error('No login set yet. In the script editor, add APP_USERNAME and APP_PASSWORD under Project Settings → Script Properties.');
    if (uname === norm_(pu) && String(password) === pp) {
      const old = getSettings_();
      user = makeUser_(pu.trim(), pu.trim(), password, true);
      user.email = Session.getEffectiveUser().getEmail();
      user.income = num_(old.income);
      user.emailReminders = bool_(old.emailReminders);
      saveUsers_([user]);
      props.deleteProperty('APP_PASSWORD');
    }
  } else {
    const u = users.find(x => norm_(x.username) === uname);
    if (u && hash_(u.salt + String(password)) === u.hash) user = u;
  }

  if (!user) {
    const fails = Number(props.getProperty('FAILS') || 0) + 1;
    if (fails >= 5) { props.setProperty('LOCK_UNTIL', String(now + 15 * 60000)); props.setProperty('FAILS', '0'); }
    else props.setProperty('FAILS', String(fails));
    Utilities.sleep(1000);
    throw new Error(fails >= 5 ? 'Too many wrong tries — locked for 15 minutes.' : 'Wrong username or password.');
  }
  props.setProperty('FAILS', '0');
  return newSession_(user.id);
}

function logout(token) {
  const h = hash_(String(token || ''));
  saveSessions_(sessions_().filter(x => x.h !== h));
  return true;
}

/** Change your own username / password. Logs out your other devices; returns a fresh token. */
function changeLogin_(user, currentPassword, newUsername, newPassword) {
  if (hash_(user.salt + String(currentPassword || '')) !== user.hash) throw new Error('Current password is wrong.');
  newUsername = String(newUsername || '').trim() || user.username;
  if (String(newPassword || '').length < 6) throw new Error('New password must be at least 6 characters.');
  if (users_().some(u => u.id !== user.id && norm_(u.username) === norm_(newUsername))) throw new Error('That username is taken.');
  updateUser_(user.id, u => { u.username = newUsername; setPassword_(u, newPassword); });
  saveSessions_(sessions_().filter(x => x.uid !== user.id));
  return newSession_(user.id);
}

/** Admin: add a person (e.g. your wife). */
function addPerson_(user, name, username, password, email) {
  return withLock_(() => {
    if (!user.admin) throw new Error('Only the admin can add people.');
    name = String(name || '').trim(); username = String(username || '').trim();
    if (!name || !username) throw new Error('Name and username are required.');
    if (String(password || '').length < 6) throw new Error('Password must be at least 6 characters.');
    const users = users_();
    if (users.some(u => norm_(u.username) === norm_(username))) throw new Error('That username is taken.');
    const p = makeUser_(name, username, password, false);
    p.email = String(email || '').trim();
    users.push(p);
    saveUsers_(users);
    return getData_(user);
  });
}

/** Admin: remove a person's login. Their personal bills stay in the sheet but are hidden. */
function removePerson_(user, id) {
  return withLock_(() => {
    if (!user.admin) throw new Error('Only the admin can remove people.');
    if (id === user.id) throw new Error("You can't remove yourself.");
    saveUsers_(users_().filter(u => u.id !== id));
    saveSessions_(sessions_().filter(x => x.uid !== id));
    return getData_(user);
  });
}

/** Admin: set a new password for someone who forgot theirs. */
function resetPersonPassword_(user, id, password) {
  if (!user.admin) throw new Error('Only the admin can reset passwords.');
  if (String(password || '').length < 6) throw new Error('Password must be at least 6 characters.');
  updateUser_(id, u => setPassword_(u, password));
  saveSessions_(sessions_().filter(x => x.uid !== id));
  return getData_(user);
}

/** Forgot the admin password? Add APP_PASSWORD in Script Properties, then run this from the editor. */
function resetLogin() {
  if (Session.getActiveUser().getEmail() !== Session.getEffectiveUser().getEmail()) throw new Error('Run this from the script editor.');
  const props = PropertiesService.getScriptProperties();
  const pp = props.getProperty('APP_PASSWORD');
  if (!pp) throw new Error('First add APP_PASSWORD (your new password) in Project Settings → Script Properties.');
  const users = users_();
  const admin = users.find(u => u.admin);
  if (!admin) { props.deleteProperty('USERS'); Logger.log('No accounts yet — just log in with APP_USERNAME / APP_PASSWORD.'); return; }
  setPassword_(admin, pp);
  saveUsers_(users);
  ['SESSIONS', 'FAILS', 'LOCK_UNTIL', 'APP_PASSWORD'].forEach(k => props.deleteProperty(k));
  Logger.log('Password reset for "' + admin.username + '". Log in with the password you set.');
}

/* ---------- account helpers ---------- */

function users_() {
  const props = PropertiesService.getScriptProperties();
  let list = [];
  try { list = JSON.parse(props.getProperty('USERS') || '[]'); } catch (e) { list = []; }
  // Upgrade from the single-login version
  if (!list.length && props.getProperty('PASS_HASH')) {
    const old = getSettings_();
    const name = props.getProperty('USERNAME') || props.getProperty('APP_USERNAME') || 'me';
    list = [{ id: Utilities.getUuid(), name: name, username: name, salt: props.getProperty('PASS_SALT'),
      hash: props.getProperty('PASS_HASH'), admin: true, email: Session.getEffectiveUser().getEmail(),
      income: num_(old.income), emailReminders: bool_(old.emailReminders) }];
    saveUsers_(list);
    ['USERNAME', 'PASS_SALT', 'PASS_HASH', 'SESSIONS'].forEach(k => props.deleteProperty(k));
  }
  return list;
}

function saveUsers_(list) { PropertiesService.getScriptProperties().setProperty('USERS', JSON.stringify(list)); }

function updateUser_(id, fn) {
  const users = users_();
  const u = users.find(x => x.id === id);
  if (!u) throw new Error('Person not found');
  fn(u);
  saveUsers_(users);
}

function makeUser_(name, username, password, admin) {
  const u = { id: Utilities.getUuid(), name: name, username: username, admin: !!admin, email: '', income: 0, emailReminders: false };
  setPassword_(u, password);
  return u;
}

function setPassword_(u, password) {
  u.salt = Utilities.getUuid();
  u.hash = hash_(u.salt + String(password));
}

function adminId_() {
  const a = users_().find(u => u.admin);
  return a ? a.id : '';
}

function checkUserById_(id) { return users_().find(u => u.id === id); }

function newSession_(uid) {
  const token = Utilities.getUuid() + Utilities.getUuid();
  const now = Date.now();
  const list = sessions_().filter(x => x.exp > now);
  list.push({ h: hash_(token), exp: now + SESSION_DAYS * 864e5, uid: uid });
  saveSessions_(list.slice(-30));
  return token;
}

/** Returns the logged-in person, or null. */
function checkToken_(token) {
  if (!token) return null;
  const now = Date.now(), h = hash_(String(token));
  const list = sessions_();
  const sess = list.find(x => x.h === h && x.exp > now);
  if (!sess) return null;
  const user = checkUserById_(sess.uid);
  if (!user) return null;
  if (sess.exp - now < (SESSION_DAYS - 7) * 864e5) {   // keep devices you use logged in
    sess.exp = now + SESSION_DAYS * 864e5;
    saveSessions_(list.filter(x => x.exp > now));
  }
  return user;
}

function sessions_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('SESSIONS') || '[]'); } catch (e) { return []; }
}
function saveSessions_(list) { PropertiesService.getScriptProperties().setProperty('SESSIONS', JSON.stringify(list)); }

function hash_(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
function norm_(s) { return String(s || '').trim().toLowerCase(); }

/* ======================================================================
   Sheet helpers
   ====================================================================== */

function ss_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SPREADSHEET_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { /* fall through */ } }
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { /* standalone */ }
  if (!ss) ss = SpreadsheetApp.create('Bill Tracker Data');
  props.setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

function sheet_(name, cols) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else {
    const lastCol = Math.max(sh.getLastColumn(), 1);
    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    cols.forEach(c => {
      if (header.indexOf(c) === -1) {
        const col = sh.getLastColumn() + 1;
        sh.getRange(1, col, sh.getMaxRows(), 1).setNumberFormat('@');
        sh.getRange(1, col).setValue(c).setFontWeight('bold');
      }
    });
  }
  return sh;
}

function read_(name, cols) {
  const sh = sheet_(name, cols);
  const values = sh.getDataRange().getValues();
  const header = values.shift();
  const tz = Session.getScriptTimeZone();
  return values.map((r, i) => {
    const o = { _row: i + 2 };
    header.forEach((h, j) => {
      const v = r[j];
      o[h] = v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v);
    });
    return o;
  }).filter(o => o.id);
}

function writeRow_(sh, rowNum, obj) {
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = header.map(h => (obj[h] === undefined || obj[h] === null) ? '' : String(obj[h]));
  const r = rowNum || sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, header.length).setNumberFormat('@').setValues([row]);
}

function deleteWhere_(name, cols, fn) {
  const sh = sheet_(name, cols);
  read_(name, cols).filter(fn).map(r => r._row).sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
}

function getSettings_() {
  const sh = sheet_(SETTINGS, SET_COLS);
  const s = Object.assign({}, DEFAULT_SETTINGS);
  sh.getDataRange().getValues().slice(1).forEach(r => { if (r[0]) s[r[0]] = String(r[1]); });
  return s;
}

/** Remember a category so it can be picked again later. */
function addCategory_(cat) {
  cat = String(cat || '').trim();
  if (!cat) return;
  let list = [];
  try { list = JSON.parse(getSettings_().categories || '[]'); } catch (e) { list = []; }
  if (list.some(c => c.toLowerCase() === cat.toLowerCase())) return;
  list.push(cat);
  setSetting_('categories', JSON.stringify(list));
}

function setSetting_(key, value) {
  const sh = sheet_(SETTINGS, SET_COLS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) { sh.getRange(i + 1, 2).setNumberFormat('@').setValue(String(value)); return; }
  }
  sh.getRange(sh.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([[key, String(value)]]);
}

function adjustBalance_(billId, delta) {
  if (!delta) return;
  const b = read_(BILLS, BILL_COLS).find(x => x.id === billId);
  if (!b || !bool_(b.trackBalance)) return;
  b.balance = round_(num_(b.balance) + delta);
  writeRow_(sheet_(BILLS, BILL_COLS), b._row, b);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ======================================================================
   Data shaping & small utils
   ====================================================================== */

function billOut_(o) {
  return {
    id: o.id, name: o.name, amount: num_(o.amount), dueDay: parseInt(o.dueDay, 10) || 1,
    category: o.category || '', frequency: o.frequency || 'monthly',
    startMonth: o.startMonth || '', endMonth: o.endMonth || '',
    variable: bool_(o.variable), autopay: bool_(o.autopay), trackBalance: bool_(o.trackBalance),
    balance: num_(o.balance), originalBalance: num_(o.originalBalance), minPayment: num_(o.minPayment),
    interestRate: num_(o.interestRate), url: o.url || '', notes: o.notes || '',
    active: o.active === '' ? true : bool_(o.active), owner: o.owner || 'shared',
    startDate: o.startDate || '', every: parseInt(o.every, 10) || 1, unit: o.unit || 'weeks', createdAt: o.createdAt || ''
  };
}

function payOut_(o) {
  return {
    id: o.id, billId: o.billId, month: o.month, amount: num_(o.amount), paidDate: o.paidDate,
    confirmation: o.confirmation || '', note: o.note || '', status: o.status || 'paid', paidBy: o.paidBy || ''
  };
}

function pcOut_(o) {
  let a = [];
  try { a = JSON.parse(o.allocations || '[]'); } catch (e) { a = []; }
  return { id: o.id, payDate: o.payDate, amount: num_(o.amount), note: o.note || '', allocations: a, owner: pcOwner_(o) };
}

function num_(v) { const n = parseFloat(v); return isNaN(n) ? 0 : n; }
function bool_(v) { return v === true || String(v).toLowerCase() === 'true'; }
function round_(n) { return Math.round(n * 100) / 100; }
function today_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function esc_(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function monthIndex_(ym) { const p = ym.split('-').map(Number); return p[0] * 12 + (p[1] - 1); }
function addMonths_(ym, n) { const i = monthIndex_(ym) + n; return Math.floor(i / 12) + '-' + String(i % 12 + 1).padStart(2, '0'); }
function daysBetween_(a, b) { return Math.round((Date.parse(b) - Date.parse(a)) / 86400000); }

/** How a bill repeats: by months (keyed 'yyyy-MM') or by days (keyed by each due date 'yyyy-MM-dd'). */
function sched_(b) {
  const f = b.frequency, n = Math.max(1, parseInt(b.every, 10) || 1);
  if (f === 'weekly') return { mode: 'day', step: 7 };
  if (f === 'biweekly') return { mode: 'day', step: 14 };
  if (f === 'custom') return b.unit === 'months' ? { mode: 'month', step: n } : { mode: 'day', step: b.unit === 'days' ? n : n * 7 };
  return { mode: 'month', step: ({ quarterly: 3, semiannual: 6, yearly: 12 })[f] || 1, once: f === 'once' };
}

function isDueIn_(b, ym) {
  const start = b.startMonth || ym;
  const idx = monthIndex_(ym), s = monthIndex_(start);
  if (idx < s) return false;
  if (b.endMonth && idx > monthIndex_(b.endMonth)) return false;
  const sc = sched_(b), diff = idx - s;
  return sc.once ? diff === 0 : diff % sc.step === 0;
}

/** Every time a bill is due in a month: [{key, due}] */
function occurrences_(b, ym) {
  const sc = sched_(b);
  if (sc.mode === 'month') return isDueIn_(b, ym) ? [{ key: ym, due: dueDate_(b, ym) }] : [];
  const anchor = b.startDate || ((b.startMonth || ym) + '-' + String(Math.min(b.dueDay || 1, 28)).padStart(2, '0'));
  const first = ym + '-01', last = addDays_(addMonths_(ym, 1) + '-01', -1);
  if (last < anchor) return [];
  let d = addDays_(anchor, Math.max(0, Math.ceil(daysBetween_(anchor, first) / sc.step)) * sc.step);
  const out = [];
  while (d <= last) {
    if (b.endMonth && d.slice(0, 7) > b.endMonth) break;
    out.push({ key: d, due: d });
    d = addDays_(d, sc.step);
  }
  return out;
}

function addDays_(d, n) { return new Date(Date.parse(d) + n * 86400000).toISOString().slice(0, 10); }

function dueDate_(b, ym) {
  const p = ym.split('-').map(Number);
  const dim = new Date(p[0], p[1], 0).getDate();
  return ym + '-' + String(Math.min(b.dueDay, dim)).padStart(2, '0');
}
