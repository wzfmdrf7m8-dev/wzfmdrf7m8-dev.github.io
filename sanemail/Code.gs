/**
 * SaneMail — a self-hosted, SaneBox-style inbox filter for Gmail / Google Workspace.
 *
 * Runs entirely inside your own Google account as a Google Apps Script.
 * Nothing leaves your account; there is no third-party server.
 *
 * What it does (every 10 minutes):
 *   • Keeps mail from people you actually correspond with in the Inbox.
 *   • Files everything else into @SaneLater (unknown senders) or @SaneNews (newsletters / bulk).
 *   • @SaneBlackHole: drop any thread here and that sender is trashed forever after.
 *   • Training: move a thread into/out of @SaneLater / @SaneNews and the sender is remembered.
 *   • Snooze: label a thread @SaneTomorrow or @SaneNextWeek and it disappears, then comes back unread.
 *   • @SaneNoReplies: tracks emails you sent that nobody has answered after N days.
 *   • Daily digest email of what was filed, with one-click "→ Inbox" / "→ Black hole" links.
 *   • A dashboard web page (deploy as a web app) with stats and your sender rules.
 *
 * Install: see README.md. Short version: paste this file into script.google.com,
 * enable nothing extra, run setup() once and approve the permissions.
 */

// ───────────────────────────── Configuration ─────────────────────────────

const CONFIG = {
  LABELS: {
    later: '@SaneLater',
    news: '@SaneNews',
    blackhole: '@SaneBlackHole',
    tomorrow: '@SaneTomorrow',
    nextWeek: '@SaneNextWeek',
    noReplies: '@SaneNoReplies',
  },

  // Domains whose mail always stays in the Inbox (your own domain is added automatically).
  VIP_DOMAINS: [],

  // Exact addresses that always stay in the Inbox (bank alerts, school, etc.).
  VIP_SENDERS: [],

  // Set true to put newsletters in @SaneLater instead of a separate @SaneNews folder.
  NEWS_INTO_LATER: false,

  // Flag sent mail with no reply after this many days.
  NO_REPLY_DAYS: 3,
  NO_REPLY_LOOKBACK_DAYS: 14,

  // Hour (0–23, script time zone) the daily digest is sent and snoozed mail wakes up.
  DIGEST_HOUR: 7,
  WAKE_HOUR: 7,

  // How far back setup() learns your contacts from Sent mail, and how much of the
  // existing inbox the very first run sorts.
  LEARN_FROM_SENT_MAX_THREADS: 1500,
  INITIAL_LOOKBACK_DAYS: 2,

  MAX_THREADS_PER_RUN: 150,

  // true = log what would happen without moving any mail. Good for a first test.
  DRY_RUN: false,
};

// Hidden marker labels so the script can tell its own filing apart from your training.
const MARKER = { later: '_sane/later', news: '_sane/news' };
const SHEET_NAME = 'SaneMail data';

// ───────────────────────────── Setup ─────────────────────────────

/** Run once. Creates labels, the rules spreadsheet, learns your contacts, installs triggers. */
function setup() {
  Object.values(CONFIG.LABELS).forEach(getLabel_);
  Object.values(MARKER).forEach(getLabel_);
  getSpreadsheet_();

  const props = PropertiesService.getUserProperties();
  if (!props.getProperty('lastRun')) {
    props.setProperty('lastRun', String(Date.now() - CONFIG.INITIAL_LOOKBACK_DAYS * 864e5));
  }

  learnFromSent();

  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('run').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('sendDigest').timeBased().everyDays(1).atHour(CONFIG.DIGEST_HOUR).create();

  Logger.log('SaneMail installed. Rules sheet: ' + getSpreadsheet_().getUrl());
  run();
}

/** Remove triggers (labels and your rules are kept). */
function uninstall() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('SaneMail triggers removed. Labels and the rules sheet were left in place.');
}

/** Scan Sent mail and mark everyone you've written to as "Inbox". */
function learnFromSent() {
  const store = loadStore_();
  let n = 0;
  for (let start = 0; start < CONFIG.LEARN_FROM_SENT_MAX_THREADS; start += 100) {
    const threads = GmailApp.search('in:sent', start, 100);
    if (!threads.length) break;
    threads.forEach(t => t.getMessages().forEach(m => {
      if (!isMe_(m.getFrom())) return;
      recipients_(m).forEach(addr => { if (store.set(addr, 'inbox', 'sent')) n++; });
    }));
  }
  store.save();
  Logger.log('Learned ' + n + ' correspondents from Sent mail.');
}

// ───────────────────────────── Main loop ─────────────────────────────

function run() {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(5000)) return;
  try {
    const props = PropertiesService.getUserProperties();
    const lastRun = Number(props.getProperty('lastRun')) || Date.now() - 864e5;
    const startedAt = Date.now();
    const store = loadStore_();
    const stats = loadStats_();

    learnNewSent_(store, lastRun);
    processTraining_(store, lastRun);
    processSnoozes_();
    processInbox_(store, stats, lastRun);
    maybeProcessNoReplies_();

    store.save();
    saveStats_(stats);
    props.setProperty('lastRun', String(startedAt));
  } finally {
    lock.releaseLock();
  }
}

function learnNewSent_(store, lastRun) {
  GmailApp.search('in:sent after:' + epoch_(lastRun - 3600e3), 0, 100).forEach(t =>
    t.getMessages().forEach(m => {
      if (m.getDate().getTime() < lastRun - 3600e3 || !isMe_(m.getFrom())) return;
      recipients_(m).forEach(addr => store.set(addr, 'inbox', 'sent'));
    }));
}

/** Sort new mail in the Inbox. */
function processInbox_(store, stats, lastRun) {
  const threads = GmailApp.search('in:inbox after:' + epoch_(lastRun - 120e3), 0, CONFIG.MAX_THREADS_PER_RUN);
  threads.forEach(thread => {
    if (thread.hasStarredMessages()) return;
    const msgs = thread.getMessages();
    const newest = msgs[msgs.length - 1];
    if (isMe_(newest.getFrom())) return;

    const sender = addr_(newest.getFrom());
    let folder = classify_(sender, newest, msgs, store);
    if (folder === 'news' && CONFIG.NEWS_INTO_LATER) folder = 'later';
    stats.bump(folder);
    file_(thread, folder, sender);
  });
}

/** Decide where a message belongs. Returns 'inbox' | 'later' | 'news' | 'blackhole'. */
function classify_(sender, msg, msgs, store) {
  // 1. Any conversation you've taken part in stays in the Inbox.
  if (msgs.some(m => isMe_(m.getFrom()))) { store.set(sender, 'inbox', 'replied'); return 'inbox'; }

  // 2. Explicit rule (your training, or something learned earlier).
  const rule = store.get(sender);
  if (rule && rule.source !== 'auto') return rule.folder;

  // 3. VIPs and your own domain.
  const domain = sender.split('@')[1] || '';
  if (CONFIG.VIP_SENDERS.map(s => s.toLowerCase()).includes(sender)) return 'inbox';
  if (vipDomains_().some(d => domain === d || domain.endsWith('.' + d))) return 'inbox';

  // 4. Have you ever emailed this person? (Cached after the first check.)
  if (!rule && GmailApp.search('in:sent to:' + sender, 0, 1).length) {
    store.set(sender, 'inbox', 'sent');
    return 'inbox';
  }
  if (rule) return rule.folder;

  // 5. Bulk / newsletter signals.
  const folder = looksBulk_(sender, msg) ? 'news' : 'later';
  store.set(sender, folder, 'auto');
  return folder;
}

function looksBulk_(sender, msg) {
  if (/^(no-?reply|do-?not-?reply|newsletter|news|marketing|notifications?|updates?|info|hello|team|mailer|digest|offers?|deals?|promo)[@+._-]/i.test(sender)) return true;
  if (msg.getHeader('List-Unsubscribe') || msg.getHeader('List-Id')) return true;
  if (/bulk|list|junk/i.test(msg.getHeader('Precedence') || '')) return true;
  if (/mailchimp|sendgrid|mailgun|constantcontact|hubspot|klaviyo|substack|campaign/i.test(msg.getHeader('X-Mailer') || '')) return true;
  return false;
}

/** Move a thread to its folder. */
function file_(thread, folder, sender) {
  const L = CONFIG.LABELS;
  Logger.log((CONFIG.DRY_RUN ? '[dry run] ' : '') + folder.padEnd(9) + ' ← ' + sender + ' — ' + thread.getFirstMessageSubject());
  if (CONFIG.DRY_RUN) return;

  ['later', 'news'].forEach(f => {
    if (f !== folder) { thread.removeLabel(getLabel_(L[f])); thread.removeLabel(getLabel_(MARKER[f])); }
  });
  if (folder === 'blackhole') { thread.moveToTrash(); return; }
  if (folder === 'inbox') { thread.moveToInbox(); return; }
  thread.addLabel(getLabel_(L[folder]));
  thread.addLabel(getLabel_(MARKER[folder]));
  thread.moveToArchive();
}

// ───────────────────────────── Training ─────────────────────────────

/**
 * Learns from what you do in Gmail:
 *   • Put a thread in @SaneLater / @SaneNews  → that sender always goes there.
 *   • Move a thread out of those to the Inbox → that sender always stays in the Inbox.
 *   • Put a thread in @SaneBlackHole          → sender is trashed from now on.
 */
function processTraining_(store, lastRun) {
  const L = CONFIG.LABELS;

  // Black hole.
  threadsWith_(L.blackhole).forEach(t => {
    const s = senderOf_(t);
    if (s) store.set(s, 'blackhole', 'trained');
    if (!CONFIG.DRY_RUN) { t.removeLabel(getLabel_(L.blackhole)); t.moveToTrash(); }
  });

  ['later', 'news'].forEach(f => {
    const label = L[f], marker = MARKER[f];

    // You put it in the folder yourself.
    threadsWith_(label).filter(t => !hasLabel_(t, marker)).forEach(t => {
      const s = senderOf_(t);
      if (s) store.set(s, f, 'trained');
      if (!CONFIG.DRY_RUN) file_(t, f, s || '?');
    });

    // You took it out of the folder (Gmail's "Move to Inbox" removes the label).
    threadsWith_(marker).filter(t => !hasLabel_(t, label)).forEach(t => {
      if (t.isInInbox()) { const s = senderOf_(t); if (s) store.set(s, 'inbox', 'trained'); }
      if (!CONFIG.DRY_RUN) t.removeLabel(getLabel_(marker));
    });

    // Still labelled but back in the Inbox with no new mail → you moved it back.
    threadsWith_(label).filter(t => hasLabel_(t, marker) && t.isInInbox()).forEach(t => {
      if (t.getLastMessageDate().getTime() > lastRun - 120e3) return; // new reply; processInbox_ handles it
      const s = senderOf_(t);
      if (s) store.set(s, 'inbox', 'trained');
      if (!CONFIG.DRY_RUN) { t.removeLabel(getLabel_(label)); t.removeLabel(getLabel_(marker)); }
    });
  });
}

/** Train a single thread's sender (used by digest links and the dashboard). */
function trainThread_(threadId, folder) {
  const t = GmailApp.getThreadById(threadId);
  if (!t) return null;
  const s = senderOf_(t);
  const store = loadStore_();
  if (s) store.set(s, folder, 'trained');
  store.save();
  file_(t, folder, s || '?');
  return s;
}

// ───────────────────────────── Snooze ─────────────────────────────

function processSnoozes_() {
  const L = CONFIG.LABELS;
  const props = PropertiesService.getUserProperties();
  const snoozes = JSON.parse(props.getProperty('snoozes') || '{}');
  const now = Date.now();

  [[L.tomorrow, wakeTomorrow_], [L.nextWeek, wakeNextMonday_]].forEach(([name, when]) => {
    const label = getLabel_(name);
    const seen = {};
    threadsWith_(name).forEach(t => {
      const id = t.getId();
      seen[id] = true;
      if (!snoozes[id]) {
        snoozes[id] = { wake: when().getTime(), label: name };
        if (!CONFIG.DRY_RUN) t.moveToArchive();
      } else if (snoozes[id].wake <= now) {
        if (!CONFIG.DRY_RUN) { t.removeLabel(label); t.moveToInbox(); t.markUnread(); }
        delete snoozes[id];
      }
    });
    // Label removed by hand → forget the snooze.
    Object.keys(snoozes).forEach(id => { if (snoozes[id].label === name && !seen[id]) delete snoozes[id]; });
  });

  props.setProperty('snoozes', JSON.stringify(snoozes));
}

function wakeTomorrow_() {
  const d = new Date(); d.setDate(d.getDate() + 1); d.setHours(CONFIG.WAKE_HOUR, 0, 0, 0); return d;
}
function wakeNextMonday_() {
  const d = new Date(); d.setDate(d.getDate() + ((8 - d.getDay()) % 7 || 7)); d.setHours(CONFIG.WAKE_HOUR, 0, 0, 0); return d;
}

// ───────────────────────────── No replies ─────────────────────────────

function maybeProcessNoReplies_() {
  const props = PropertiesService.getUserProperties();
  if (Date.now() - Number(props.getProperty('noRepliesAt') || 0) < 3600e3) return;
  props.setProperty('noRepliesAt', String(Date.now()));
  if (CONFIG.DRY_RUN) return;

  const label = getLabel_(CONFIG.LABELS.noReplies);

  // Clear the flag once someone answers.
  threadsWith_(CONFIG.LABELS.noReplies).forEach(t => {
    const m = t.getMessages();
    if (!isMe_(m[m.length - 1].getFrom())) t.removeLabel(label);
  });

  // Flag sent threads still waiting on a reply.
  const query = 'in:sent older_than:' + CONFIG.NO_REPLY_DAYS + 'd newer_than:' + CONFIG.NO_REPLY_LOOKBACK_DAYS + 'd';
  GmailApp.search(query, 0, 100).forEach(t => {
    if (hasLabel_(t, CONFIG.LABELS.noReplies)) return;
    const m = t.getMessages();
    const last = m[m.length - 1];
    if (!isMe_(last.getFrom())) return;
    const others = recipients_(last).filter(a => !isMe_(a));
    if (others.length) t.addLabel(label);
  });
}

// ───────────────────────────── Digest ─────────────────────────────

function sendDigest() {
  const L = CONFIG.LABELS;
  const url = webAppUrl_();
  const section = (title, name) => {
    const threads = threadsWith_(name).filter(t => t.getLastMessageDate().getTime() > Date.now() - 864e5);
    if (!threads.length) return '';
    const rows = threads.map(t => {
      const m = t.getMessages()[t.getMessageCount() - 1];
      const id = t.getId();
      const actions = url
        ? ' <a href="' + url + '?a=inbox&t=' + id + '">→ Inbox</a> · <a href="' + url + '?a=blackhole&t=' + id + '" style="color:#b00">Black hole</a>'
        : '';
      return '<tr><td style="padding:6px 8px;border-bottom:1px solid #eee"><b>' + esc_(nameOf_(m.getFrom())) + '</b><br>' +
        '<a href="https://mail.google.com/mail/u/0/#all/' + id + '">' + esc_(t.getFirstMessageSubject() || '(no subject)') + '</a>' +
        '<div style="color:#777;font-size:12px">' + esc_(m.getPlainBody().replace(/\s+/g, ' ').slice(0, 140)) + '</div></td>' +
        '<td style="padding:6px 8px;border-bottom:1px solid #eee;white-space:nowrap;font-size:12px">' + actions + '</td></tr>';
    }).join('');
    return '<h3 style="margin:20px 0 6px">' + title + ' (' + threads.length + ')</h3><table style="border-collapse:collapse;width:100%">' + rows + '</table>';
  };

  const body = section('📥 SaneLater', L.later) + section('📰 SaneNews', L.news);
  if (!body) return;
  const noReply = threadsWith_(L.noReplies).length;
  GmailApp.sendEmail(me_(), 'SaneMail digest — ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'EEE d MMM'), '', {
    htmlBody: '<div style="font-family:system-ui,Arial,sans-serif;max-width:720px">' + body +
      (noReply ? '<p style="margin-top:20px">⏳ ' + noReply + ' sent email(s) still waiting on a reply in ' + L.noReplies + '.</p>' : '') +
      (url ? '<p style="color:#999;font-size:12px">Dashboard: <a href="' + url + '">' + url + '</a></p>' : '') + '</div>',
    name: 'SaneMail',
  });
}

// ───────────────────────────── Web app (dashboard + digest actions) ─────────────────────────────
// Deploy → New deployment → Web app → Execute as: Me, Who has access: Only myself.

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.a && p.t && ['inbox', 'later', 'news', 'blackhole'].includes(p.a)) {
    const s = trainThread_(p.t, p.a);
    return page_('<p>' + (s ? 'Got it — mail from <b>' + esc_(s) + '</b> now goes to <b>' + p.a + '</b>.' : 'Thread not found.') +
      '</p><p><a href="' + webAppUrl_() + '">Back to dashboard</a></p>');
  }
  if (p.a === 'rule' && p.s && ['inbox', 'later', 'news', 'blackhole'].includes(p.f)) {
    const store = loadStore_(); store.set(p.s.toLowerCase(), p.f, 'trained'); store.save();
  }
  return page_(dashboard_());
}

function dashboard_() {
  const stats = loadStats_().data;
  const days = Object.keys(stats).sort().slice(-14);
  const tot = { inbox: 0, later: 0, news: 0, blackhole: 0 };
  days.forEach(d => Object.keys(tot).forEach(k => tot[k] += stats[d][k] || 0));
  const all = tot.inbox + tot.later + tot.news + tot.blackhole || 1;
  const max = Math.max(1, ...days.map(d => Object.values(stats[d]).reduce((a, b) => a + b, 0)));
  const colors = { inbox: '#2563eb', later: '#f59e0b', news: '#10b981', blackhole: '#111' };

  const tiles = Object.keys(tot).map(k =>
    '<div class="tile"><div class="n" style="color:' + colors[k] + '">' + tot[k] + '</div><div>' + k + '</div></div>').join('');

  const bars = days.map(d => {
    const segs = Object.keys(tot).map(k => {
      const v = stats[d][k] || 0;
      return v ? '<div title="' + k + ': ' + v + '" style="height:' + (v / max * 140) + 'px;background:' + colors[k] + '"></div>' : '';
    }).reverse().join('');
    return '<div class="bar"><div class="stack">' + segs + '</div><small>' + d.slice(5) + '</small></div>';
  }).join('');

  const store = loadStore_();
  const rows = Object.entries(store.map)
    .filter(([, r]) => r.folder !== 'inbox' || r.source === 'trained')
    .sort((a, b) => String(b[1].updated).localeCompare(String(a[1].updated)))
    .slice(0, 200)
    .map(([s, r]) => '<tr><td>' + esc_(s) + '</td><td>' + r.folder + '</td><td>' + r.source + '</td><td>' +
      ['inbox', 'later', 'news', 'blackhole'].filter(f => f !== r.folder)
        .map(f => '<a href="?a=rule&s=' + encodeURIComponent(s) + '&f=' + f + '">' + f + '</a>').join(' · ') + '</td></tr>').join('');

  return '<h1>SaneMail</h1><p class="muted">' + esc_(me_()) + ' · last 14 days · ' +
    Math.round((tot.later + tot.news + tot.blackhole) / all * 100) + '% of mail kept out of your Inbox</p>' +
    '<div class="tiles">' + tiles + '</div>' +
    '<h2>Per day</h2><div class="chart">' + (bars || '<p class="muted">No data yet.</p>') + '</div>' +
    '<h2>Sender rules</h2><p class="muted">Full list (editable) in the <a href="' + getSpreadsheet_().getUrl() + '" target="_blank">rules sheet</a>.</p>' +
    '<table><tr><th>Sender</th><th>Folder</th><th>Learned from</th><th>Move to</th></tr>' + rows + '</table>';
}

function page_(html) {
  return HtmlService.createHtmlOutput(
    '<style>body{font-family:system-ui,Arial,sans-serif;max-width:900px;margin:24px auto;padding:0 16px;color:#222}' +
    '.muted{color:#777}.tiles{display:flex;gap:12px;flex-wrap:wrap}.tile{flex:1;min-width:120px;border:1px solid #eee;border-radius:10px;padding:14px}' +
    '.tile .n{font-size:30px;font-weight:700}.chart{display:flex;gap:6px;align-items:flex-end;height:170px;overflow-x:auto}' +
    '.bar{display:flex;flex-direction:column;align-items:center;width:34px}.stack{display:flex;flex-direction:column;width:22px;justify-content:flex-end;height:140px}' +
    'table{border-collapse:collapse;width:100%;font-size:14px}td,th{text-align:left;padding:6px;border-bottom:1px solid #eee}</style>' + html
  ).setTitle('SaneMail').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ───────────────────────────── Storage ─────────────────────────────

const SOURCE_RANK = { auto: 0, sent: 1, replied: 1, trained: 2 };

/** Sender rules live in a Google Sheet so you can read and edit them. */
function loadStore_() {
  const sheet = getSheet_();
  const values = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues() : [];
  const map = {};
  values.forEach(([email, folder, source, updated]) => {
    if (email) map[String(email).toLowerCase().trim()] = { folder: String(folder).trim(), source: String(source || 'trained').trim(), updated };
  });
  let dirty = false;
  return {
    map,
    get: s => map[s],
    /** Returns true if something changed. Lower-ranked sources never overwrite higher ones. */
    set(s, folder, source) {
      if (!s || isMe_(s)) return false;
      const cur = map[s];
      if (cur && (SOURCE_RANK[cur.source] || 0) > SOURCE_RANK[source]) return false;
      if (cur && cur.folder === folder && cur.source === source) return false;
      map[s] = { folder, source, updated: new Date().toISOString().slice(0, 16).replace('T', ' ') };
      dirty = true;
      return true;
    },
    save() {
      if (!dirty) return;
      const rows = Object.entries(map).sort().map(([e, r]) => [e, r.folder, r.source, r.updated]);
      sheet.getRange(2, 1, Math.max(sheet.getLastRow(), 2), 4).clearContent();
      if (rows.length) sheet.getRange(2, 1, rows.length, 4).setValues(rows);
      dirty = false;
    },
  };
}

function getSpreadsheet_() {
  const props = PropertiesService.getUserProperties();
  const id = props.getProperty('sheetId');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { /* recreate */ } }
  const ss = SpreadsheetApp.create('SaneMail rules');
  const sh = ss.getSheets()[0].setName(SHEET_NAME);
  sh.getRange(1, 1, 1, 4).setValues([['sender', 'folder (inbox/later/news/blackhole)', 'source', 'updated']]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 320).setColumnWidth(2, 240);
  props.setProperty('sheetId', ss.getId());
  return ss;
}
function getSheet_() { return getSpreadsheet_().getSheetByName(SHEET_NAME) || getSpreadsheet_().getSheets()[0]; }

function loadStats_() {
  const props = PropertiesService.getUserProperties();
  const data = JSON.parse(props.getProperty('stats') || '{}');
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return { data, bump(k) { (data[today] = data[today] || {})[k] = (data[today][k] || 0) + 1; } };
}
function saveStats_(stats) {
  const keys = Object.keys(stats.data).sort();
  keys.slice(0, Math.max(0, keys.length - 60)).forEach(k => delete stats.data[k]);
  PropertiesService.getUserProperties().setProperty('stats', JSON.stringify(stats.data));
}

// ───────────────────────────── Helpers ─────────────────────────────

const EMAIL_RE = /[A-Z0-9._%+'-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;
const labelCache_ = {};
let myAddrs_ = null;

function getLabel_(name) {
  return labelCache_[name] || (labelCache_[name] = GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name));
}
function addr_(from) { const m = String(from).match(EMAIL_RE); return m ? m[m.length - 1].toLowerCase() : ''; }
function nameOf_(from) { const m = String(from).match(/^\s*"?([^"<]+?)"?\s*</); return m ? m[1] : addr_(from); }
function recipients_(m) { return ((m.getTo() + ',' + m.getCc()).match(EMAIL_RE) || []).map(s => s.toLowerCase()); }
function me_() { return Session.getEffectiveUser().getEmail(); }
function myAddresses_() {
  if (!myAddrs_) myAddrs_ = [me_()].concat(GmailApp.getAliases()).map(s => s.toLowerCase());
  return myAddrs_;
}
function isMe_(from) { return myAddresses_().includes(addr_(from)); }
function vipDomains_() {
  return CONFIG.VIP_DOMAINS.concat(myAddresses_().map(a => a.split('@')[1]))
    .map(d => d.toLowerCase())
    .filter(d => !/^(gmail|googlemail)\.com$/.test(d));
}
function senderOf_(thread) {
  const msgs = thread.getMessages();
  for (let i = msgs.length - 1; i >= 0; i--) if (!isMe_(msgs[i].getFrom())) return addr_(msgs[i].getFrom());
  return '';
}
function epoch_(ms) { return Math.floor(ms / 1000); }
function threadsWith_(name) { return getLabel_(name).getThreads(0, 100); }
function hasLabel_(thread, name) { return thread.getLabels().some(l => l.getName() === name); }
function esc_(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function webAppUrl_() { try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } }
