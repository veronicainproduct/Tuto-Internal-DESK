/**
 * TUTO Desk — backend (Google Apps Script)
 *
 * Deploy as a web app: Execute as "Me", Who has access "Anyone".
 * Sign-in: "Continue with Google" (also grants read-only access to that
 * person's calendar, so their calls load automatically) or an emailed 6-digit
 * code. Either gives a 30-day session. Only emails on the "Team" tab of the
 * data sheet (or in the ALLOWED_EMAILS property) can sign in. WhatsApp numbers
 * come from that same Team tab, so teammates don't have to enter anything.
 *
 * Visibility: tasks in the "Personal" category are seen only by the person
 * who created them. Every other task is seen by the whole team.
 *
 * Calendars are read only to show each person their own calls and to send
 * them reminders; nobody else ever sees another person's calls in the app.
 */

const PROPS = PropertiesService.getScriptProperties();

const TABLES = {
  Users: ['email', 'name', 'picture', 'phone', 'prefs', 'calKey', 'calUrl', 'lastSeen', 'ical', 'gtoken'],
  Tasks: ['id', 'title', 'notes', 'due', 'dueTime', 'priority', 'category', 'spoc', 'tags', 'status',
          'createdBy', 'createdAt', 'updatedAt', 'completedAt', 'source'],
  Comments: ['id', 'taskId', 'author', 'text', 'createdAt'],
  Prompts: ['id', 'email', 'title', 'endedAt', 'attendees', 'status', 'createdAt'],
  Notifications: ['id', 'key', 'email', 'kind', 'text', 'status', 'createdAt'],
  Sessions: ['id', 'email', 'device', 'createdAt', 'expiresAt'],
};
const SESSION_DAYS = 30;
const TEAM_SHEET = 'Team';
const GOOGLE_SCOPES = 'openid email profile https://www.googleapis.com/auth/calendar.readonly';

const PERSONAL = 'Personal';
const CATEGORIES = ['Work', 'Meetings', 'Follow-ups', PERSONAL];
const PRIORITIES = ['low', 'medium', 'high'];
const STATUSES = ['todo', 'doing', 'done'];
const DEFAULT_PREFS = { callAlerts: true, newCalls: true, digest: true, overdue: true, assign: true, tagged: true, actionItems: true };
const PREF_FOR = { call: 'callAlerts', newcall: 'newCalls', digest: 'digest', overdue: 'overdue', assign: 'assign',
                   tagged: 'tagged', actions: 'actionItems', test: null };
const EMAIL_RE = /^[^\s@,;]+@[^\s@,;]+\.[a-z]{2,}$/i;

// ---------------------------------------------------------------- helpers

function prop_(key, fallback) { const v = PROPS.getProperty(key); return v == null || v === '' ? fallback : v; }
function fmt_(d, pattern) { return Utilities.formatDate(d, Session.getScriptTimeZone(), pattern); }
function uid_(p) { return p + '_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }
function fail_(code, msg) { const e = new Error(msg); e.code = code; return e; }
function split_(s) { return String(s || '').split(',').map(e => e.trim().toLowerCase()).filter(Boolean); }
function clean_(r) { const o = Object.assign({}, r); delete o._row; return o; }
function link_(q) { const b = prop_('APP_URL', ''); return b ? '\n' + b + (q || '') : ''; }
function hm12_(hm) { const [h, m] = hm.split(':').map(Number); return (h % 12 || 12) + (m ? ':' + String(m).padStart(2, '0') : '') + (h < 12 ? ' AM' : ' PM'); }
function dueAt_(t) {
  const [y, mo, d] = t.due.split('-').map(Number);
  const [h, mi] = t.dueTime ? t.dueTime.split(':').map(Number) : [23, 59];
  return new Date(y, mo - 1, d, h, mi);
}
function key_() { return Utilities.getUuid().replace(/-/g, ''); }
function sha_(s) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s)).map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join(''); }

// ---------------------------------------------------------------- setup

/** Run once from the editor (safe to run again). */
function setup() {
  const id = prop_('SHEET_ID', '');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create('TUTO Desk — data');
  PROPS.setProperty('SHEET_ID', ss.getId());
  Object.keys(TABLES).forEach(name => {
    const cols = TABLES[name];
    let sh = ss.getSheetByName(name);
    if (sh && sh.getLastColumn() > 0) {
      const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].filter(String);
      if (head.join('|') !== cols.join('|')) { sh.setName(name + ' (old ' + fmt_(new Date(), 'd MMM HH:mm') + ')'); sh = null; }
    }
    sh = sh || ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  let team = ss.getSheetByName(TEAM_SHEET);
  if (!team) {
    team = ss.insertSheet(TEAM_SHEET, 0);
    team.getRange('A:C').setNumberFormat('@');
    const rows = [['Email', 'Name', 'WhatsApp (with +91)']];
    const seed = [Session.getEffectiveUser().getEmail().toLowerCase()].concat(split_(prop_('ALLOWED_EMAILS', '')));
    seed.filter((e, i, a) => e && a.indexOf(e) === i).forEach(e => rows.push([e, '', '']));
    team.getRange(1, 1, rows.length, 3).setValues(rows);
    team.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#F3E3FB');
    team.setFrozenRows(1);
    team.setColumnWidths(1, 3, 260);
  }
  const blank = ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1) ss.deleteSheet(blank);

  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('tick').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('morningDigest').timeBased().atHour(9).nearMinute(30).everyDays(1).create();
  ScriptApp.newTrigger('overdueSweep').timeBased().everyHours(1).create();
  MailApp.getRemainingDailyQuota();
  Logger.log('Done. Data sheet: ' + ss.getUrl());
}

// ---------------------------------------------------------------- data layer

const _mem = {};
function ss_() { return _mem.ss || (_mem.ss = SpreadsheetApp.openById(prop_('SHEET_ID'))); }
function sh_(n) { return ss_().getSheetByName(n); }

function all_(name) {
  if (_mem[name]) return _mem[name];
  const sh = sh_(name), cols = TABLES[name], last = sh.getLastRow();
  const vals = last < 2 ? [] : sh.getRange(2, 1, last - 1, cols.length).getDisplayValues();
  _mem[name] = vals.map((r, i) => { const o = { _row: i + 2 }; cols.forEach((c, j) => { o[c] = r[j]; }); return o; })
                   .filter(o => o[cols[0]]);
  return _mem[name];
}

function put_(name, obj) {
  const cols = TABLES[name], rows = all_(name), sh = sh_(name);
  const hit = rows.find(r => r[cols[0]] === String(obj[cols[0]]));
  const row = hit ? hit._row : Math.max(sh.getLastRow(), 1) + 1;
  const vals = cols.map(c => (obj[c] == null ? '' : String(obj[c])));
  sh.getRange(row, 1, 1, cols.length).setNumberFormat('@').setValues([vals]);
  const rec = { _row: row };
  cols.forEach((c, j) => { rec[c] = vals[j]; });
  if (hit) Object.assign(hit, rec); else rows.push(rec);
  return rec;
}

function drop_(name, ids) {
  if (!ids.length) return;
  const k = TABLES[name][0], sh = sh_(name);
  all_(name).filter(r => ids.indexOf(r[k]) >= 0).sort((a, b) => b._row - a._row).forEach(r => sh.deleteRow(r._row));
  delete _mem[name];
}

// ---------------------------------------------------------------- team & auth

/** Rows of the "Team" tab: Email | Name | WhatsApp. The owner fills this in once. */
function teamRows_() {
  if (_mem.team) return _mem.team;
  const sh = ss_().getSheetByName(TEAM_SHEET), rows = [];
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 3).getDisplayValues().forEach(r => {
      const email = String(r[0]).trim().toLowerCase();
      if (EMAIL_RE.test(email)) rows.push({ email: email, name: String(r[1]).trim(), phone: String(r[2]).trim() });
    });
  }
  return (_mem.team = rows);
}
function teamRow_(email) { return teamRows_().find(r => r.email === email); }
function allowedList_() {
  const l = teamRows_().map(r => r.email);
  split_(prop_('ALLOWED_EMAILS', '')).forEach(e => { if (l.indexOf(e) < 0) l.push(e); });
  return l;
}
function allowed_(email) { const l = allowedList_(); return !l.length || l.indexOf(email) >= 0; }
function user_(email) { return all_('Users').find(u => u.email === email); }
function nameOf_(email) { const t = teamRow_(email), u = user_(email); return (t && t.name) || (u && u.name) || email.split('@')[0]; }
function phoneOf_(email) { const t = teamRow_(email), u = user_(email); return (u && u.phone) || (t && t.phone) || ''; }
function handleOf_(email) { return nameOf_(email).split(/\s+/)[0].toLowerCase().replace(/[^a-z]/g, ''); }

/** Everyone on the team: invited emails plus anyone who has signed in. */
function team_() {
  const out = allowedList_().slice();
  all_('Users').forEach(u => { if (out.indexOf(u.email) < 0 && allowed_(u.email)) out.push(u.email); });
  return out;
}

function verify_(token) {
  if (!token || String(token).indexOf('sess_') !== 0) throw fail_('AUTH', 'Sign in to continue.');
  const s = all_('Sessions').find(x => x.id === sha_(token));
  if (!s || Date.parse(s.expiresAt) < Date.now()) throw fail_('AUTH', 'Session expired. Sign in again.');
  if (!allowed_(s.email)) throw fail_('AUTH', 'Access removed.');
  const u = user_(s.email);
  return { email: s.email, name: u ? u.name : s.email.split('@')[0], picture: u ? u.picture : '', token: token };
}

function newSession_(user, device) {
  const token = 'sess_' + key_() + key_();
  const now = Date.now();
  put_('Sessions', { id: sha_(token), email: user.email, device: String(device || '').slice(0, 60),
                     createdAt: new Date(now).toISOString(), expiresAt: new Date(now + SESSION_DAYS * 864e5).toISOString() });
  ensureUser_(user);
  return { session: token };
}

/** Actions that work before sign-in. */
const PUBLIC = {
  /** "Continue with Google": one consent gives sign-in plus read-only calendar access. */
  googleCode(p) {
    const g = exchangeCode_(p.code);
    if (!allowed_(g.email)) throw fail_('FORBIDDEN', g.email + ' isn\'t on the team. Ask Veronica to add it to the Team tab.');
    const out = newSession_({ email: g.email, name: g.name || g.email.split('@')[0], picture: g.picture }, p.device);
    storeGoogle_(g);
    return out;
  },

  sendCode(p) {
    const email = String(p.email || '').trim().toLowerCase();
    if (!EMAIL_RE.test(email)) throw fail_('BAD_REQUEST', 'Enter a valid email.');
    if (allowedList_().indexOf(email) < 0) throw fail_('FORBIDDEN', 'This email isn\'t invited. Ask the owner to add it.');
    const cache = CacheService.getScriptCache();
    const sent = Number(cache.get('sent_' + email) || 0);
    if (sent >= 5) throw fail_('LIMIT', 'Too many codes. Try again in an hour.');
    const code = String(parseInt(key_().slice(0, 10), 16) % 1000000).padStart(6, '0');
    cache.put('code_' + email, JSON.stringify({ h: sha_(code), tries: 0 }), 600);
    cache.put('sent_' + email, String(sent + 1), 3600);
    MailApp.sendEmail(email, 'TUTO Desk code: ' + code, 'Your TUTO Desk sign-in code is ' + code + '.\nIt expires in 10 minutes.');
    return { sent: true };
  },

  verifyCode(p) {
    const email = String(p.email || '').trim().toLowerCase();
    const cache = CacheService.getScriptCache();
    const raw = cache.get('code_' + email);
    if (!raw) throw fail_('AUTH', 'Code expired. Send a new one.');
    const c = JSON.parse(raw);
    if (c.tries >= 5) { cache.remove('code_' + email); throw fail_('AUTH', 'Too many tries. Send a new code.'); }
    if (sha_(String(p.code || '').trim()) !== c.h) {
      c.tries++; cache.put('code_' + email, JSON.stringify(c), 600);
      throw fail_('AUTH', 'Wrong code. Check the latest email.');
    }
    cache.remove('code_' + email);
    const u = user_(email);
    return newSession_({ email: email, name: u ? u.name : email.split('@')[0], picture: u ? u.picture : '' }, p.device);
  },

  config() {
    const on = !!(prop_('GOOGLE_CLIENT_ID', '') && prop_('GOOGLE_CLIENT_SECRET', ''));
    return { google: on, clientId: on ? prop_('GOOGLE_CLIENT_ID', '') : '', scopes: GOOGLE_SCOPES };
  },
};

function exchangeCode_(code) {
  const id = prop_('GOOGLE_CLIENT_ID', ''), secret = prop_('GOOGLE_CLIENT_SECRET', '');
  if (!id || !secret) throw fail_('AUTH', 'Google sign-in isn\'t set up yet. Use the email code.');
  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post', muteHttpExceptions: true,
    payload: { code: String(code || ''), client_id: id, client_secret: secret, redirect_uri: 'postmessage', grant_type: 'authorization_code' },
  });
  const j = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200 || !j.id_token) throw fail_('AUTH', 'Google sign-in failed (' + (j.error_description || j.error || res.getResponseCode()) + '). Try again.');
  let part = j.id_token.split('.')[1];
  part += '='.repeat((4 - part.length % 4) % 4);
  const c = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(part)).getDataAsString());
  if (c.aud !== id || String(c.email_verified) !== 'true') throw fail_('AUTH', 'Google sign-in failed. Try again.');
  return { email: String(c.email).toLowerCase(), name: c.name || '', picture: c.picture || '',
           access: j.access_token || '', expires: Number(j.expires_in || 3600), refresh: j.refresh_token || '',
           calendar: String(j.scope || '').indexOf('calendar') >= 0 };
}

/** Keeps the person's calendar access (refresh token) so reminders keep working without them doing anything. */
function storeGoogle_(g) {
  if (!g.calendar) return;
  const rec = user_(g.email);
  if (!rec) return;
  if (g.access) CacheService.getScriptCache().put('gat_' + sha_(g.email).slice(0, 20), g.access, Math.max(60, g.expires - 120));
  if (g.refresh) put_('Users', Object.assign({}, rec, { gtoken: g.refresh }));
  CacheService.getScriptCache().remove('cal_' + sha_(g.email).slice(0, 20));
}

// ---------------------------------------------------------------- HTTP

function doGet() { return json_({ ok: true, service: 'TUTO Desk API' }); }

function doPost(e) {
  let out;
  const lock = LockService.getScriptLock();
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    lock.waitLock(20000);
    if (PUBLIC[body.action]) out = { ok: true, data: PUBLIC[body.action](body.payload || {}) };
    else {
      const user = verify_(body.token);
      const fn = API[body.action];
      if (!fn) throw fail_('BAD_REQUEST', 'Unknown action.');
      out = { ok: true, data: fn(user, body.payload || {}) };
    }
  } catch (err) {
    console.error(err);
    out = { ok: false, code: err.code || 'ERROR', error: err.message || String(err) };
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
  return json_(out);
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// ---------------------------------------------------------------- domain

function prefs_(u) { let p = {}; try { p = JSON.parse(u.prefs || '{}'); } catch (e) {} return Object.assign({}, DEFAULT_PREFS, p); }
function canSee_(t, email) { return t.category !== PERSONAL || t.createdBy === email; }
function taskOut_(t) { return Object.assign(clean_(t), { tags: split_(t.tags) }); }

/** @handles → team emails. Handle = first name or email prefix. */
function mentions_(text, people) {
  const out = [];
  String(text || '').replace(/(^|\s)@([\w.+-]+)/g, (m, sp, h) => {
    h = h.toLowerCase();
    const e = people.find(p => p.split('@')[0] === h) || people.find(p => handleOf_(p) === h);
    if (e && out.indexOf(e) < 0) out.push(e);
    return m;
  });
  return out;
}

function ensureUser_(u) {
  const now = new Date().toISOString();
  let rec = user_(u.email);
  if (!rec) {
    rec = put_('Users', { email: u.email, name: u.name, picture: u.picture, phone: '', prefs: JSON.stringify(DEFAULT_PREFS),
                          calKey: key_(), calUrl: '', lastSeen: now });
  } else if (!rec.calKey || Date.now() - Date.parse(rec.lastSeen || 0) > 10 * 60000) {
    rec = put_('Users', Object.assign({}, rec, { lastSeen: now, calKey: rec.calKey || key_(), picture: u.picture || rec.picture }));
  }
  return rec;
}

function meOut_(r) {
  return { email: r.email, name: nameOf_(r.email), picture: r.picture, phone: phoneOf_(r.email), prefs: prefs_(r),
           calendar: !!(r.gtoken || r.ical), calSource: r.gtoken ? 'google' : (r.ical ? 'link' : '') };
}

function saveTask_(u, t) {
  const prev = t.id ? all_('Tasks').find(x => x.id === t.id) : null;
  if (prev && !canSee_(prev, u.email)) throw fail_('FORBIDDEN', 'Task not found.');
  const title = String(t.title || '').trim().slice(0, 300);
  if (!title) throw fail_('BAD_REQUEST', 'Add a title.');
  const category = CATEGORIES.indexOf(t.category) >= 0 ? t.category : 'Work';
  const owner = prev ? prev.createdBy : u.email;
  const personal = category === PERSONAL;
  if (personal && owner !== u.email) throw fail_('FORBIDDEN', 'Only the creator can make a task Personal.');

  const team = team_();
  let spoc = String(t.spoc || '').trim().toLowerCase();
  if (personal) spoc = '';
  if (spoc && team.indexOf(spoc) < 0) throw fail_('BAD_REQUEST', spoc + ' isn\'t on the team.');
  const notes = String(t.notes || '').slice(0, 4000);
  const tags = personal ? [] : mentions_(title + ' ' + notes, team);
  const now = new Date().toISOString();
  const status = STATUSES.indexOf(t.status) >= 0 ? t.status : 'todo';
  const rec = {
    id: prev ? prev.id : (/^t_[a-z0-9]{4,24}$/.test(t.id || '') ? t.id : uid_('t')),
    title: title, notes: notes,
    due: /^\d{4}-\d{2}-\d{2}$/.test(t.due || '') ? t.due : '',
    dueTime: /^\d{2}:\d{2}$/.test(t.dueTime || '') ? t.dueTime : '',
    priority: PRIORITIES.indexOf(t.priority) >= 0 ? t.priority : 'medium',
    category: category, spoc: spoc, tags: tags.join(','), status: status,
    createdBy: owner, createdAt: prev ? prev.createdAt : now, updatedAt: now,
    completedAt: status === 'done' ? ((prev && prev.completedAt) || now) : '',
    source: prev ? prev.source : String(t.source || '').slice(0, 40),
  };
  put_('Tasks', rec);

  if (!personal) {
    const who = nameOf_(u.email), q = link_('?task=' + rec.id);
    if (spoc && spoc !== u.email && (!prev || prev.spoc !== spoc)) {
      notify_(spoc, 'assign', who + ' made you SPOC:\n' + rec.title + (rec.due ? '\nDue ' + fmt_(dueAt_(rec), 'EEE d MMM') : '') + q,
              'assign:' + rec.id + ':' + spoc + ':' + now);
    }
    const before = prev ? split_(prev.tags) : [];
    tags.filter(e => e !== u.email && e !== spoc && before.indexOf(e) < 0)
        .forEach(e => notify_(e, 'tagged', who + ' tagged you:\n' + rec.title + q, 'tag:' + rec.id + ':' + e));
  }
  return taskOut_(rec);
}

// ---------------------------------------------------------------- API (signed in)

const API = {
  bootstrap(u) {
    const me = ensureUser_(u);
    const tasks = all_('Tasks').filter(t => canSee_(t, u.email));
    const ids = {}; tasks.forEach(t => { ids[t.id] = true; });
    return {
      me: meOut_(me),
      users: team_().map(e => { const r = user_(e); return { email: e, name: nameOf_(e), picture: r ? r.picture : '', joined: !!r }; }),
      tasks: tasks.map(taskOut_),
      comments: all_('Comments').filter(c => ids[c.taskId]).map(clean_),
      prompts: all_('Prompts').filter(p => p.email === u.email && p.status === 'pending').map(p => Object.assign(clean_(p), { attendees: split_(p.attendees) })),
      notifications: all_('Notifications').filter(n => n.email === u.email).slice(-40).reverse()
        .map(n => ({ id: n.id, kind: n.kind, email: n.email, text: n.text, status: n.status, createdAt: n.createdAt })),
      categories: CATEGORIES,
    };
  },

  saveTask(u, p) { return saveTask_(u, p.task || {}); },

  saveTasks(u, p) {
    const out = (p.tasks || []).slice(0, 50).map(t => saveTask_(u, t));
    if (p.promptId) API.resolvePrompt(u, { id: p.promptId, status: 'done' });
    return out;
  },

  deleteTask(u, p) {
    const t = all_('Tasks').find(x => x.id === p.id);
    if (!t) return { id: p.id };
    if (!canSee_(t, u.email)) throw fail_('FORBIDDEN', 'Task not found.');
    drop_('Tasks', [t.id]);
    drop_('Comments', all_('Comments').filter(c => c.taskId === t.id).map(c => c.id));
    return { id: t.id };
  },

  addComment(u, p) {
    const t = all_('Tasks').find(x => x.id === p.taskId);
    if (!t || !canSee_(t, u.email)) throw fail_('NOT_FOUND', 'Task not found.');
    const text = String(p.text || '').trim().slice(0, 2000);
    if (!text) throw fail_('BAD_REQUEST', 'Write something first.');
    const rec = put_('Comments', { id: uid_('c'), taskId: t.id, author: u.email, text: text, createdAt: new Date().toISOString() });
    if (t.category !== PERSONAL) {
      mentions_(text, team_()).filter(e => e !== u.email).forEach(e =>
        notify_(e, 'tagged', nameOf_(u.email) + ' on "' + t.title + '":\n' + text + link_('?task=' + t.id), 'ctag:' + rec.id + ':' + e));
    }
    return clean_(rec);
  },

  saveProfile(u, p) {
    const rec = ensureUser_(u);
    const prefs = prefs_(rec);
    Object.keys(DEFAULT_PREFS).forEach(k => { if (p.prefs && typeof p.prefs[k] === 'boolean') prefs[k] = p.prefs[k]; });
    let ical = rec.ical;
    if (typeof p.ical === 'string') {
      ical = p.ical.trim();
      if (ical) {
        if (!/^https:\/\//i.test(ical) || /\s/.test(ical)) throw fail_('BAD_REQUEST', 'Paste the full link that starts with https://');
        icsWindow_(ical, u.email, new Date(), new Date(Date.now() + 864e5));   // throws a clear error if the link doesn't work
      }
      CacheService.getScriptCache().remove('cal_' + sha_(u.email).slice(0, 20));
    }
    return meOut_(put_('Users', Object.assign({}, rec, {
      name: String(p.name || rec.name).trim().slice(0, 60) || rec.name,
      phone: String(p.phone || '').replace(/[^\d+ ]/g, '').trim().slice(0, 20),
      prefs: JSON.stringify(prefs), ical: ical,
    })));
  },

  connectCalendar(u, p) {
    const g = exchangeCode_(p.code);
    if (g.email !== u.email) throw fail_('BAD_REQUEST', 'Pick ' + u.email + ' in the Google window.');
    if (!g.calendar) throw fail_('BAD_REQUEST', 'Tick "See your calendar" in the Google window to show your calls.');
    storeGoogle_(g);
    return meOut_(user_(u.email));
  },

  disconnectCalendar(u) {
    const rec = user_(u.email);
    if (rec && rec.gtoken) { try { UrlFetchApp.fetch('https://oauth2.googleapis.com/revoke?token=' + encodeURIComponent(rec.gtoken), { method: 'post', muteHttpExceptions: true }); } catch (e) {} }
    const next = put_('Users', Object.assign({}, rec, { gtoken: '', ical: '' }));
    CacheService.getScriptCache().remove('cal_' + sha_(u.email).slice(0, 20));
    CacheService.getScriptCache().remove('gat_' + sha_(u.email).slice(0, 20));
    return meOut_(next);
  },

  calendar(u, p) {
    const from = new Date(p.from), to = new Date(p.to);
    if (isNaN(from) || isNaN(to) || to <= from || to - from > 43 * 864e5) throw fail_('BAD_REQUEST', 'Bad date range.');
    const rec = user_(u.email);
    if (!rec || !(rec.gtoken || rec.ical)) return [];
    try { return eventsCached_(rec, from, to); }
    catch (e) { throw fail_('CAL', 'Couldn\'t read your calendar: ' + e.message); }
  },

  resolvePrompt(u, p) {
    const pr = all_('Prompts').find(x => x.id === p.id && x.email === u.email);
    if (pr) put_('Prompts', Object.assign({}, pr, { status: p.status === 'done' ? 'done' : 'dismissed' }));
    return { id: p.id };
  },

  signOut(u) { drop_('Sessions', [sha_(u.token)]); return { ok: true }; },

  testNotify(u) { notify_(u.email, 'test', 'TUTO Desk test. Nudges will land here.', null); return { ok: true }; },
};

// ---------------------------------------------------------------- calendar

/** A person's own events: from Google (connected at sign-in) or a pasted iCal link. */
function calendarEvents_(rec, from, to) {
  if (rec.gtoken) return googleEvents_(rec, from, to);
  if (rec.ical) return icsWindow_(rec.ical, rec.email, from, to);
  return [];
}

function googleAccess_(rec) {
  const cache = CacheService.getScriptCache(), key = 'gat_' + sha_(rec.email).slice(0, 20);
  const hit = cache.get(key);
  if (hit) return hit;
  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post', muteHttpExceptions: true,
    payload: { client_id: prop_('GOOGLE_CLIENT_ID', ''), client_secret: prop_('GOOGLE_CLIENT_SECRET', ''), refresh_token: rec.gtoken, grant_type: 'refresh_token' },
  });
  const j = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200) {
    if (j.error === 'invalid_grant') put_('Users', Object.assign({}, rec, { gtoken: '' }));
    throw fail_('CAL', 'Calendar access ended. Sign out and continue with Google again.');
  }
  cache.put(key, j.access_token, Math.max(60, Number(j.expires_in || 3600) - 120));
  return j.access_token;
}

function googleEvents_(rec, from, to) {
  const token = googleAccess_(rec);
  const url = 'https://www.googleapis.com/calendar/v3/calendars/primary/events?singleEvents=true&orderBy=startTime&maxResults=250' +
              '&timeMin=' + encodeURIComponent(from.toISOString()) + '&timeMax=' + encodeURIComponent(to.toISOString());
  const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    const msg = res.getContentText();
    if (/has not been used|is disabled/i.test(msg)) throw fail_('CAL', 'Google Calendar API is off. Veronica: enable it in Google Cloud (setup step 2.3).');
    throw fail_('CAL', 'Couldn\'t read the calendar (error ' + res.getResponseCode() + ').');
  }
  const day = d => Utilities.parseDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return (JSON.parse(res.getContentText()).items || []).filter(ev => {
    if (ev.status === 'cancelled') return false;
    const me = (ev.attendees || []).find(a => a.self);
    return !(me && me.responseStatus === 'declined');
  }).map(ev => {
    const allDay = !ev.start.dateTime;
    const s = allDay ? day(ev.start.date) : new Date(ev.start.dateTime);
    const e = allDay ? day(ev.end.date) : new Date(ev.end.dateTime);
    return { id: sha_(ev.id).slice(0, 24), title: ev.summary || 'Busy', start: s.toISOString(), end: e.toISOString(), allDay: allDay,
             attendees: (ev.attendees || []).map(a => String(a.email || '').toLowerCase()), location: ev.hangoutLink ? 'Google Meet' : (ev.location || ''),
             created: ev.created || '' };
  });
}


function eventsCached_(rec, from, to) {
  const cache = CacheService.getScriptCache();
  const key = 'cal_' + sha_(rec.email).slice(0, 20);
  const hit = cache.get(key);
  if (hit) {
    const c = JSON.parse(hit);
    if (c.from <= from.getTime() && c.to >= to.getTime()) {
      return c.events.filter(e => Date.parse(e.end) > from.getTime() && Date.parse(e.start) < to.getTime());
    }
  }
  const wFrom = new Date(Math.min(from.getTime(), Date.now() - 864e5)), wTo = new Date(Math.max(to.getTime(), Date.now() + 8 * 864e5));
  const events = calendarEvents_(rec, wFrom, wTo);
  const payload = JSON.stringify({ from: wFrom.getTime(), to: wTo.getTime(), events: events });
  if (payload.length < 95000) cache.put(key, payload, 300);
  return events.filter(e => Date.parse(e.end) > from.getTime() && Date.parse(e.start) < to.getTime());
}

/** Fetches a person's own .ics feed and returns their events between from and to (recurring ones expanded). */
function icsWindow_(url, email, from, to) {
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  if (res.getResponseCode() !== 200) throw fail_('CAL', 'The calendar link didn\'t work (error ' + res.getResponseCode() + '). Copy it again.');
  const text = res.getContentText();
  if (text.indexOf('BEGIN:VCALENDAR') < 0) throw fail_('CAL', 'That isn\'t a calendar link. Use "Secret address in iCal format".');
  const cal = icsParse_(text);
  const moved = {};
  cal.events.filter(e => e.recurrenceId).forEach(e => { (moved[e.uid] = moved[e.uid] || {})[e.recurrenceId] = true; });
  const out = [];
  const add = (e, startMs) => {
    const dur = e.end ? e.end.getTime() - e.start.getTime() : durMs_(e.duration, e.allDay);
    const s = new Date(startMs), en = new Date(startMs + dur);
    if (en <= from || s >= to || e.status === 'CANCELLED') return;
    const me = e.attendees.find(a => a.email === email);
    if (me && me.partstat === 'DECLINED') return;
    out.push({ id: sha_(e.uid + '@' + startMs).slice(0, 24), title: e.title || 'Busy', start: s.toISOString(), end: en.toISOString(),
               allDay: !!e.allDay, attendees: e.attendees.map(a => a.email), location: e.location || '', created: e.created ? e.created.toISOString() : '' });
  };
  cal.events.forEach(e => {
    if (!e.start) return;
    if (e.recurrenceId || !e.rrule) { add(e, e.start.getTime()); return; }
    occurrences_(e, to).forEach(t => {
      if (e.exdates.indexOf(t) >= 0 || (moved[e.uid] && moved[e.uid][t])) return;
      add(e, t);
    });
  });
  return out.sort((a, b) => (a.start < b.start ? -1 : 1));
}

function icsParse_(text) {
  const lines = text.replace(/\r?\n[ \t]/g, '').split(/\r?\n/);
  const events = [];
  let cur = null, calTz = '';
  lines.forEach(line => {
    if (line === 'BEGIN:VEVENT') { cur = { exdates: [], attendees: [] }; return; }
    if (line === 'END:VEVENT') { if (cur) events.push(cur); cur = null; return; }
    const i = line.indexOf(':');
    if (i < 0) return;
    const parts = line.slice(0, i).split(';'), val = line.slice(i + 1), name = parts[0].toUpperCase();
    const params = {};
    parts.slice(1).forEach(x => { const j = x.indexOf('='); if (j > 0) params[x.slice(0, j).toUpperCase()] = x.slice(j + 1).replace(/^"|"$/g, ''); });
    if (!cur) { if (name === 'X-WR-TIMEZONE') calTz = val.trim(); return; }
    try {
      switch (name) {
        case 'UID': cur.uid = val; break;
        case 'SUMMARY': cur.title = val.replace(/\\n/gi, ' ').replace(/\\([,;\\])/g, '$1'); break;
        case 'LOCATION': cur.location = val.replace(/\\n/gi, ' ').replace(/\\([,;\\])/g, '$1'); break;
        case 'STATUS': cur.status = val.trim().toUpperCase(); break;
        case 'DTSTART': cur.start = icsDate_(val, params.TZID || calTz); cur.allDay = params.VALUE === 'DATE' || /^\d{8}$/.test(val.trim()); break;
        case 'DTEND': cur.end = icsDate_(val, params.TZID || calTz); break;
        case 'DURATION': cur.duration = val.trim(); break;
        case 'RRULE': cur.rrule = val.trim(); break;
        case 'EXDATE': val.split(',').forEach(v => cur.exdates.push(icsDate_(v, params.TZID || calTz).getTime())); break;
        case 'RECURRENCE-ID': cur.recurrenceId = icsDate_(val, params.TZID || calTz).getTime(); break;
        case 'CREATED': cur.created = icsDate_(val, 'UTC'); break;
        case 'ATTENDEE': cur.attendees.push({ email: val.replace(/^mailto:/i, '').trim().toLowerCase(), partstat: (params.PARTSTAT || '').toUpperCase() }); break;
      }
    } catch (e) { /* skip one unreadable field, keep the event */ }
  });
  return { events: events.filter(e => e.uid) };
}

function icsDate_(v, tz) {
  v = String(v).trim();
  const local = Session.getScriptTimeZone();
  if (/^\d{8}$/.test(v)) return Utilities.parseDate(v, local, 'yyyyMMdd');
  if (/Z$/i.test(v)) return Utilities.parseDate(v.slice(0, 15), 'UTC', "yyyyMMdd'T'HHmmss");
  try { return Utilities.parseDate(v.slice(0, 15), tz || local, "yyyyMMdd'T'HHmmss"); }
  catch (e) { return Utilities.parseDate(v.slice(0, 15), local, "yyyyMMdd'T'HHmmss"); }
}

function durMs_(d, allDay) {
  const m = /^P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/.exec(String(d || ''));
  if (!m) return allDay ? 864e5 : 0;
  return (((+m[1] || 0) * 7 + (+m[2] || 0)) * 864e5) + (+m[3] || 0) * 36e5 + (+m[4] || 0) * 6e4 + (+m[5] || 0) * 1e3;
}

/** Start times (ms) of a recurring event up to `to`. Covers DAILY / WEEKLY / MONTHLY / YEARLY with the usual options. */
function occurrences_(e, to) {
  const r = {};
  e.rrule.split(';').forEach(x => { const [k, v] = x.split('='); r[String(k).toUpperCase()] = v; });
  const interval = Math.max(1, Number(r.INTERVAL || 1));
  const count = r.COUNT ? Number(r.COUNT) : Infinity;
  const until = r.UNTIL ? icsDate_(r.UNTIL, 'UTC').getTime() : Infinity;
  const limit = Math.min(to.getTime(), until);
  const start = e.start;
  const DAYS = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];
  const byDay = r.BYDAY ? r.BYDAY.split(',').map(d => ({ n: parseInt(d, 10) || 0, wd: DAYS.indexOf(d.slice(-2)) })) : null;
  const out = [];
  let n = 0, guard = 0;
  const push = t => { if (t >= start.getTime() && t <= limit && n < count) { n++; out.push(t); } };

  if (r.FREQ === 'DAILY' || r.FREQ === 'WEEKLY') {
    const startMonday = new Date(start); startMonday.setHours(0, 0, 0, 0); startMonday.setDate(startMonday.getDate() - ((startMonday.getDay() + 6) % 7));
    const wds = byDay ? byDay.map(b => b.wd) : (r.FREQ === 'WEEKLY' ? [start.getDay()] : null);
    for (let d = new Date(start); d.getTime() <= limit && n < count && guard++ < 40000; d.setDate(d.getDate() + 1)) {
      const k = Math.round((new Date(d.getFullYear(), d.getMonth(), d.getDate()) - new Date(start.getFullYear(), start.getMonth(), start.getDate())) / 864e5);
      if (r.FREQ === 'DAILY' && k % interval !== 0) continue;
      if (r.FREQ === 'WEEKLY') {
        const dm = new Date(d); dm.setHours(0, 0, 0, 0); dm.setDate(dm.getDate() - ((dm.getDay() + 6) % 7));
        const w = Math.round((dm - startMonday) / (7 * 864e5));
        if (w % interval !== 0) continue;
      }
      if (wds && wds.indexOf(d.getDay()) < 0) continue;
      push(d.getTime());
    }
  } else if (r.FREQ === 'MONTHLY' || r.FREQ === 'YEARLY') {
    const step = r.FREQ === 'MONTHLY' ? interval : 12 * interval;
    for (let i = 0; guard++ < 2000 && n < count; i += step) {
      const base = new Date(start.getFullYear(), start.getMonth() + i, 1, start.getHours(), start.getMinutes(), start.getSeconds());
      if (base.getTime() > limit) break;
      let days = [];
      if (byDay && byDay.some(b => b.n)) {
        byDay.forEach(b => {
          const matches = [];
          for (let dd = new Date(base); dd.getMonth() === base.getMonth(); dd.setDate(dd.getDate() + 1)) if (dd.getDay() === b.wd) matches.push(new Date(dd));
          const pick = b.n > 0 ? matches[b.n - 1] : matches[matches.length + b.n];
          if (pick) days.push(pick);
        });
      } else {
        const dom = r.BYMONTHDAY ? Number(r.BYMONTHDAY) : start.getDate();
        const dd = new Date(base); dd.setDate(dom);
        if (dd.getMonth() === base.getMonth()) days.push(dd);
      }
      days.sort((a, b) => a - b).forEach(dd => push(dd.getTime()));
    }
  } else {
    push(start.getTime());
  }
  return out;
}

/** Every 5 minutes: reminders, new calls and "call ended" prompts, per person, from their own calendar. */
function tick() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    const now = new Date();
    const last = new Date(Number(prop_('LAST_TICK', now.getTime() - 5 * 60000)));
    all_('Users').filter(u => (u.gtoken || u.ical) && allowed_(u.email)).forEach(u => {
      let evs;
      try { evs = calendarEvents_(u, new Date(now - 40 * 60000), new Date(+now + 7 * 864e5)); }
      catch (e) { console.warn(u.email + ': ' + e.message); return; }
      evs.filter(ev => !ev.allDay).forEach(ev => {
        const start = new Date(ev.start), end = new Date(ev.end), mins = Math.round((start - now) / 60000);
        const when = fmt_(start, 'EEE d MMM, h:mm') + '–' + fmt_(end, 'h:mm a');
        if (mins > 0 && mins <= 20) notify_(u.email, 'call', 'Call in ' + mins + ' min: ' + ev.title + '\n' + when, 'call:' + ev.id + ':' + u.email);
        if (ev.created && new Date(ev.created) > last && start > now) notify_(u.email, 'newcall', 'New call: ' + ev.title + '\n' + when, 'new:' + ev.id + ':' + u.email);
        if (end <= now && now - end < 30 * 60000 && ev.attendees.length > 1) {
          const pid = 'prompt:' + ev.id + ':' + u.email;
          if (!all_('Prompts').some(x => x.id === pid)) {
            put_('Prompts', { id: pid, email: u.email, title: ev.title, endedAt: ev.end, attendees: ev.attendees.slice(0, 30).join(','), status: 'pending', createdAt: now.toISOString() });
            notify_(u.email, 'actions', '"' + ev.title + '" ended. Add its action items.' + link_(''), pid);
          }
        }
      });
    });
    PROPS.setProperty('LAST_TICK', String(now.getTime()));
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- scheduled

function mine_(t, email) { return t.spoc === email || (!t.spoc && t.createdBy === email) || split_(t.tags).indexOf(email) >= 0; }

function digestFor_(u, calls) {
  const now = new Date(), today = fmt_(now, 'yyyy-MM-dd');
  const mine = all_('Tasks').filter(t => canSee_(t, u.email) && t.status !== 'done' && mine_(t, u.email));
  const late = mine.filter(t => t.due && t.due < today);
  const due = mine.filter(t => t.due === today);
  if (!late.length && !due.length && !(calls && calls.length)) return;
  const L = ['Morning, ' + u.name.split(' ')[0] + ' · ' + fmt_(now, 'EEE d MMM')];
  if (calls && calls.length) { L.push('', 'Calls'); calls.forEach(c => L.push('• ' + String(c.time || '') + '  ' + String(c.title || ''))); }
  if (due.length) { L.push('', 'Due today'); due.forEach(t => L.push('• ' + t.title + (t.dueTime ? ' · ' + hm12_(t.dueTime) : ''))); }
  if (late.length) { L.push('', 'Overdue'); late.forEach(t => L.push('• ' + t.title)); }
  notify_(u.email, 'digest', L.join('\n') + link_(''), 'digest:' + today + ':' + u.email);
}

/** ~9:30 AM daily: today's calls (if the person linked a calendar), tasks due today, and overdue tasks. */
function morningDigest() {
  const a = new Date(); a.setHours(0, 0, 0, 0);
  const b = new Date(a.getTime() + 864e5);
  all_('Users').forEach(u => {
    let calls = null;
    if (u.gtoken || u.ical) {
      try { calls = calendarEvents_(u, a, b).filter(e => !e.allDay).map(e => ({ time: fmt_(new Date(e.start), 'h:mm a'), title: e.title })); }
      catch (e) { calls = null; }
    }
    digestFor_(u, calls);
  });
  const cut = new Date(Date.now() - 45 * 864e5).toISOString();
  drop_('Notifications', all_('Notifications').filter(n => n.createdAt < cut).map(n => n.id));
  drop_('Prompts', all_('Prompts').filter(p => p.createdAt < new Date(Date.now() - 7 * 864e5).toISOString()).map(p => p.id));
  drop_('Sessions', all_('Sessions').filter(x => Date.parse(x.expiresAt) < Date.now()).map(x => x.id));
}

/** Hourly: High priority, 24 h late → SPOC (and creator, if different), once per due date. */
function overdueSweep() {
  const now = Date.now();
  all_('Tasks').forEach(t => {
    if (t.status === 'done' || t.priority !== 'high' || !t.due || now - dueAt_(t).getTime() < 864e5) return;
    const text = 'Overdue · High: ' + t.title + '\nDue ' + fmt_(dueAt_(t), 'EEE d MMM') + link_('?task=' + t.id);
    const who = t.spoc || t.createdBy;
    notify_(who, 'overdue', text, 'late:' + t.id + ':' + t.due + ':' + who);
    if (t.category !== PERSONAL && t.createdBy !== who) notify_(t.createdBy, 'overdue', text + '\nSPOC: ' + nameOf_(who), 'late:' + t.id + ':' + t.due + ':' + t.createdBy);
  });
}

// ---------------------------------------------------------------- notifications

function notify_(email, kind, text, key) {
  if (key && all_('Notifications').some(n => n.key === key)) return;
  const u = user_(email);
  const pref = PREF_FOR[kind];
  if (pref && u && prefs_(u)[pref] === false) return;
  let status;
  try { status = sendWhatsApp_(phoneOf_(email), text); } catch (e) { status = 'failed: ' + String(e.message).slice(0, 120); }
  if (status !== 'sent' && prop_('EMAIL_FALLBACK', 'false') === 'true') {
    try { MailApp.sendEmail(email, 'TUTO Desk', text); status += ' +email'; } catch (e) {}
  }
  put_('Notifications', { id: uid_('n'), key: key || '', email: email, kind: kind, text: text, status: status, createdAt: new Date().toISOString() });
}

function sendWhatsApp_(phone, text) {
  const provider = prop_('WA_PROVIDER', 'stub');
  if (provider === 'stub') return 'logged';
  if (!phone) return 'no-phone';
  const send = PROVIDERS[provider];
  if (!send) throw new Error('Unknown WA_PROVIDER ' + provider);
  return send(String(phone).replace(/[^\d]/g, ''), text);
}

function oneLine_(t) { return t.replace(/\s*\n+\s*/g, ' · ').slice(0, 1000); }
function ok_(res) { return res.getResponseCode() < 300 ? 'sent' : 'failed: ' + res.getContentText().slice(0, 120); }

const PROVIDERS = {
  meta(to, text) {
    const tpl = prop_('WA_TEMPLATE', '');
    const body = tpl
      ? { messaging_product: 'whatsapp', to: to, type: 'template', template: { name: tpl, language: { code: prop_('WA_TEMPLATE_LANG', 'en') },
          components: [{ type: 'body', parameters: [{ type: 'text', text: oneLine_(text) }] }] } }
      : { messaging_product: 'whatsapp', to: to, type: 'text', text: { body: text } };
    return ok_(UrlFetchApp.fetch('https://graph.facebook.com/' + prop_('WA_API_VERSION', 'v21.0') + '/' + prop_('WA_PHONE_NUMBER_ID') + '/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + prop_('WA_TOKEN') }, payload: JSON.stringify(body) }));
  },
  twilio(to, text) {
    const sid = prop_('TWILIO_SID');
    return ok_(UrlFetchApp.fetch('https://api.twilio.com/2010-04-01/Accounts/' + sid + '/Messages.json', {
      method: 'post', muteHttpExceptions: true,
      headers: { Authorization: 'Basic ' + Utilities.base64Encode(sid + ':' + prop_('TWILIO_TOKEN')) },
      payload: { From: 'whatsapp:' + prop_('TWILIO_FROM'), To: 'whatsapp:+' + to, Body: text } }));
  },
  interakt(to, text) {
    const cc = prop_('WA_DEFAULT_CC', '91');
    return ok_(UrlFetchApp.fetch('https://api.interakt.ai/v1/public/message/', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Basic ' + prop_('INTERAKT_KEY') },
      payload: JSON.stringify({ countryCode: '+' + cc, phoneNumber: to.indexOf(cc) === 0 && to.length > 10 ? to.slice(cc.length) : to,
        type: 'Template', template: { name: prop_('WA_TEMPLATE'), languageCode: prop_('WA_TEMPLATE_LANG', 'en'), bodyValues: [oneLine_(text)] } }) }));
  },
};
