/**
 * TUTO Desk — Cloudflare Worker.
 *
 * Serves the app (./public), its API (/api/*), and runs the reminders:
 *   every 5 min  → call reminders, new calls, "call ended" prompts (per person, from their own Google Calendar)
 *   hourly       → High-priority tasks a day late
 *   09:30 IST    → morning digest
 * Notifications are Web Push (works on Android, desktop, and iPhone home-screen apps).
 *
 * Visibility: tasks in the "Personal" category are seen only by their creator; everything else by the whole team.
 */
import { sendPush, unb64url } from './push.js';

const PERSONAL = 'Personal';
const CATEGORIES = ['Work', 'Meetings', 'Follow-ups', PERSONAL];
const PRIORITIES = ['low', 'medium', 'high'];
const STATUSES = ['todo', 'doing', 'done'];
const SCOPES = 'openid email profile https://www.googleapis.com/auth/calendar.readonly';
const SESSION_DAYS = 120;
const DEFAULT_PREFS = { callAlerts: true, newCalls: true, digest: true, overdue: true, assign: true, tagged: true, actionItems: true };
const PREF_FOR = { call: 'callAlerts', newcall: 'newCalls', digest: 'digest', overdue: 'overdue', assign: 'assign', tagged: 'tagged', actions: 'actionItems' };

// ---------------------------------------------------------------- small helpers

class HttpError extends Error { constructor(code, msg) { super(msg); this.code = code; } }
const fail = (code, msg) => { throw new HttpError(code, msg); };
const nowIso = () => new Date().toISOString();
const uid = p => p + '_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const list = s => String(s || '').split(',').map(x => x.trim().toLowerCase()).filter(Boolean);
const json = (obj, status = 200) => new Response(JSON.stringify(obj), { status, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' } });

async function sha256(s) {
  const d = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(s));
  return [...new Uint8Array(d)].map(b => b.toString(16).padStart(2, '0')).join('');
}
function randomToken(bytes = 32) {
  return [...crypto.getRandomValues(new Uint8Array(bytes))].map(b => b.toString(16).padStart(2, '0')).join('');
}

// Local time (IST by default) for dates, the digest and "today".
const offMs = env => Number(env.TZ_OFFSET_MIN || 330) * 60000;
const localDay = (env, d = new Date()) => new Date(d.getTime() + offMs(env)).toISOString().slice(0, 10);
function localClock(env, iso) {
  const d = new Date(new Date(iso).getTime() + offMs(env));
  const h = d.getUTCHours(), m = d.getUTCMinutes();
  return `${h % 12 || 12}${m ? ':' + String(m).padStart(2, '0') : ''} ${h < 12 ? 'AM' : 'PM'}`;
}
function dueAt(env, t) {
  const [y, mo, d] = t.due.split('-').map(Number);
  const [h, mi] = t.due_time ? t.due_time.split(':').map(Number) : [23, 59];
  return Date.UTC(y, mo - 1, d, h, mi) - offMs(env);
}

// ---------------------------------------------------------------- data access

const all = async (env, sql, ...args) => (await env.DB.prepare(sql).bind(...args).all()).results || [];
const one = (env, sql, ...args) => env.DB.prepare(sql).bind(...args).first();
const run = (env, sql, ...args) => env.DB.prepare(sql).bind(...args).run();

const prefsOf = u => { let p = {}; try { p = JSON.parse(u.prefs || '{}'); } catch (e) {} return { ...DEFAULT_PREFS, ...p }; };
const admins = env => list(env.ADMIN_EMAILS);
const isAdmin = (env, email) => admins(env).includes(email);

async function team(env) { return all(env, 'SELECT email, name, picture FROM users WHERE disabled = 0 ORDER BY created_at'); }
const firstName = u => String(u.name || u.email).split(/\s+/)[0].toLowerCase().replace(/[^a-z]/g, '');

function mentions(text, people) {
  const out = [];
  for (const m of String(text || '').matchAll(/(^|\s)@([\w.+-]+)/g)) {
    const h = m[2].toLowerCase();
    const p = people.find(p => p.email.split('@')[0] === h) || people.find(p => firstName(p) === h);
    if (p && !out.includes(p.email)) out.push(p.email);
  }
  return out;
}

const taskOut = r => ({
  id: r.id, title: r.title, notes: r.notes, due: r.due, dueTime: r.due_time, priority: r.priority, category: r.category,
  spoc: r.spoc, tags: list(r.tags), status: r.status, createdBy: r.created_by, createdAt: r.created_at,
  updatedAt: r.updated_at, completedAt: r.completed_at, source: r.source,
});
const canSee = (t, email) => t.category !== PERSONAL || t.created_by === email;
const meOut = (env, u) => ({
  email: u.email, name: u.name, picture: u.picture, prefs: prefsOf(u),
  calendar: !!u.refresh_token, admin: isAdmin(env, u.email),
});

// ---------------------------------------------------------------- notifications

async function notify(env, email, kind, title, body, url, key) {
  const u = await one(env, 'SELECT * FROM users WHERE email = ? AND disabled = 0', email);
  if (!u) return;
  const pref = PREF_FOR[kind];
  if (pref && prefsOf(u)[pref] === false) return;
  const id = uid('n');
  const ins = await run(env, 'INSERT OR IGNORE INTO notifications (id, key, email, kind, title, body, url, status, created_at) VALUES (?,?,?,?,?,?,?,?,?)',
    id, key || id, email, kind, title, body || '', url || '/', '', nowIso());
  if (!ins.meta.changes) return; // already sent
  const subs = await all(env, 'SELECT * FROM push_subs WHERE email = ?', email);
  let sent = 0;
  for (const s of subs) {
    try {
      const code = await sendPush(env, s, { title, body, url: url || '/', tag: key || id });
      if (code === 404 || code === 410) await run(env, 'DELETE FROM push_subs WHERE endpoint = ?', s.endpoint);
      else if (code < 300) sent++;
    } catch (e) { console.warn('push failed', e.message); }
  }
  await run(env, 'UPDATE notifications SET status = ? WHERE id = ?', subs.length ? (sent ? 'sent' : 'failed') : 'no-device', id);
}

// ---------------------------------------------------------------- Google

async function exchangeCode(env, code) {
  if (!env.GOOGLE_CLIENT_ID || !env.GOOGLE_CLIENT_SECRET) fail('SETUP', 'Google sign-in isn\'t set up yet.');
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ code: String(code || ''), client_id: env.GOOGLE_CLIENT_ID, client_secret: env.GOOGLE_CLIENT_SECRET, redirect_uri: 'postmessage', grant_type: 'authorization_code' }),
  });
  const j = await res.json().catch(() => ({}));
  if (!res.ok || !j.id_token) fail('AUTH', 'Google sign-in failed (' + (j.error_description || j.error || res.status) + '). Try again.');
  const payload = JSON.parse(new TextDecoder().decode(unb64url(j.id_token.split('.')[1])));
  if (payload.aud !== env.GOOGLE_CLIENT_ID || String(payload.email_verified) !== 'true') fail('AUTH', 'Google sign-in failed. Try again.');
  return {
    email: String(payload.email).toLowerCase(), name: payload.name || '', picture: payload.picture || '',
    access: j.access_token || '', expiresIn: Number(j.expires_in || 3600), refresh: j.refresh_token || '',
    calendar: String(j.scope || '').includes('calendar'),
  };
}

async function saveGoogle(env, g) {
  if (!g.calendar) return;
  const exp = Date.now() + (g.expiresIn - 120) * 1000;
  if (g.refresh) await run(env, 'UPDATE users SET refresh_token = ?, access_token = ?, access_exp = ? WHERE email = ?', g.refresh, g.access, exp, g.email);
  else await run(env, 'UPDATE users SET access_token = ?, access_exp = ? WHERE email = ?', g.access, exp, g.email);
}

async function googleAccess(env, u) {
  if (u.access_token && u.access_exp > Date.now()) return u.access_token;
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: env.GOOGLE_CLIENT_ID, client_secret: env.GOOGLE_CLIENT_SECRET, refresh_token: u.refresh_token, grant_type: 'refresh_token' }),
  });
  const j = await res.json().catch(() => ({}));
  if (!res.ok) {
    if (j.error === 'invalid_grant') await run(env, 'UPDATE users SET refresh_token = \'\', access_token = \'\', access_exp = 0 WHERE email = ?', u.email);
    fail('CAL', 'Calendar access ended. Tap "Connect calendar" to reconnect.');
  }
  const exp = Date.now() + (Number(j.expires_in || 3600) - 120) * 1000;
  await run(env, 'UPDATE users SET access_token = ?, access_exp = ? WHERE email = ?', j.access_token, exp, u.email);
  return j.access_token;
}

async function calendarEvents(env, u, from, to) {
  if (!u.refresh_token) return [];
  const token = await googleAccess(env, u);
  const url = 'https://www.googleapis.com/calendar/v3/calendars/primary/events?singleEvents=true&orderBy=startTime&maxResults=250' +
    '&timeMin=' + encodeURIComponent(new Date(from).toISOString()) + '&timeMax=' + encodeURIComponent(new Date(to).toISOString());
  const res = await fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  if (!res.ok) {
    const t = await res.text();
    if (/has not been used|is disabled/i.test(t)) fail('CAL', 'The Google Calendar API is switched off in Google Cloud.');
    fail('CAL', 'Couldn\'t read the calendar (' + res.status + ').');
  }
  const day = d => new Date(Date.parse(d + 'T00:00:00Z') - offMs(env)).toISOString();
  return ((await res.json()).items || [])
    .filter(ev => ev.status !== 'cancelled' && !((ev.attendees || []).find(a => a.self) || {}).responseStatus?.startsWith('declined'))
    .map(ev => {
      const allDay = !ev.start.dateTime;
      return {
        id: ev.id, title: ev.summary || 'Busy',
        start: allDay ? day(ev.start.date) : new Date(ev.start.dateTime).toISOString(),
        end: allDay ? day(ev.end.date) : new Date(ev.end.dateTime).toISOString(),
        allDay, attendees: (ev.attendees || []).map(a => String(a.email || '').toLowerCase()),
        location: ev.hangoutLink ? 'Google Meet' : (ev.location || ''), created: ev.created || '',
      };
    });
}

// ---------------------------------------------------------------- sessions

async function newSession(env, email, device) {
  const token = 'sess_' + randomToken();
  await run(env, 'INSERT INTO sessions (id, email, device, expires_at) VALUES (?,?,?,?)', await sha256(token), email, String(device || '').slice(0, 60), Date.now() + SESSION_DAYS * 864e5);
  return token;
}

async function authed(env, req) {
  const token = (req.headers.get('Authorization') || '').replace(/^Bearer\s+/i, '');
  if (!token.startsWith('sess_')) fail('AUTH', 'Sign in to continue.');
  const id = await sha256(token);
  const s = await one(env, 'SELECT * FROM sessions WHERE id = ?', id);
  if (!s || s.expires_at < Date.now()) fail('AUTH', 'Your session ended. Sign in again.');
  const u = await one(env, 'SELECT * FROM users WHERE email = ?', s.email);
  if (!u || u.disabled) fail('AUTH', 'Your access was removed.');
  if (s.expires_at - Date.now() < 60 * 864e5) await run(env, 'UPDATE sessions SET expires_at = ? WHERE id = ?', Date.now() + SESSION_DAYS * 864e5, id);
  if (!u.last_seen || Date.now() - Date.parse(u.last_seen) > 10 * 60000) await run(env, 'UPDATE users SET last_seen = ? WHERE email = ?', nowIso(), u.email);
  return { user: u, sessionId: id };
}

async function signIn(env, g, invite, device) {
  let u = await one(env, 'SELECT * FROM users WHERE email = ?', g.email);
  if (u && u.disabled) fail('FORBIDDEN', 'Your access was removed. Ask Veronica.');
  if (!u) {
    const domains = list(env.ALLOWED_DOMAINS);
    const ok = isAdmin(env, g.email) || (env.INVITE_CODE && invite === env.INVITE_CODE) || domains.includes(g.email.split('@')[1]);
    if (!ok) fail('FORBIDDEN', 'You need an invite link to join. Ask Veronica for it.');
    await run(env, 'INSERT INTO users (email, name, picture, prefs, created_at, last_seen) VALUES (?,?,?,?,?,?)', g.email, g.name || g.email.split('@')[0], g.picture, '{}', nowIso(), nowIso());
  } else if (g.picture && g.picture !== u.picture) {
    await run(env, 'UPDATE users SET picture = ? WHERE email = ?', g.picture, g.email);
  }
  await saveGoogle(env, g);
  return { session: await newSession(env, g.email, device) };
}

// ---------------------------------------------------------------- tasks

async function saveTask(env, u, t, ppl) {
  const prev = t.id ? await one(env, 'SELECT * FROM tasks WHERE id = ?', t.id) : null;
  if (prev && !canSee(prev, u.email)) fail('FORBIDDEN', 'Task not found.');
  const title = String(t.title || '').trim().slice(0, 300);
  if (!title) fail('BAD_REQUEST', 'Add a title.');
  const category = CATEGORIES.includes(t.category) ? t.category : 'Work';
  const owner = prev ? prev.created_by : u.email;
  const personal = category === PERSONAL;
  if (personal && owner !== u.email) fail('FORBIDDEN', 'Only the creator can make a task Personal.');
  const people = ppl || await team(env);
  let spoc = personal ? '' : String(t.spoc || '').trim().toLowerCase();
  if (spoc && !people.some(p => p.email === spoc)) fail('BAD_REQUEST', spoc + ' isn\'t on the team.');
  const notes = String(t.notes || '').slice(0, 4000);
  const tags = personal ? [] : mentions(title + ' ' + notes, people);
  const now = nowIso();
  const status = STATUSES.includes(t.status) ? t.status : 'todo';
  const rec = {
    id: prev ? prev.id : (/^t_[a-z0-9]{4,24}$/.test(t.id || '') ? t.id : uid('t')),
    title, notes,
    due: /^\d{4}-\d{2}-\d{2}$/.test(t.due || '') ? t.due : '',
    due_time: /^\d{2}:\d{2}$/.test(t.dueTime || '') ? t.dueTime : '',
    priority: PRIORITIES.includes(t.priority) ? t.priority : 'medium',
    category, spoc, tags: tags.join(','), status,
    created_by: owner, created_at: prev ? prev.created_at : now, updated_at: now,
    completed_at: status === 'done' ? ((prev && prev.completed_at) || now) : '',
    source: prev ? prev.source : String(t.source || '').slice(0, 40),
  };
  await run(env, `INSERT INTO tasks (id, title, notes, due, due_time, priority, category, spoc, tags, status, created_by, created_at, updated_at, completed_at, source)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    ON CONFLICT(id) DO UPDATE SET title=excluded.title, notes=excluded.notes, due=excluded.due, due_time=excluded.due_time, priority=excluded.priority,
      category=excluded.category, spoc=excluded.spoc, tags=excluded.tags, status=excluded.status, updated_at=excluded.updated_at, completed_at=excluded.completed_at`,
    rec.id, rec.title, rec.notes, rec.due, rec.due_time, rec.priority, rec.category, rec.spoc, rec.tags, rec.status, rec.created_by, rec.created_at, rec.updated_at, rec.completed_at, rec.source);

  if (!personal) {
    const who = u.name.split(' ')[0];
    const url = '/?task=' + rec.id;
    if (spoc && spoc !== u.email && (!prev || prev.spoc !== spoc)) {
      await notify(env, spoc, 'assign', `${who} made you SPOC`, rec.title, url, `assign:${rec.id}:${spoc}:${now}`);
    }
    const before = prev ? list(prev.tags) : [];
    for (const e of tags.filter(e => e !== u.email && e !== spoc && !before.includes(e))) {
      await notify(env, e, 'tagged', `${who} tagged you`, rec.title, url, `tag:${rec.id}:${e}`);
    }
  }
  return taskOut(rec);
}

// ---------------------------------------------------------------- API

const PUBLIC = {
  async config(env) {
    return { clientId: env.GOOGLE_CLIENT_ID || '', scopes: SCOPES, vapidPublic: env.VAPID_PUBLIC || '', ready: !!(env.GOOGLE_CLIENT_ID && env.GOOGLE_CLIENT_SECRET) };
  },
  async signIn(env, p) {
    const g = await exchangeCode(env, p.code);
    return signIn(env, g, String(p.invite || ''), p.device);
  },
  async devSignIn(env, p) {
    if (env.DEV_LOGIN !== '1') fail('NOT_FOUND', 'Unknown action.');
    const email = String(p.email || '').toLowerCase();
    return signIn(env, { email, name: p.name || email.split('@')[0], picture: '', calendar: false }, String(p.invite || ''), 'dev');
  },
};

const API = {
  async bootstrap(env, u) {
    const people = await team(env);
    const tasks = (await all(env, 'SELECT * FROM tasks WHERE category != ? OR created_by = ? ORDER BY created_at DESC', PERSONAL, u.email));
    const ids = new Set(tasks.map(t => t.id));
    const comments = (await all(env, 'SELECT * FROM comments ORDER BY created_at')).filter(c => ids.has(c.task_id))
      .map(c => ({ id: c.id, taskId: c.task_id, author: c.author, text: c.text, createdAt: c.created_at }));
    const prompts = (await all(env, 'SELECT * FROM prompts WHERE email = ? AND status = ? ORDER BY created_at DESC', u.email, 'pending'))
      .map(p => ({ id: p.id, title: p.title, endedAt: p.ended_at, attendees: list(p.attendees) }));
    const notifications = (await all(env, 'SELECT * FROM notifications WHERE email = ? ORDER BY created_at DESC LIMIT 40', u.email))
      .map(n => ({ id: n.id, kind: n.kind, email: n.email, text: n.title + (n.body ? '\n' + n.body : ''), status: n.status, url: n.url, createdAt: n.created_at }));
    const devices = (await one(env, 'SELECT COUNT(*) AS n FROM push_subs WHERE email = ?', u.email)).n;
    return {
      me: { ...meOut(env, u), devices },
      users: people.map(p => ({ email: p.email, name: p.name, picture: p.picture, joined: true })),
      tasks: tasks.map(taskOut), comments, prompts, notifications, categories: CATEGORIES,
    };
  },

  async saveTask(env, u, p) { return saveTask(env, u, p.task || {}); },

  async saveTasks(env, u, p) {
    const people = await team(env);
    const out = [];
    for (const t of (p.tasks || []).slice(0, 50)) out.push(await saveTask(env, u, t, people));
    if (p.promptId) await run(env, 'UPDATE prompts SET status = ? WHERE id = ? AND email = ?', 'done', p.promptId, u.email);
    return out;
  },

  async deleteTask(env, u, p) {
    const t = await one(env, 'SELECT * FROM tasks WHERE id = ?', p.id);
    if (!t) return { id: p.id };
    if (!canSee(t, u.email)) fail('FORBIDDEN', 'Task not found.');
    await env.DB.batch([env.DB.prepare('DELETE FROM tasks WHERE id = ?').bind(t.id), env.DB.prepare('DELETE FROM comments WHERE task_id = ?').bind(t.id)]);
    return { id: t.id };
  },

  async addComment(env, u, p) {
    const t = await one(env, 'SELECT * FROM tasks WHERE id = ?', p.taskId);
    if (!t || !canSee(t, u.email)) fail('NOT_FOUND', 'Task not found.');
    const text = String(p.text || '').trim().slice(0, 2000);
    if (!text) fail('BAD_REQUEST', 'Write something first.');
    const c = { id: uid('c'), taskId: t.id, author: u.email, text, createdAt: nowIso() };
    await run(env, 'INSERT INTO comments (id, task_id, author, text, created_at) VALUES (?,?,?,?,?)', c.id, c.taskId, c.author, c.text, c.createdAt);
    if (t.category !== PERSONAL) {
      for (const e of mentions(text, await team(env)).filter(e => e !== u.email)) {
        await notify(env, e, 'tagged', `${u.name.split(' ')[0]} mentioned you`, `"${t.title}": ${text}`, '/?task=' + t.id, `ctag:${c.id}:${e}`);
      }
    }
    return c;
  },

  async saveProfile(env, u, p) {
    const prefs = prefsOf(u);
    for (const k of Object.keys(DEFAULT_PREFS)) if (p.prefs && typeof p.prefs[k] === 'boolean') prefs[k] = p.prefs[k];
    const name = String(p.name || u.name).trim().slice(0, 60) || u.name;
    await run(env, 'UPDATE users SET name = ?, prefs = ? WHERE email = ?', name, JSON.stringify(prefs), u.email);
    return meOut(env, { ...u, name, prefs: JSON.stringify(prefs) });
  },

  async calendar(env, u, p) {
    const from = Date.parse(p.from), to = Date.parse(p.to);
    if (!from || !to || to <= from || to - from > 43 * 864e5) fail('BAD_REQUEST', 'Bad date range.');
    return calendarEvents(env, u, from, to);
  },

  async connectCalendar(env, u, p) {
    const g = await exchangeCode(env, p.code);
    if (g.email !== u.email) fail('BAD_REQUEST', 'Pick ' + u.email + ' in the Google window.');
    if (!g.calendar) fail('BAD_REQUEST', 'Tick the calendar box in the Google window to show your calls.');
    await saveGoogle(env, g);
    return meOut(env, await one(env, 'SELECT * FROM users WHERE email = ?', u.email));
  },

  async disconnectCalendar(env, u) {
    if (u.refresh_token) await fetch('https://oauth2.googleapis.com/revoke?token=' + encodeURIComponent(u.refresh_token), { method: 'POST' }).catch(() => {});
    await run(env, 'UPDATE users SET refresh_token = \'\', access_token = \'\', access_exp = 0 WHERE email = ?', u.email);
    return meOut(env, { ...u, refresh_token: '' });
  },

  async pushSubscribe(env, u, p) {
    const s = p.subscription || {};
    if (!/^https:\/\//.test(s.endpoint || '') || !s.keys || !s.keys.p256dh || !s.keys.auth) fail('BAD_REQUEST', 'This browser didn\'t return a valid subscription.');
    await run(env, `INSERT INTO push_subs (endpoint, email, p256dh, auth, device, created_at) VALUES (?,?,?,?,?,?)
      ON CONFLICT(endpoint) DO UPDATE SET email=excluded.email, p256dh=excluded.p256dh, auth=excluded.auth, device=excluded.device`,
      s.endpoint, u.email, s.keys.p256dh, s.keys.auth, String(p.device || '').slice(0, 60), nowIso());
    return { ok: true };
  },

  async pushUnsubscribe(env, u, p) {
    await run(env, 'DELETE FROM push_subs WHERE endpoint = ? AND email = ?', String(p.endpoint || ''), u.email);
    return { ok: true };
  },

  async testNotify(env, u) {
    const n = (await one(env, 'SELECT COUNT(*) AS n FROM push_subs WHERE email = ?', u.email)).n;
    await notify(env, u.email, 'test', 'TUTO Desk test', 'Notifications are working.', '/', null);
    return { devices: n };
  },

  async resolvePrompt(env, u, p) {
    await run(env, 'UPDATE prompts SET status = ? WHERE id = ? AND email = ?', p.status === 'done' ? 'done' : 'dismissed', String(p.id || ''), u.email);
    return { id: p.id };
  },

  async signOut(env, u, p, ctx) {
    await run(env, 'DELETE FROM sessions WHERE id = ?', ctx.sessionId);
    if (p.endpoint) await run(env, 'DELETE FROM push_subs WHERE endpoint = ?', String(p.endpoint));
    return { ok: true };
  },

  // ---- admin
  async invite(env, u, p, ctx) {
    if (!isAdmin(env, u.email)) fail('FORBIDDEN', 'Only admins can invite.');
    const members = await all(env, 'SELECT email, name, last_seen, disabled, refresh_token != \'\' AS calendar FROM users ORDER BY created_at');
    const devices = await all(env, 'SELECT email, COUNT(*) AS n FROM push_subs GROUP BY email');
    return {
      link: env.INVITE_CODE ? `${ctx.origin}/?invite=${env.INVITE_CODE}` : '',
      members: members.map(m => ({ email: m.email, name: m.name, lastSeen: m.last_seen, removed: !!m.disabled, calendar: !!m.calendar, devices: (devices.find(d => d.email === m.email) || {}).n || 0 })),
    };
  },

  async removeMember(env, u, p) {
    if (!isAdmin(env, u.email)) fail('FORBIDDEN', 'Only admins can remove people.');
    const email = String(p.email || '').toLowerCase();
    if (email === u.email) fail('BAD_REQUEST', 'You can\'t remove yourself.');
    await env.DB.batch([
      env.DB.prepare('UPDATE users SET disabled = ?, refresh_token = \'\' WHERE email = ?').bind(p.restore ? 0 : 1, email),
      env.DB.prepare('DELETE FROM sessions WHERE email = ?').bind(email),
      env.DB.prepare('DELETE FROM push_subs WHERE email = ?').bind(email),
    ]);
    return { ok: true };
  },
};

// ---------------------------------------------------------------- scheduled jobs

async function callsTick(env) {
  const now = Date.now();
  const last = Number((await one(env, 'SELECT value FROM meta WHERE key = ?', 'last_tick') || {}).value || now - 5 * 60000);
  const users = await all(env, 'SELECT * FROM users WHERE disabled = 0 AND refresh_token != \'\'');
  for (const u of users) {
    let evs;
    try { evs = await calendarEvents(env, u, now - 40 * 60000, now + 7 * 864e5); }
    catch (e) { console.warn(u.email, e.message); continue; }
    for (const ev of evs.filter(e => !e.allDay)) {
      const start = Date.parse(ev.start), end = Date.parse(ev.end), mins = Math.round((start - now) / 60000);
      const key = ev.id + '@' + start;
      if (mins > 0 && mins <= 15) {
        await notify(env, u.email, 'call', `Call in ${mins} min`, `${ev.title} · ${localClock(env, ev.start)}${ev.location ? ' · ' + ev.location : ''}`, '/', `call:${key}:${u.email}`);
      }
      if (ev.created && Date.parse(ev.created) > last && start > now) {
        await notify(env, u.email, 'newcall', 'New call booked', `${ev.title} · ${localDay(env, new Date(start)).slice(5)} ${localClock(env, ev.start)}`, '/', `new:${key}:${u.email}`);
      }
      if (end <= now && now - end < 30 * 60000 && ev.attendees.length > 1) {
        const pid = `prompt:${key}:${u.email}`;
        const ins = await run(env, 'INSERT OR IGNORE INTO prompts (id, email, title, ended_at, attendees, status, created_at) VALUES (?,?,?,?,?,?,?)',
          pid, u.email, ev.title, ev.end, ev.attendees.slice(0, 30).join(','), 'pending', nowIso());
        if (ins.meta.changes) await notify(env, u.email, 'actions', `"${ev.title}" ended`, 'Add its action items while they\'re fresh.', '/', pid);
      }
    }
  }
  await run(env, 'INSERT INTO meta (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value', 'last_tick', String(now));
}

async function overdueSweep(env) {
  const now = Date.now();
  const tasks = await all(env, 'SELECT * FROM tasks WHERE status != ? AND priority = ? AND due != ?', 'done', 'high', '');
  for (const t of tasks) {
    if (now - dueAt(env, t) < 864e5) continue;
    const who = t.spoc || t.created_by;
    await notify(env, who, 'overdue', 'Overdue · High priority', t.title, '/?task=' + t.id, `late:${t.id}:${t.due}:${who}`);
    if (t.category !== PERSONAL && t.created_by !== who) await notify(env, t.created_by, 'overdue', 'Overdue · High priority', t.title, '/?task=' + t.id, `late:${t.id}:${t.due}:${t.created_by}`);
  }
  const cut = new Date(now - 45 * 864e5).toISOString();
  await env.DB.batch([
    env.DB.prepare('DELETE FROM notifications WHERE created_at < ?').bind(cut),
    env.DB.prepare('DELETE FROM prompts WHERE created_at < ?').bind(new Date(now - 7 * 864e5).toISOString()),
    env.DB.prepare('DELETE FROM sessions WHERE expires_at < ?').bind(now),
  ]);
}

async function morningDigest(env) {
  const today = localDay(env);
  const users = await all(env, 'SELECT * FROM users WHERE disabled = 0');
  const tasks = await all(env, 'SELECT * FROM tasks WHERE status != ?', 'done');
  const dayStart = Date.parse(today + 'T00:00:00Z') - offMs(env);
  for (const u of users) {
    const mine = tasks.filter(t => canSee(t, u.email) && (t.spoc === u.email || (!t.spoc && t.created_by === u.email) || list(t.tags).includes(u.email)));
    const due = mine.filter(t => t.due === today).length, late = mine.filter(t => t.due && t.due < today).length;
    let calls = 0;
    try { calls = (await calendarEvents(env, u, dayStart, dayStart + 864e5)).filter(e => !e.allDay).length; } catch (e) {}
    if (!due && !late && !calls) continue;
    const bits = [calls && `${calls} call${calls > 1 ? 's' : ''}`, due && `${due} due today`, late && `${late} overdue`].filter(Boolean);
    await notify(env, u.email, 'digest', `Morning, ${u.name.split(' ')[0]}`, bits.join(' · '), '/', `digest:${today}:${u.email}`);
  }
}

// ---------------------------------------------------------------- entry points

export default {
  async fetch(req, env, ctx) {
    const url = new URL(req.url);
    if (!url.pathname.startsWith('/api/')) return env.ASSETS.fetch(req);
    if (req.method !== 'POST') return json({ ok: false, error: 'Use POST.' }, 405);
    const action = url.pathname.slice(5);
    try {
      const p = await req.json().catch(() => ({}));
      if (PUBLIC[action]) return json({ ok: true, data: await PUBLIC[action](env, p) });
      const fn = API[action];
      if (!fn) fail('NOT_FOUND', 'Unknown action.');
      const a = await authed(env, req);
      return json({ ok: true, data: await fn(env, a.user, p, { sessionId: a.sessionId, origin: url.origin }) });
    } catch (e) {
      if (!(e instanceof HttpError)) console.error(e);
      return json({ ok: false, code: e.code || 'ERROR', error: e instanceof HttpError ? e.message : 'Something went wrong. Try again.' });
    }
  },

  async scheduled(event, env, ctx) {
    if (event.cron === '0 4 * * *') { ctx.waitUntil(morningDigest(env)); return; }
    ctx.waitUntil((async () => {
      await callsTick(env);
      if (new Date(event.scheduledTime).getUTCMinutes() < 5) await overdueSweep(env);
    })());
  },
};
