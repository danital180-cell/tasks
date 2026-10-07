// Daniel Assistant — משימות + דוח שבועי (קובץ רביעי באותו פרויקט Apps Script)
// המשימות נשמרות בגיליון גוגל "משימות – העוזר האישי" (נוצר אוטומטית בדרייב).
// דף הסימון: כתובת ה-Web App + ?p=tasks&k=<TASKS_KEY>  (המפתח נוצר אוטומטית ונשמר במאפייני הסקריפט)

const TASKS_SHEET_NAME = 'משימות – העוזר האישי';
const TASK_COLS = ['id', 'created', 'title', 'area', 'contact', 'phone', 'owner', 'due', 'status',
                   'note', 'source', 'weeks', 'updated', 'doneAt', 'ref', 'time'];
const TASK_AREAS = ['הרצל בוטיק', 'פסגות העמק', 'טאבונוצי', 'RETAIN', 'אישי', 'אחר'];
const TASK_STATUSES = ['open', 'progress', 'done', 'postponed', 'cancelled'];
// כתובת קבועה של דף המשימות (GitHub Pages). המפתח מתווסף אחרי # ולא נשלח לשום שרת.
const TASKS_PUBLIC_URL = 'https://danital180-cell.github.io/tasks/';

// ---------- גיליון ----------
function tasksSheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('TASKS_SHEET');
  let ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create(TASKS_SHEET_NAME);
    props.setProperty('TASKS_SHEET', ss.getId());
    const sh0 = ss.getSheets()[0];
    sh0.setName('משימות');
    sh0.getRange(1, 1, sh0.getMaxRows(), TASK_COLS.length).setNumberFormat('@'); // טקסט: שומר 0 בטלפונים
    sh0.getRange(1, 1, 1, TASK_COLS.length).setValues([TASK_COLS]).setFontWeight('bold');
    sh0.setFrozenRows(1);
    sh0.setRightToLeft(true);
  }
  const sh = ss.getSheetByName('משימות') || ss.getSheets()[0];
  const hc = sh.getRange(1, TASK_COLS.length); // עמודת שעה (time) – נוספה מאוחר יותר
  if (hc.getValue() !== TASK_COLS[TASK_COLS.length - 1]) hc.setValue(TASK_COLS[TASK_COLS.length - 1]).setFontWeight('bold');
  return sh;
}
function tasksKey_() {
  const props = PropertiesService.getScriptProperties();
  let k = props.getProperty('TASKS_KEY');
  if (!k) { k = Utilities.getUuid().replace(/-/g, '').slice(0, 20); props.setProperty('TASKS_KEY', k); }
  return k;
}
function tasksAll_() {
  const sh = tasksSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const rows = sh.getRange(2, 1, n, TASK_COLS.length).getValues();
  return rows.map(function (r, i) {
    const t = { row: i + 2 };
    TASK_COLS.forEach(function (c, j) {
      let v = r[j];
      if (c === 'due') v = dueStr_(v);
      else if (v instanceof Date) v = v.toISOString();
      t[c] = v === '' ? '' : v;
    });
    t.weeks = Number(t.weeks) || 0;
    return t;
  });
}
// תאריך יעד תמיד כטקסט YYYY-MM-DD בשעון ישראל (תיקון: תאריך שנשמר כ-ISO הוזז ביום)
function dueStr_(v) {
  if (v === '' || v === null || v === undefined) return '';
  const tz = 'Asia/Jerusalem';
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const s = String(v);
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) { const d = new Date(s); if (!isNaN(d)) return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); }
  return s.slice(0, 10);
}
function taskWrite_(t) {
  const sh = tasksSheet_();
  t.due = dueStr_(t.due);
  const vals = TASK_COLS.map(function (c) { return t[c] === undefined ? '' : t[c]; });
  const row = t.row || sh.getLastRow() + 1;
  const rng = sh.getRange(row, 1, 1, TASK_COLS.length);
  rng.setNumberFormat('@'); // טקסט – כדי שגוגל לא יהפוך תאריכים/טלפונים
  rng.setValues([vals]);
  if (!t.row) t.row = row;
}
function weekStart_(d) { // יום ראשון 00:00 של השבוע (שעון ישראל)
  const tz = 'Asia/Jerusalem';
  const day = Number(Utilities.formatDate(d, tz, 'u')) % 7; // 1=שני...7=ראשון → 0=ראשון
  const s = new Date(d.getTime() - day * 86400000);
  return new Date(Utilities.formatDate(s, tz, 'yyyy-MM-dd') + 'T00:00:00' + Utilities.formatDate(s, tz, 'XXX'));
}

// ---------- פעולות לבוט ----------
// השוואת כותרות: בלי רווחים כפולים, סימני פיסוק ו-✅/❌
function normTitle_(s) {
  return String(s || '').replace(/^[✅❌]\s*/, '').replace(/[\s\-–—_.,:;!?"'״׳()]+/g, ' ').trim().toLowerCase();
}
function isActive_(t) { return t.status !== 'done' && t.status !== 'cancelled'; }

function addTask_(p) {
  p = p || {};
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {}
  try {
    return addTaskLocked_(p);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
function addTaskLocked_(p) {
  const all = tasksAll_();
  if (p.ref) { // מניעת כפילות (למשל אירוע יומן שכבר נוסף)
    const ex = all.filter(function (t) { return t.ref && t.ref === p.ref; })[0];
    if (ex) return { task: ex, duplicate: true };
  }
  // מניעת כפילות לפי כותרת: כבר יש משימה פעילה עם אותה כותרת
  const nt = normTitle_(p.title);
  if (nt) {
    const same = all.filter(function (t) { return isActive_(t) && normTitle_(t.title) === nt; })[0];
    if (same) {
      // אם זה אירוע יומן ולמשימה הקיימת אין קישור – מחברים ביניהם
      if (p.ref && !same.ref) {
        same.ref = p.ref;
        if (p.due) same.due = p.due;
        same.updated = new Date().toISOString();
        taskWrite_(same);
      }
      return { task: same, duplicate: true };
    }
  }
  const now = new Date().toISOString();
  const t = {
    id: 't' + Utilities.getUuid().replace(/-/g, '').slice(0, 8),
    created: now,
    title: String(p.title || '').slice(0, 200),
    area: TASK_AREAS.indexOf(p.area) >= 0 ? p.area : 'אחר',
    contact: p.contact || '', phone: p.phone || '', owner: p.owner || '',
    due: p.due || '', time: p.time || '',
    status: TASK_STATUSES.indexOf(p.status) >= 0 ? p.status : 'open',
    note: p.note || '', source: p.source || 'daniel',
    weeks: 0, updated: now, doneAt: p.status === 'done' ? now : '', ref: p.ref || '',
  };
  if (!t.title) return { error: 'empty title' };
  taskWrite_(t);
  return { task: t };
}

// משימות פתוחות (אופציונלי: לפי איש קשר / אחראי)
function listTasks_(p) {
  p = p || {};
  syncFromCalendar_();
  const q = String(p.query || '').trim();
  let list = tasksAll_().filter(function (t) { return ['open', 'progress', 'postponed'].indexOf(t.status) >= 0; });
  if (q) list = list.filter(function (t) { return (t.contact + ' ' + t.owner + ' ' + t.title + ' ' + t.area).indexOf(q) >= 0; });
  // מיון לפי תאריך ושעה (בלי תאריך – בסוף), כדי שה-30 שמוצגות יהיו המוקדמות ביותר
  list.sort(function (a, b) {
    const ka = (String(a.due || '').slice(0, 10) || '9999-99-99') + ' ' + (a.time || '99:99');
    const kb = (String(b.due || '').slice(0, 10) || '9999-99-99') + ' ' + (b.time || '99:99');
    return ka < kb ? -1 : ka > kb ? 1 : 0;
  });
  // limit: ברירת מחדל 30 (תצוגה בוואטסאפ). בדיקות פנימיות מבקשות מספר גדול יותר.
  const limit = Math.max(1, Math.min(Number(p.limit) || 30, 500));
  return { items: list.slice(0, limit), total: list.length, url: tasksPageUrl_(p.base) };
}

// סימון משימה כבוצעה מהוואטסאפ ("סיימתי עם הספק"): מחפש במשימות הפתוחות
function doneTask_(p) {
  p = p || {};
  const words = String(p.query || '').trim().split(/\s+/).filter(function (w) { return w.length > 1; });
  if (!words.length) return { matches: [] };
  const open = listTasks_({ limit: 500 }).items;
  const hit = open.filter(function (t) {
    const hay = t.title + ' ' + t.contact + ' ' + t.owner;
    return words.every(function (w) { return hay.indexOf(w) >= 0; });
  });
  if (hit.length !== 1) return { matches: hit.slice(0, 5) };
  const t = hit[0];
  t.status = p.status && TASK_STATUSES.indexOf(p.status) >= 0 ? p.status : 'done';
  t.updated = new Date().toISOString();
  t.doneAt = t.status === 'done' ? t.updated : '';
  taskWrite_(t);
  syncCalendar_(t);
  return { task: t };
}

// סיכום לשבוע הנוכחי + קישור לדף
function weeklySummary_(p) {
  p = p || {};
  syncFromCalendar_();
  const ws = weekStart_(new Date()).toISOString();
  const all = tasksAll_();
  const week = all.filter(function (t) {
    const active = ['open', 'progress', 'postponed'].indexOf(t.status) >= 0;
    return active || t.created >= ws || (t.doneAt && t.doneAt >= ws) || t.updated >= ws;
  });
  const c = { total: week.length, done: 0, progress: 0, open: 0, postponed: 0, cancelled: 0, stuck: 0, waiting: 0 };
  week.forEach(function (t) {
    c[t.status] = (c[t.status] || 0) + 1;
    if (t.weeks >= 2 && t.status !== 'done' && t.status !== 'cancelled') c.stuck++;
    if (t.owner && t.status !== 'done' && t.status !== 'cancelled') c.waiting++;
  });
  return { counts: c, url: tasksPageUrl_(p.base) };
}

function tasksPageUrl_(base) {
  if (TASKS_PUBLIC_URL) return TASKS_PUBLIC_URL + '#k=' + tasksKey_();
  base = base || ScriptApp.getService().getUrl() || '';
  return base + '?p=tasks&k=' + tasksKey_();
}

// סגירת שבוע (יום ראשון בבוקר): משימות שלא הסתיימו עוברות לשבוע הבא ומונה הדחיות עולה
function closeWeek_() {
  const ws = weekStart_(new Date()).toISOString();
  let moved = 0;
  tasksAll_().forEach(function (t) {
    if (['open', 'progress', 'postponed'].indexOf(t.status) >= 0 && t.created < ws) {
      t.weeks = (Number(t.weeks) || 0) + 1;
      if (t.status === 'postponed') t.status = 'open';
      t.updated = new Date().toISOString();
      taskWrite_(t);
      moved++;
    }
  });
  const open = listTasks_({ limit: 500 }).items;
  return { moved: moved, open: open };
}

// ---------- דף הסימון (Web App: doGet) ----------
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.p === 'api') return tasksApi_(p); // לדף הקבוע ב-GitHub Pages
  if (p.p !== 'tasks' || p.k !== tasksKey_()) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">אין גישה.</p>');
  }
  const t = HtmlService.createTemplate(TASKS_PAGE_HTML);
  t.key = p.k;
  return t.evaluate()
    .setTitle('סגירת שבוע – משימות')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// נקרא מהדף (google.script.run) — בודק מפתח
function tasksPageData(key) {
  if (key !== tasksKey_()) throw new Error('denied');
  syncFromCalendar_();
  const ws = weekStart_(new Date()).toISOString();
  const items = tasksAll_().filter(function (t) {
    return ['open', 'progress', 'postponed'].indexOf(t.status) >= 0 || t.updated >= ws || t.created >= ws;
  });
  return { items: items, areas: TASK_AREAS };
}
function tasksPageUpdate(key, id, status, note, area, title) {
  if (key !== tasksKey_()) throw new Error('denied');
  if (TASK_STATUSES.indexOf(status) < 0) throw new Error('bad status');
  const t = tasksAll_().filter(function (x) { return x.id === id; })[0];
  if (!t) throw new Error('not found');
  t.status = status;
  if (note !== undefined && note !== null) t.note = String(note).slice(0, 500);
  if (area && TASK_AREAS.indexOf(area) >= 0) t.area = area;
  if (title && String(title).trim()) t.title = String(title).trim().slice(0, 200);
  t.updated = new Date().toISOString();
  t.doneAt = status === 'done' ? (t.doneAt || t.updated) : '';
  taskWrite_(t);
  syncCalendar_(t);
  return { ok: true };
}
function tasksPageAdd(key, title, area, date, time) {
  if (key !== tasksKey_()) throw new Error('denied');
  const r = addTask_({ title: title, area: area, source: 'page' });
  if (r.task && !r.duplicate && date) tasksPageWhen(key, r.task.id, date, time);
  return r;
}

// תאריך + שעה מהדף → נשמר במשימה, והאירוע ביומן נוצר / זז בהתאם
function tasksPageWhen(key, id, date, time) {
  if (key !== tasksKey_()) throw new Error('denied');
  date = String(date || '').trim(); time = String(time || '').trim();
  if (date && !/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('bad date');
  if (time && !/^\d{2}:\d{2}$/.test(time)) throw new Error('bad time');
  if (time && !date) date = Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyyy-MM-dd');
  const t = tasksAll_().filter(function (x) { return x.id === id; })[0];
  if (!t) throw new Error('not found');
  t.due = date; t.time = date ? time : '';
  t.updated = new Date().toISOString();
  if (date) {
    try { calendarWhen_(t); } catch (e) { console.log('calendar when failed', e); t.calError = String(e.message || e); }
  }
  taskWrite_(t);
  return { ok: true, ref: t.ref, error: t.calError };
}
function calendarWhen_(t) {
  const tz = 'Asia/Jerusalem';
  const ref = String(t.ref || '');
  let ev = null;
  if (ref.indexOf('cal:') === 0) {
    try { ev = Calendar.Events.get(TASKS_CALENDAR_ID, ref.slice(4)); } catch (e) { ev = null; }
    if (ev && ev.status === 'cancelled') ev = null;
  }
  let start, end;
  if (t.time) {
    let mins = 30; // משך ברירת מחדל, או המשך הקיים של האירוע
    if (ev && ev.start && ev.start.dateTime && ev.end && ev.end.dateTime) {
      mins = Math.max(5, Math.round((new Date(ev.end.dateTime) - new Date(ev.start.dateTime)) / 60000));
    }
    const off = Utilities.formatDate(new Date(t.due + 'T12:00:00Z'), tz, 'XXX');
    const s0 = new Date(t.due + 'T' + t.time + ':00' + off);
    const e0 = new Date(s0.getTime() + mins * 60000);
    start = { dateTime: Utilities.formatDate(s0, tz, "yyyy-MM-dd'T'HH:mm:ssXXX"), timeZone: tz };
    end = { dateTime: Utilities.formatDate(e0, tz, "yyyy-MM-dd'T'HH:mm:ssXXX"), timeZone: tz };
  } else { // בלי שעה → אירוע של יום שלם
    const n = new Date(t.due + 'T12:00:00Z'); n.setUTCDate(n.getUTCDate() + 1);
    start = { date: t.due };
    end = { date: Utilities.formatDate(n, 'UTC', 'yyyy-MM-dd') };
  }
  const prefix = t.status === 'done' ? '✅ ' : t.status === 'cancelled' ? '❌ ' : '';
  if (ev) {
    ev.start = start; ev.end = end; ev.summary = prefix + t.title;
    if (prefix) ev.reminders = { useDefault: false, overrides: [] };
    Calendar.Events.update(ev, TASKS_CALENDAR_ID, ev.id);
  } else {
    const made = Calendar.Events.insert({ summary: prefix + t.title, start: start, end: end,
      description: 'נוצר מדף המשימות' }, TASKS_CALENDAR_ID);
    t.ref = 'cal:' + made.id;
  }
}

// API לדף הקבוע (GET, JSON). כל קריאה חייבת את מפתח הדף.
function tasksApi_(p) {
  let out;
  try {
    if (p.k !== tasksKey_()) throw new Error('denied');
    if (p.a === 'list') out = tasksPageData(p.k);
    else if (p.a === 'update') out = tasksPageUpdate(p.k, p.id, p.status, p.note, p.area, p.title);
    else if (p.a === 'add') out = tasksPageAdd(p.k, p.title, p.area, p.date, p.time);
    else if (p.a === 'when') out = tasksPageWhen(p.k, p.id, p.date, p.time);
    else throw new Error('bad action');
  } catch (err) {
    out = { error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// משימה שנוצרה מהיומן → האירוע ביומן גוגל מתעדכן: שם, ובוצע ✅ / בוטל ❌ בתחילת השם
// (הבוט לא שולח תזכורת לאירוע שמסומן ✅ או ❌)
// דורש: שירותים (+) → Google Calendar API, והרצה אחת של authorizeCalendar
const TASKS_CALENDAR_ID = 'primary';
function syncCalendar_(t) {
  const ref = String(t.ref || '');
  if (ref.indexOf('cal:') !== 0) return;
  const prefix = t.status === 'done' ? '✅ ' : t.status === 'cancelled' ? '❌ ' : '';
  try {
    // בוצע / בוטל → בלי התראות; חזר לפתוח → התראות ברירת המחדל של היומן
    const closed = t.status === 'done' || t.status === 'cancelled';
    Calendar.Events.patch({
      summary: prefix + t.title,
      reminders: closed ? { useDefault: false, overrides: [] } : { useDefault: true },
    }, TASKS_CALENDAR_ID, ref.slice(4));
  } catch (e) {
    console.log('calendar sync failed', e);
  }
}
// היומן → המשימות: שינוי שם / תאריך / ✅ / ❌ / מחיקה ביומן מתעדכנים במשימות,
// ואירועים חדשים של השבוע נכנסים כמשימות. רץ לכל היותר פעם ב-2 דקות (בפתיחת הדף / שאלה לבוט).
function syncFromCalendar_() {
  const cache = CacheService.getScriptCache();
  if (cache.get('calsync')) return;
  cache.put('calsync', '1', 120);
  try {
    const tz = 'Asia/Jerusalem';
    const byRef = {};
    tasksAll_().forEach(function (t) { if (t.ref) byRef[t.ref] = t; });
    const ws = weekStart_(new Date());
    const we = new Date(ws.getTime() + 7 * 86400000);
    const res = Calendar.Events.list(TASKS_CALENDAR_ID, {
      timeMin: ws.toISOString(), timeMax: we.toISOString(), singleEvents: true, showDeleted: true, maxResults: 250,
    });
    (res.items || []).forEach(function (ev) {
      const ref = 'cal:' + ev.id;
      const t = byRef[ref];
      let summary = String(ev.summary || '').trim();
      let st = ev.status === 'cancelled' ? 'cancelled' : null;
      if (/^✅/.test(summary)) { st = 'done'; summary = summary.replace(/^✅\s*/, ''); }
      else if (/^❌/.test(summary)) { st = 'cancelled'; summary = summary.replace(/^❌\s*/, ''); }
      const day = !ev.start ? '' : ev.start.date ? ev.start.date
        : Utilities.formatDate(new Date(ev.start.dateTime), tz, 'yyyy-MM-dd');
      const hm = ev.start && ev.start.dateTime
        ? Utilities.formatDate(new Date(ev.start.dateTime), tz, 'HH:mm') : '';
      if (t) {
        let changed = false;
        if (summary && summary !== t.title) { t.title = summary; changed = true; }
        if (day && day !== t.due) { t.due = day; changed = true; }
        if (day && hm !== String(t.time || '')) { t.time = hm; changed = true; }
        if (st && t.status !== st && t.status !== 'done') {
          t.status = st; t.doneAt = st === 'done' ? new Date().toISOString() : ''; changed = true;
        }
        if (changed) { t.updated = new Date().toISOString(); taskWrite_(t); }
      } else if (summary && !st) {
        addTask_({ title: summary, area: 'אחר', source: 'calendar', due: day, time: hm, ref: ref });
      }
    });
  } catch (e) {
    console.log('calendar → tasks failed', e);
  }
}

function authorizeCalendar() {
  Logger.log('יומן: ' + Calendar.Calendars.get(TASKS_CALENDAR_ID).summary);
}

// להרצה ידנית פעם אחת: יוצר את הגיליון והמפתח, ומדפיס את הקישור לדף
function setupTasks() {
  tasksSheet_();
  Logger.log('דף המשימות: ' + tasksPageUrl_());
  Logger.log('גיליון: ' + SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('TASKS_SHEET')).getUrl());
}

const TASKS_PAGE_HTML = `<!doctype html>
<html dir="rtl" lang="he"><head><meta charset="utf-8">
<style>
  :root{--bg:#f6f5f2;--card:#fff;--ink:#1f2328;--mute:#6b7075;--line:#e4e2dc;--ok:#1a7f4b;--warn:#b8860b;--bad:#b3261e;--acc:#2f5bd3}
  *{box-sizing:border-box} body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.45 system-ui,-apple-system,"Segoe UI",Arial,sans-serif}
  header{position:sticky;top:0;background:var(--bg);padding:14px 16px 8px;border-bottom:1px solid var(--line);z-index:2}
  h1{font-size:20px;margin:0 0 6px} .stats{display:flex;gap:8px;flex-wrap:wrap;font-size:13px;color:var(--mute)}
  .pill{background:var(--card);border:1px solid var(--line);border-radius:99px;padding:2px 10px}
  main{padding:8px 16px 90px;max-width:720px;margin:0 auto}
  h2{font-size:15px;margin:18px 0 6px;color:var(--mute)}
  .t{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:10px 12px;margin:8px 0}
  .t.done .ttl{text-decoration:line-through;color:var(--mute)}
  .row{display:flex;align-items:center;gap:10px}
  .row input[type=checkbox]{width:24px;height:24px;flex:none;accent-color:var(--ok)}
  .ttl{flex:1;font-weight:600;border-radius:6px;padding:2px 4px;outline:none}
  .ttl:focus{background:var(--bg);box-shadow:0 0 0 2px var(--acc)}
  .meta{font-size:12.5px;color:var(--mute);margin:4px 34px 0 0}
  .stuck{color:var(--bad);font-weight:600}
  .ctl{display:flex;gap:6px;margin:8px 34px 0 0;flex-wrap:wrap}
  select,textarea,input[type=text]{font:inherit;font-size:14px;border:1px solid var(--line);border-radius:8px;padding:6px 8px;background:#fff;color:var(--ink)}
  textarea{width:100%;min-height:36px;margin:6px 0 0}
  .btn{display:inline-block;text-decoration:none;font-size:13px;border:1px solid var(--line);border-radius:8px;padding:5px 10px;color:var(--ink);background:#fff}
  .saved{font-size:12px;color:var(--ok);margin-inline-start:6px}
  .add{display:flex;gap:6px;margin:12px 0} .add input{flex:1}
  @media (prefers-color-scheme:dark){:root{--bg:#16181b;--card:#1f2226;--ink:#e8e6e1;--mute:#9aa0a6;--line:#30343a}select,textarea,input[type=text],.btn{background:#26292e}}
</style></head><body>
<header><h1>סגירת שבוע – משימות</h1><div class="stats" id="stats">טוען…</div></header>
<main>
  <div class="add"><input type="text" id="newTitle" placeholder="משימה חדשה…"><select id="newArea"></select><button class="btn" onclick="addNew()">הוסף</button></div>
  <div id="list"></div>
</main>
<script>
const KEY = '<?!= key ?>';
const ST = {open:'פתוח',progress:'בתהליך',done:'בוצע',postponed:'נדחה',cancelled:'בוטל'};
let DATA = [], AREAS = [];
function esc(s){return String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
function fmt(d){if(!d)return'';const x=new Date(d);return isNaN(x)?d:x.toLocaleDateString('he-IL',{day:'numeric',month:'numeric'})}
function tel(p){let d=String(p||'').replace(/\\D/g,'');if(d.startsWith('0'))d='972'+d.slice(1);return d}
function load(){google.script.run.withSuccessHandler(r=>{DATA=r.items;AREAS=r.areas;render()}).withFailureHandler(e=>{document.getElementById('stats').textContent='שגיאה: '+e.message}).tasksPageData(KEY)}
function render(){
  const c={done:0,progress:0,open:0,postponed:0,cancelled:0};DATA.forEach(t=>c[t.status]=(c[t.status]||0)+1);
  document.getElementById('stats').innerHTML=['<span class="pill">סה"כ '+DATA.length+'</span>','<span class="pill">✅ '+c.done+' בוצעו</span>','<span class="pill">🔄 '+c.progress+' בתהליך</span>','<span class="pill">⬜ '+c.open+' פתוחות</span>','<span class="pill">⏭ '+c.postponed+' נדחו</span>'].join('');
  document.getElementById('newArea').innerHTML=AREAS.map(a=>'<option>'+esc(a)+'</option>').join('');
  const waiting=DATA.filter(t=>t.owner&&t.status!=='done'&&t.status!=='cancelled');
  const mine=DATA.filter(t=>!waiting.includes(t));
  let h='';
  AREAS.forEach(a=>{const g=mine.filter(t=>(t.area||'אחר')===a);if(g.length)h+='<h2>'+esc(a)+'</h2>'+g.map(card).join('')});
  if(waiting.length)h+='<h2>⏳ מחכה לאחרים</h2>'+waiting.map(card).join('');
  document.getElementById('list').innerHTML=h||'<p>אין משימות השבוע 🙂</p>';
}
function card(t){
  const opts=Object.keys(ST).map(s=>'<option value="'+s+'"'+(s===t.status?' selected':'')+'>'+ST[s]+'</option>').join('');
  const m=[];
  if(t.owner)m.push('באחריות: '+esc(t.owner));
  if(t.contact)m.push('👤 '+esc(t.contact));
  if(t.due)m.push('עד: '+fmt(t.due));
  if(t.source&&t.source!=='daniel'&&t.source!=='page')m.push('נוסף ע"י הבוט');
  if(t.weeks>=2&&t.status!=='done')m.push('<span class="stuck">נדחה '+t.weeks+' שבועות</span>');
  let links='';
  if(t.phone){const d=tel(t.phone);links+='<a class="btn" target="_top" href="tel:+'+d+'">📞</a><a class="btn" target="_blank" href="https://wa.me/'+d+(t.owner?'?text='+encodeURIComponent('היי, רק מזכיר: '+t.title):'')+'">💬</a>'}
  return '<div class="t '+(t.status==='done'?'done':'')+'" id="t_'+t.id+'"><div class="row"><input type="checkbox" '+(t.status==='done'?'checked':'')+' onchange="setSt(\\''+t.id+'\\',this.checked?\\'done\\':\\'open\\')"><div class="ttl" contenteditable="true" title="לחיצה לעריכת השם" onblur="setTitle(\\''+t.id+'\\',this.innerText)">'+esc(t.title)+'</div></div>'+
   (m.length?'<div class="meta">'+m.join(' · ')+'</div>':'')+
   '<div class="ctl"><select onchange="setSt(\\''+t.id+'\\',this.value)">'+opts+'</select>'+'<select onchange="setArea(\\''+t.id+'\\',this.value)">'+AREAS.map(a=>'<option'+(a===(t.area||'אחר')?' selected':'')+'>'+esc(a)+'</option>').join('')+'</select>'+links+'<span class="saved" id="s_'+t.id+'"></span></div>'+
   '<div class="ctl" style="display:block"><textarea placeholder="הערה…" onchange="setNote(\\''+t.id+'\\',this.value)">'+esc(t.note)+'</textarea></div></div>';
}
function save(id,st,note,area,title){const s=document.getElementById('s_'+id);if(s)s.textContent='שומר…';google.script.run.withSuccessHandler(()=>{if(s)s.textContent='✓ נשמר'}).withFailureHandler(e=>{if(s)s.textContent='שגיאה'}).tasksPageUpdate(KEY,id,st,note,area||'',title||'')}
function setSt(id,st){const t=DATA.find(x=>x.id===id);t.status=st;save(id,st,t.note);render()}
function setTitle(id,v){const t=DATA.find(x=>x.id===id);v=String(v||'').replace(/\\s+/g,' ').trim();if(!v||v===t.title)return;t.title=v;save(id,t.status,t.note,'',v)}
function setArea(id,a){const t=DATA.find(x=>x.id===id);t.area=a;save(id,t.status,t.note,a);render()}
function setNote(id,v){const t=DATA.find(x=>x.id===id);t.note=v;save(id,t.status,v)}
function addNew(){const i=document.getElementById('newTitle');const v=i.value.trim();if(!v)return;const a=document.getElementById('newArea').value;google.script.run.withSuccessHandler(()=>{i.value='';load()}).tasksPageAdd(KEY,v,a)}
load();
</script></body></html>`;

function testTasks() {
  function t(p) { p.secret = SECRET; const o = JSON.parse(doPost({ postData: { contents: JSON.stringify(p) } }).getContent()); Logger.log(p.action + ' → ' + JSON.stringify(o).slice(0, 300)); }
  t({ action: 'addTask', title: 'בדיקת מערכת המשימות – אפשר לסמן בוצע', area: 'אישי', source: 'daniel' });
  t({ action: 'listTasks' });
  t({ action: 'weeklySummary' });
}
