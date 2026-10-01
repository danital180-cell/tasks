// Daniel Assistant — עוזר אישי בוואטסאפ (Supabase Edge Function)
// תזכורות + יומן גוגל + ג'ימייל + דרייב + סיכום בוקר + חיפוש באינטרנט + אנשי קשר + חיוג + שליחת מייל + משימות ודוח שבועי
import { createClient } from "npm:@supabase/supabase-js@2";
import { decodeBase64, encodeBase64 } from "jsr:@std/encoding@1/base64";

const TZ = "Asia/Jerusalem";
const GRAPH = "https://graph.facebook.com/v25.0";

const env = (k: string, def = "") => Deno.env.get(k) ?? def;
const WHATSAPP_TOKEN = env("WHATSAPP_TOKEN");
const PHONE_NUMBER_ID = env("PHONE_NUMBER_ID");
const VERIFY_TOKEN = env("VERIFY_TOKEN");
const APP_SECRET = env("APP_SECRET");
const OWNER_PHONE = env("OWNER_PHONE");
const GEMINI_API_KEY = env("GEMINI_API_KEY");
const GEMINI_MODEL = env("GEMINI_MODEL", "gemini-flash-lite-latest");
const GEMINI_FALLBACK_MODEL = env("GEMINI_FALLBACK_MODEL", "gemini-flash-latest");
const CALENDAR_ID = env("CALENDAR_ID");
const CRON_SECRET = env("CRON_SECRET");
const TEMPLATE_NAME = env("TEMPLATE_NAME"); // תבנית לתזכורות מחוץ לחלון 24 השעות (בהמשך)
const TEMPLATE_LANG = env("TEMPLATE_LANG", "he");
const SUMMARY_TIME = env("SUMMARY_TIME", "07:30");
const GAS_URL = env("GAS_URL");       // כתובת ה-Apps Script (ג'ימייל + דרייב)
const GAS_SECRET = env("GAS_SECRET"); // אותו ערך כמו SECRET בסקריפט
const HOME_AREA = env("HOME_AREA", "נהריה"); // אזור ברירת מחדל לשאלות "מה פתוח / איפה"
const WEEKLY_TIME = env("WEEKLY_TIME", "15:00"); // דוח משימות שבועי – יום שישי בשעה הזו
const AREAS = ["הרצל בוטיק", "פסגות העמק", "טאבונוצי", "RETAIN", "אישי", "אחר"];

// מפתח השרת: קודם המפתחות החדשים של Supabase, ואם אין — המפתח הישן
function serverKey(): string {
  try {
    const keys = JSON.parse(env("SUPABASE_SECRET_KEYS", "{}"));
    if (keys?.default) return keys.default;
  } catch { /* ignore */ }
  return env("SUPABASE_SERVICE_ROLE_KEY");
}
const db = createClient(env("SUPABASE_URL"), serverKey());

const HELP =
  "אני העוזר האישי שלך 🤖\nאפשר לכתוב לי למשל:\n" +
  "• תזכיר לי מחר ב-10 להתקשר לשמעון\n" +
  "• פגישה ביום חמישי ב-14:00 עם משפחת כהן\n" +
  "• מה יש לי היום? / מחר?\n" +
  "• מה חדש במייל? / יש מייל משמעון?\n" +
  "• שלח לי את החוזה של כהן (מהדרייב)\n" +
  "• תכין הודעה ל-0501234567: אני מגיע ב-5\n" +
  "• מה עם החשבוניות? (ייבוא חשבוניות לדרייב)\n" +
  "• האם קופת חולים בנהריה פתוחה עכשיו? (חיפוש באינטרנט)\n" +
  "• מה הטלפון של שמעון? / תחייג לשמעון\n" +
  "• תשלח מייל לשמעון שהפגישה נדחית למחר (אשאל לפני שליחה)\n" +
  "• צילום חשבונית → נשמר ב\"חשבוניות למע\"מ\" ונשלח מיד לפייפרלס\n" +
  "• קובץ או תמונה אחרים יישמרו בדרייב בתיקייה \"עוזר אישי\"\n" +
  "• משימה: לסגור מחיר עם הספק של טאבונוצי (איש קשר: מתן)\n" +
  "• משימה: אורלי צריכה לשלוח הצעת מחיר (מחכה לאחרים)\n" +
  "• מה המשימות שלי? / סיימתי עם הספק\n" +
  "• בכל שישי ב-15:00 אשלח לך דוח שבועי עם דף לסימון";

// ---------- זמן (שעון ישראל, כולל שעון קיץ/חורף) ----------
function tzOffsetMin(d: Date): number {
  const part = new Intl.DateTimeFormat("en-US", { timeZone: TZ, timeZoneName: "longOffset" })
    .formatToParts(d).find((p) => p.type === "timeZoneName")?.value ?? "GMT";
  const m = part.match(/GMT([+-])(\d{1,2}):?(\d{2})?/);
  if (!m) return 0;
  return (m[1] === "-" ? -1 : 1) * (Number(m[2]) * 60 + Number(m[3] ?? 0));
}
function localToUtc(local: string): Date {
  const s = local.length === 16 ? local + ":00" : local.slice(0, 19);
  const base = new Date(s + "Z");
  const first = new Date(base.getTime() - tzOffsetMin(base) * 60000);
  return new Date(base.getTime() - tzOffsetMin(first) * 60000);
}
function localParts(d = new Date()) {
  const s = d.toLocaleString("sv-SE", { timeZone: TZ });
  return { date: s.slice(0, 10), time: s.slice(11, 16), full: s.slice(0, 16).replace(" ", "T") };
}
function addDays(dateStr: string, n: number) {
  const d = new Date(dateStr + "T12:00:00Z");
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}
const heTime = (d: Date) =>
  new Intl.DateTimeFormat("he-IL", { timeZone: TZ, hour: "2-digit", minute: "2-digit" }).format(d);
const heDateTime = (d: Date) =>
  new Intl.DateTimeFormat("he-IL", {
    timeZone: TZ, weekday: "long", day: "numeric", month: "numeric", hour: "2-digit", minute: "2-digit",
  }).format(d);
const heShort = (d: Date) =>
  new Intl.DateTimeFormat("he-IL", { timeZone: TZ, day: "numeric", month: "numeric" }).format(d);
const heDay = (dateStr: string) =>
  new Intl.DateTimeFormat("he-IL", { timeZone: TZ, weekday: "long", day: "numeric", month: "numeric" })
    .format(new Date(dateStr + "T12:00:00Z"));
const heWeekday = () => new Intl.DateTimeFormat("he-IL", { timeZone: TZ, weekday: "long" }).format(new Date());
// 0=ראשון ... 5=שישי (שעון ישראל)
const dowIL = (d = new Date()) =>
  ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].indexOf(
    new Intl.DateTimeFormat("en-US", { timeZone: TZ, weekday: "short" }).format(d));

// ---------- וואטסאפ ----------
const waAuth = () => ({ Authorization: `Bearer ${WHATSAPP_TOKEN}` });
async function waPost(payload: unknown) {
  const r = await fetch(`${GRAPH}/${PHONE_NUMBER_ID}/messages`, {
    method: "POST",
    headers: { ...waAuth(), "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(JSON.stringify(j?.error ?? j));
  return j;
}
const sendText = (to: string, body: string) =>
  waPost({ messaging_product: "whatsapp", to, type: "text", text: { body: body.slice(0, 4000), preview_url: false } });

async function notify(body: string, templateParam: string) {
  try {
    await sendText(OWNER_PHONE, body);
  } catch (e) {
    if (TEMPLATE_NAME && String(e).includes("131047")) {
      await waPost({
        messaging_product: "whatsapp", to: OWNER_PHONE, type: "template",
        template: {
          name: TEMPLATE_NAME, language: { code: TEMPLATE_LANG },
          components: [{ type: "body", parameters: [{ type: "text", text: templateParam.replace(/\s+/g, " ").slice(0, 900) }] }],
        },
      });
    } else throw e;
  }
}

async function uploadMedia(bytes: Uint8Array, mime: string, name: string): Promise<string> {
  const fd = new FormData();
  fd.append("messaging_product", "whatsapp");
  fd.append("type", mime);
  fd.append("file", new Blob([bytes], { type: mime }), name);
  const r = await fetch(`${GRAPH}/${PHONE_NUMBER_ID}/media`, { method: "POST", headers: waAuth(), body: fd });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error("upload: " + JSON.stringify(j?.error ?? j));
  return j.id;
}

async function downloadMedia(mediaId: string) {
  const meta = await (await fetch(`${GRAPH}/${mediaId}`, { headers: waAuth() })).json();
  if (!meta?.url) throw new Error("media meta: " + JSON.stringify(meta));
  const r = await fetch(meta.url, { headers: waAuth() });
  if (!r.ok) throw new Error("media download " + r.status);
  return { bytes: new Uint8Array(await r.arrayBuffer()), mime: meta.mime_type as string };
}

async function verifySignature(raw: string, header: string | null) {
  if (!header?.startsWith("sha256=")) return false;
  const key = await crypto.subtle.importKey(
    "raw", new TextEncoder().encode(APP_SECRET), { name: "HMAC", hash: "SHA-256" }, false, ["sign"],
  );
  const sig = new Uint8Array(await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(raw)));
  const hex = [...sig].map((b) => b.toString(16).padStart(2, "0")).join("");
  return hex === header.slice(7);
}

// ---------- יומן גוגל (Service Account) ----------
let gToken = { value: "", exp: 0 };
function b64url(bytes: Uint8Array) {
  let s = "";
  bytes.forEach((b) => (s += String.fromCharCode(b)));
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
async function googleToken(): Promise<string> {
  const now = Math.floor(Date.now() / 1000);
  if (gToken.value && gToken.exp - 60 > now) return gToken.value;
  const sa = JSON.parse(env("GOOGLE_SERVICE_ACCOUNT_JSON"));
  const enc = (o: unknown) => b64url(new TextEncoder().encode(JSON.stringify(o)));
  const unsigned = enc({ alg: "RS256", typ: "JWT" }) + "." + enc({
    iss: sa.client_email, scope: "https://www.googleapis.com/auth/calendar",
    aud: "https://oauth2.googleapis.com/token", iat: now, exp: now + 3600,
  });
  const pem = sa.private_key.replace(/-----[^-]+-----/g, "").replace(/\s+/g, "");
  const der = Uint8Array.from(atob(pem), (c) => c.charCodeAt(0));
  const key = await crypto.subtle.importKey(
    "pkcs8", der, { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" }, false, ["sign"],
  );
  const sig = new Uint8Array(await crypto.subtle.sign("RSASSA-PKCS1-v1_5", key, new TextEncoder().encode(unsigned)));
  const r = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer", assertion: unsigned + "." + b64url(sig) }),
  });
  const j = await r.json();
  if (!r.ok) throw new Error("Google auth: " + JSON.stringify(j));
  gToken = { value: j.access_token, exp: now + (j.expires_in ?? 3600) };
  return gToken.value;
}
const calUrl = () => `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(CALENDAR_ID)}/events`;

async function createEvent(title: string, start: Date, durationMin: number): Promise<string> {
  const end = new Date(start.getTime() + durationMin * 60000);
  const r = await fetch(calUrl(), {
    method: "POST",
    headers: { Authorization: `Bearer ${await googleToken()}`, "Content-Type": "application/json" },
    body: JSON.stringify({
      summary: title,
      start: { dateTime: start.toISOString(), timeZone: TZ },
      end: { dateTime: end.toISOString(), timeZone: TZ },
    }),
  });
  const j = await r.json();
  if (!r.ok) throw new Error("Calendar: " + JSON.stringify(j?.error ?? j));
  return j.id;
}

async function listEvents(dateStr: string) {
  const qs = new URLSearchParams({
    singleEvents: "true", orderBy: "startTime",
    timeMin: localToUtc(dateStr + "T00:00").toISOString(),
    timeMax: localToUtc(addDays(dateStr, 1) + "T00:00").toISOString(),
  });
  const r = await fetch(`${calUrl()}?${qs}`, { headers: { Authorization: `Bearer ${await googleToken()}` } });
  const j = await r.json();
  if (!r.ok) throw new Error("Calendar: " + JSON.stringify(j?.error ?? j));
  return (j.items ?? []) as any[];
}

function formatAgenda(items: any[], header: string) {
  if (!items.length) return `${header}\nאין אירועים ביומן 🙂`;
  const lines = items.map((ev) => {
    const t = ev.start?.dateTime ? heTime(new Date(ev.start.dateTime)) : "כל היום";
    return `• ${t} – ${ev.summary ?? "(ללא כותרת)"}`;
  });
  return `${header}\n${lines.join("\n")}`;
}

// ---------- ג'ימייל + דרייב + אנשי קשר (דרך Apps Script) ----------
async function gas(action: string, params: Record<string, unknown> = {}) {
  if (!GAS_URL || !GAS_SECRET) throw new Error("GAS not configured");
  const r = await fetch(GAS_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ secret: GAS_SECRET, action, ...params }),
    redirect: "follow",
  });
  const text = await r.text();
  let j: any;
  try { j = JSON.parse(text); } catch { throw new Error("GAS bad response: " + text.slice(0, 200)); }
  if (j?.error) throw new Error("GAS: " + j.error);
  return j;
}

const senderName = (from: string) => (from.split("<")[0].replace(/"/g, "").trim() || from);

function formatMails(items: any[], header: string, withSnippet = false) {
  if (!items?.length) return `${header}\nלא נמצאו מיילים 📭`;
  const lines = items.map((m) => {
    let line = `• ${heShort(new Date(m.date))} ${senderName(m.from)} – ${m.subject}`;
    if (withSnippet && m.snippet) line += `\n   ${m.snippet.slice(0, 120)}…`;
    return line;
  });
  return `${header}\n${lines.join("\n")}`;
}

async function sendDriveFile(f: any) {
  const res = await gas("getFile", { id: f.id });
  if (res.tooLarge || !res.base64) {
    await sendText(OWNER_PHONE, `📁 ${res.name ?? f.name}\nהקובץ גדול מדי לשליחה, הנה קישור:\n${res.url ?? f.url}`);
    return;
  }
  try {
    const bytes = decodeBase64(res.base64);
    const mediaId = await uploadMedia(bytes, res.mimeType, res.name);
    const isImage = String(res.mimeType).startsWith("image/");
    await waPost(isImage
      ? { messaging_product: "whatsapp", to: OWNER_PHONE, type: "image", image: { id: mediaId, caption: res.name } }
      : { messaging_product: "whatsapp", to: OWNER_PHONE, type: "document", document: { id: mediaId, filename: res.name } });
  } catch (e) {
    console.error(e); // סוג קובץ שוואטסאפ לא תומך בו → קישור
    await sendText(OWNER_PHONE, `📁 ${res.name}\n${res.url}`);
  }
}

async function saveIncomingMedia(msg: any) {
  const kind = msg.type as string;
  const media = msg[kind];
  const { bytes, mime } = await downloadMedia(media.id);
  if (bytes.length > 20 * 1024 * 1024) {
    await sendText(OWNER_PHONE, "⚠️ הקובץ גדול מדי לשמירה אוטומטית (מעל 20MB).");
    return;
  }
  const ext = (mime.split("/")[1] ?? "bin").split(";")[0].replace("jpeg", "jpg");
  const stamp = localParts().full.replace("T", "_").replace(":", "-");
  const name = media.filename || `${kind}_${stamp}.${ext}`;
  const caption = String(media.caption ?? "").trim();
  const b64 = encodeBase64(bytes);

  // חשבונית? (תמונה או PDF) → תיקיית החודש ב"חשבוניות למע"מ" + שליחה מיידית לפייפרלס
  const canCheck = /^image\//.test(mime) || mime === "application/pdf";
  if (canCheck && !/לא חשבונית/.test(caption)) {
    let doc: any = null;
    try { doc = await classifyDoc(b64, mime, caption); } catch (e) { console.error("classify", e); }
    const forced = /חשבונית|קבלה/.test(caption);
    if (doc?.is_invoice || forced) {
      const inv = await gas("saveInvoice", {
        name, mimeType: mime, base64: b64,
        date: doc?.date || "", vendor: doc?.vendor || "", sendNow: true,
      });
      await autoTask({
        title: `חשבונית ${doc?.vendor || ""} ${inv.sent ? "נשלחה לפייפרלס" : "נשמרה (טרם נשלחה)"}`.replace(/\s+/g, " "),
        area: "אחר", source: "invoice", status: inv.sent ? "done" : "open", ref: "inv:" + inv.url,
      });
      const details = [doc?.vendor && `ספק: ${doc.vendor}`, doc?.date && `תאריך: ${doc.date}`, doc?.amount && `סכום: ${doc.amount}`]
        .filter(Boolean).join(" | ");
      await sendText(OWNER_PHONE,
        `🧾 זיהיתי חשבונית${details ? `\n${details}` : ""}\n` +
        `📁 נשמרה ב"חשבוניות למע"מ" / ${inv.folder}\n` +
        (inv.sent ? "📤 נשלחה עכשיו לפייפרלס ✅" : "⚠️ לא נשלחה לפייפרלס – תישלח במשלוח של ה-3 לחודש") +
        `\n\n(אם זו לא חשבונית – שלח שוב עם הכיתוב "לא חשבונית")`);
      return;
    }
  }

  const saved = await gas("saveFile", { name, mimeType: mime, base64: b64 });
  await sendText(OWNER_PHONE, `💾 נשמר בדרייב בתיקייה "עוזר אישי":\n${saved.name}\n(חשבונית שלא זוהתה? שלח שוב עם הכיתוב "חשבונית")`);
}

// זיהוי מסמך בעזרת Gemini: האם זו חשבונית, ומה התאריך/הספק/הסכום
async function classifyDoc(b64: string, mime: string, caption: string) {
  const j = await geminiCall({
    contents: [{
      role: "user",
      parts: [
        { inline_data: { mime_type: mime, data: b64 } },
        {
          text: `דניאל (עוסק מורשה) שלח את המסמך הזה לעוזר האישי שלו${caption ? ` עם הכיתוב: "${caption}"` : ""}.
האם זו חשבונית מס / חשבונית מס קבלה / קבלה על הוצאה (למשל תדלוק, מסעדה, ספק)? צילום רגיל, מסמך אחר או צילום מסך – לא.
החזר JSON בלבד: {"is_invoice":true,"date":"YYYY-MM-DD","vendor":"שם העסק בקצרה","amount":"הסכום הכולל עם ₪"}
date = תאריך המסמך עצמו (לא היום). אם לא ברור – השאר ריק.`,
        },
      ],
    }],
    generationConfig: { responseMimeType: "application/json", temperature: 0 },
  });
  return JSON.parse(geminiText(j).replace(/```json|```/g, "").trim());
}

function formatInvoices(st: any, full: boolean) {
  const lines: string[] = [];
  const stale = !st.lastRun || Date.now() - new Date(st.lastRun).getTime() > 48 * 3600 * 1000;
  if (st.lastError) lines.push(`⚠️ ייבוא החשבוניות נכשל לאחרונה:\n${String(st.lastError).slice(0, 200)}`);
  else if (stale) lines.push("⚠️ ייבוא החשבוניות לא רץ ביומיים האחרונים – כדאי לבדוק.");
  if (full || st.newCount) {
    lines.push(`🧾 ${st.newCount} חשבוניות חדשות${full ? " מאז הסיכום האחרון" : " מאז אתמול"} (החודש: ${st.monthCount})`);
    if (full && st.newItems?.length) lines.push(st.newItems.map((x: string) => `• ${x}`).join("\n"));
  }
  if (st.plLastError) lines.push(`⚠️ השליחה לפייפרלס נכשלה:\n${String(st.plLastError).slice(0, 200)}`);
  if (st.plLast && (st.plNew || full)) {
    const when = heShort(new Date(st.plLast.at));
    lines.push(st.plLast.count
      ? `📤 נשלחו לפייפרלס ${st.plLast.count} חשבוניות (${st.plLast.months}) ב-${when}`
      : `📤 בדיקת השליחה לפייפרלס ב-${when}: לא היו חשבוניות חדשות לשלוח`);
  }
  if (full && st.folderUrl) lines.push(`📁 ${st.folderUrl}`);
  return lines.join("\n");
}

function normPhone(p: string) {
  let d = String(p ?? "").replace(/\D/g, "");
  if (d.startsWith("0")) d = "972" + d.slice(1);
  return d;
}
// מספר בפורמט בינלאומי – וואטסאפ הופך אותו ללחיץ (חיוג / הודעה)
const dialable = (p: string) => "+" + normPhone(p);

// ---------- Gemini ----------
// קריאה עם ניסיונות חוזרים: מודל מהיר קודם, ואם עמוס/נגמרה מכסה – המודל השני
async function geminiCall(body: unknown): Promise<any> {
  const models = [GEMINI_MODEL, GEMINI_FALLBACK_MODEL];
  let lastErr = "";
  for (let round = 0; round < 2; round++) {
    for (const model of models) {
      try {
        const r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`, {
          method: "POST",
          headers: { "Content-Type": "application/json", "x-goog-api-key": GEMINI_API_KEY },
          body: JSON.stringify(body),
          signal: AbortSignal.timeout(25000),
        });
        const j = await r.json().catch(() => ({}));
        if (r.ok) return j;
        lastErr = `${model} ${r.status}: ${JSON.stringify(j?.error ?? j).slice(0, 300)}`;
        console.error("Gemini", lastErr);
        if (![429, 500, 503].includes(r.status)) throw new Error(lastErr);
      } catch (e) {
        lastErr = `${model}: ${String(e)}`;
        console.error("Gemini", lastErr);
      }
    }
    await new Promise((res) => setTimeout(res, 2000));
  }
  throw new Error("Gemini: " + lastErr);
}
const geminiText = (j: any) =>
  (j.candidates?.[0]?.content?.parts ?? []).filter((p: any) => !p.thought).map((p: any) => p.text ?? "").join("").trim();

function geminiErrorReply(e: unknown) {
  const s = String(e);
  if (s.includes("429") || s.includes("RESOURCE_EXHAUSTED")) {
    return "😕 נגמרה המכסה של Gemini (מנוע ההבנה שלי) להיום. תזכורות קיימות וסיכום הבוקר ממשיכים לעבוד. שאלות על חשבוניות עובדות גם עכשיו.";
  }
  return "😕 לא הצלחתי לעבד את ההודעה כרגע. נסה שוב בעוד דקה.";
}

// ---------- הבנת ההודעה ----------
async function understand(text: string) {
  const now = localParts();
  const system = `אתה העוזר האישי של דניאל. עכשיו: ${now.full} (${heWeekday()}), שעון ישראל. דניאל גר ב${HOME_AREA}.
החזר JSON בלבד, במבנה:
{"action":"create|agenda|mail_unread|mail_search|drive_search|wa_draft|invoices|web|contact|call|email|task|tasks_list|task_done|chat","title":"","start":"YYYY-MM-DDTHH:MM","duration_minutes":0,"remind_minutes_before":0,"date":"YYYY-MM-DD","query":"","phone":"","to":"","subject":"","message":"","area":"","contact":"","owner":"","due":"","reply":""}
כללים:
- "תזכיר לי..." / "תזכורת..." => create, duration_minutes=15, remind_minutes_before=0.
- פגישה/אירוע => create, duration_minutes=60 אלא אם צוין, remind_minutes_before=30 אלא אם צוין.
- לא צוינה שעה => 09:00. צוינה רק שעה שכבר עברה היום => מחר. "ב-5"/"ב-6" אחה"צ סביר => 17:00/18:00.
- "מה יש לי היום/מחר/ביום X" => agenda עם date.
- "מה חדש במייל" / "מיילים שלא קראתי" => mail_unread.
- חיפוש מייל ("יש מייל משמעון", "המייל מהבנק על המשכנתא") => mail_search, ו-query בתחביר חיפוש של ג'ימייל, למשל: שמעון newer_than:30d  או  from:bank subject:משכנתא.
- בקשת קובץ ("שלח לי את החוזה של כהן", "תמצא בדרייב...") => drive_search, ו-query = מילות מפתח קצרות משם הקובץ בלבד (למשל: חוזה כהן).
- "תכין/תנסח הודעה ל-<מספר>: <טקסט>" או "תענה לשמעון 050... ש..." => wa_draft, phone = המספר כפי שנכתב, message = נוסח ההודעה בגוף ראשון מדניאל, מנומס וקצר.
- שאלה כללית על חשבוניות ("מה עם חשבוניות", "כמה חשבוניות נכנסו החודש") => invoices. drive_search רק כשמבקשים קובץ מסוים לשליחה.
- שאלה שצריך בשבילה מידע עדכני מהאינטרנט: שעות פתיחה ("האם קופת חולים פתוחה"), טלפון/כתובת של עסק, מוסד, רופא או חנות, מזג אוויר, חדשות, מחירים, תוצאות, "מה זה X" => web, query = השאלה המלאה בעברית; אם לא צוין מקום ורלוונטי – הוסף "${HOME_AREA}".
- "מה הטלפון/המייל של <אדם>", "תמצא לי את המספר של <אדם>" => contact, query = שם האדם בלבד. אם זה עסק או מוסד (ולא אדם שדניאל מכיר) => web.
- "תחייג/תתקשר/תחבר אותי ל<שם או מספר>" => call, query = השם, phone = המספר אם נכתב. עסק/מוסד בלי מספר => web ו-query = "מספר הטלפון של <העסק>".
- "תשלח/תכתוב מייל ל<שם או כתובת> ש..." => email: to = כתובת המייל אם נכתבה, אחרת ריק; query = שם הנמען; subject = נושא קצר; message = גוף המייל בעברית מנומסת, בגוף ראשון מדניאל, עם "שלום <שם>," בהתחלה ו"בברכה,\nדניאל טרבלסי" בסוף.
- "משימה: ..." / "תוסיף משימה" / "צריך ל..." / "לא לשכוח ל..." בלי שעה מסוימת => task. עם יום ושעה => create.
  title = המשימה בקצרה. contact = שם האדם שקשור למשימה (אם יש). owner = שם מי שצריך לבצע אם זה לא דניאל ("מחכה לשמעון שישלח", "אורלי צריכה ל...") – אחרת ריק. due = YYYY-MM-DD אם צוין מועד ("עד יום חמישי"), אחרת ריק. phone = מספר אם נכתב.
- area (ל-task ול-create): הרצל בוטיק (עכו, דקלי פסגות, הרצל 26) | פסגות העמק (עפולה) | טאבונוצי (פיצה, סדנאות, אירועים, מעלות, מתן) | RETAIN | אישי (בית, משפחה, בריאות, רכב, בנק) | אחר. בחר לפי ההקשר.
- "מה המשימות שלי" / "מה פתוח" / "מה מחכה לי" / "מה עם המשימות של שמעון" => tasks_list, query = שם אדם או תחום אם צוין, אחרת ריק.
- "סיימתי את/עם ..." / "בוצע ..." / "תסמן שסיימתי ..." => task_done, query = 1-3 מילות מפתח מהמשימה.
- title: קצר, בעברית, בלי "תזכיר לי" ובלי "משימה:".
- כל דבר אחר => chat, ו-reply קצר בעברית.`;
  const j = await geminiCall({
    systemInstruction: { parts: [{ text: system }] },
    contents: [{ role: "user", parts: [{ text }] }],
    generationConfig: { responseMimeType: "application/json", temperature: 0.1 },
  });
  return JSON.parse(geminiText(j).replace(/```json|```/g, "").trim());
}

// ---------- חיפוש באינטרנט (Gemini + Google Search) ----------
async function webAnswer(question: string) {
  const now = localParts();
  const j = await geminiCall({
    systemInstruction: {
      parts: [{
        text: `אתה העוזר האישי של דניאל. עכשיו: ${now.full} (${heWeekday()}), שעון ישראל. דניאל גר ב${HOME_AREA}.
ענה בעברית, קצר (עד 6 שורות), רק לפי מה שמצאת בחיפוש.
שעות פתיחה: ציין את הסניף, והאם פתוח עכשיו לפי היום והשעה. טלפון: כתוב את המספר המלא.
אם המידע לא ודאי או סותר – אמור את זה ותמליץ להתקשר לוודא. בלי הקדמות.`,
      }],
    },
    contents: [{ role: "user", parts: [{ text: question }] }],
    tools: [{ google_search: {} }],
    generationConfig: { temperature: 0.2 },
  });
  const answer = geminiText(j) || "לא מצאתי תשובה ברורה.";
  const chunks = (j.candidates?.[0]?.groundingMetadata?.groundingChunks ?? [])
    .map((c: any) => c.web).filter((w: any) => w?.uri);
  const seen = new Set<string>();
  const sources = chunks.filter((w: any) => {
    const k = w.title ?? w.uri;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  }).slice(0, 2);
  let out = `🔎 ${answer}`;
  if (sources.length) out += "\n\nמקור: " + sources.map((w: any) => w.title ?? "").filter(Boolean).join(", ");
  return out;
}

// ---------- אנשי קשר / חיוג / מייל ----------
async function findContacts(query: string) {
  if (!query) return [];
  const res = await gas("searchContacts", { query, max: 5 });
  return (res.items ?? []) as { name: string; phones: string[]; emails: string[] }[];
}

function formatContacts(items: any[], forCall = false) {
  return items.map((c) => {
    const lines = [`👤 ${c.name || "(ללא שם)"}`];
    c.phones?.forEach((p: string) => lines.push(`📞 ${dialable(p)}`));
    if (!forCall) c.emails?.forEach((m: string) => lines.push(`✉️ ${m}`));
    return lines.join("\n");
  }).join("\n\n");
}

const PENDING_EMAIL = "pending_email";
const PENDING_TTL_MS = 30 * 60 * 1000;

async function getPendingEmail() {
  const { data } = await db.from("bot_state").select("value").eq("key", PENDING_EMAIL).maybeSingle();
  if (!data?.value) return null;
  try {
    const p = JSON.parse(data.value);
    if (Date.now() - new Date(p.at).getTime() > PENDING_TTL_MS) return null;
    return p;
  } catch { return null; }
}
const clearPendingEmail = () => db.from("bot_state").delete().eq("key", PENDING_EMAIL);

async function draftEmail(intent: any) {
  let to = String(intent.to ?? "").trim();
  let name = String(intent.query ?? "").trim();
  if (!/@/.test(to)) {
    const found = (await findContacts(name)).filter((c) => c.emails?.length);
    if (!found.length) {
      await sendText(OWNER_PHONE, `✉️ לא מצאתי כתובת מייל של "${name}" באנשי הקשר.\nכתוב שוב עם הכתובת, למשל: תשלח מייל ל-name@gmail.com ש...`);
      return;
    }
    const all = found.flatMap((c) => c.emails.map((m) => ({ name: c.name, email: m })));
    if (all.length > 1) {
      const lines = all.slice(0, 5).map((x) => `• ${x.name} – ${x.email}`);
      await sendText(OWNER_PHONE, `✉️ מצאתי כמה כתובות. כתוב שוב את הבקשה עם הכתובת הנכונה:\n${lines.join("\n")}`);
      return;
    }
    to = all[0].email;
    name = all[0].name || name;
  }
  const pending = { to, name, subject: intent.subject || "(ללא נושא)", body: intent.message || "", at: new Date().toISOString() };
  await db.from("bot_state").upsert({ key: PENDING_EMAIL, value: JSON.stringify(pending) });
  await sendText(OWNER_PHONE,
    `✉️ טיוטת מייל – עוד לא נשלח\nאל: ${name ? name + " " : ""}<${to}>\nנושא: ${pending.subject}\n\n${pending.body}\n\n` +
    `לשליחה כתוב *שלח*. לביטול כתוב *בטל*. (בתוקף 30 דקות)`);
}

// ---------- משימות (גיליון גוגל דרך Apps Script) ----------
const normArea = (a: string) => (AREAS.includes(a) ? a : "אחר");

// הוספה שקטה (אוטומטית מהבוט) – לא מפילה את הפעולה הראשית אם נכשלה
async function autoTask(p: Record<string, unknown>) {
  try { await gas("addTask", p); } catch (e) { console.error("autoTask", e); }
}

function taskLine(t: any) {
  let l = `• ${t.title}`;
  if (t.status === "progress") l += " 🔄";
  if (t.contact) l += ` · 👤 ${t.contact}`;
  if (t.due) l += ` · עד ${heShort(new Date(String(t.due).slice(0, 10) + "T12:00:00Z"))}`;
  if (Number(t.weeks) >= 2) l += ` · ⚠️ ${t.weeks} שבועות`;
  return l;
}

function formatTasks(items: any[], header: string) {
  if (!items.length) return `${header}\nאין משימות פתוחות 🙂`;
  const waiting = items.filter((t) => t.owner);
  const mine = items.filter((t) => !t.owner);
  const out = [header];
  for (const a of AREAS) {
    const g = mine.filter((t) => (t.area || "אחר") === a);
    if (g.length) out.push(`\n*${a}*\n` + g.map(taskLine).join("\n"));
  }
  if (waiting.length) out.push(`\n*⏳ מחכה לאחרים*\n` + waiting.map((t) => taskLine(t) + ` (${t.owner})`).join("\n"));
  return out.join("\n");
}

async function addTaskFromIntent(intent: any) {
  let phone = String(intent.phone ?? "").trim();
  const contact = String(intent.contact ?? "").trim();
  let contactName = contact;
  if (contact && !phone) {
    try {
      const c = (await findContacts(contact)).find((x) => x.phones?.length);
      if (c) { phone = c.phones[0]; contactName = c.name || contact; }
    } catch (e) { console.error("task contact", e); }
  }
  const res = await gas("addTask", {
    title: intent.title, area: normArea(intent.area), contact: contactName, phone,
    owner: String(intent.owner ?? "").trim(), due: /^\d{4}-\d{2}-\d{2}$/.test(intent.due ?? "") ? intent.due : "",
    source: "daniel",
  });
  const t = res.task ?? {};
  let reply = `📝 משימה נוספה: ${t.title}\n📂 ${t.area}`;
  if (t.owner) reply += `\n⏳ באחריות: ${t.owner}`;
  if (t.contact) reply += `\n👤 ${t.contact}${t.phone ? ` ${dialable(t.phone)}` : " (לא מצאתי מספר)"}`;
  if (t.due) reply += `\n📅 עד ${heShort(new Date(t.due + "T12:00:00Z"))}`;
  await sendText(OWNER_PHONE, reply);
}

// ---------- טיפול בהודעה נכנסת ----------
async function handleIncoming(msg: any) {
  if (msg.from !== OWNER_PHONE) return; // הבוט עונה רק לדניאל

  const { error: dupErr } = await db.from("processed_messages").insert({ id: msg.id });
  if (dupErr?.code === "23505") return;
  if (dupErr) console.error("dedupe", dupErr);

  // קובץ / תמונה / וידאו → חשבונית (פייפרלס) או שמירה בדרייב
  if (["image", "document", "video"].includes(msg.type)) {
    try {
      await saveIncomingMedia(msg);
    } catch (e) {
      console.error(e);
      await sendText(OWNER_PHONE, "😕 לא הצלחתי לשמור את הקובץ בדרייב.");
    }
    return;
  }
  if (msg.type !== "text") {
    await sendText(OWNER_PHONE, "כרגע אני מבין טקסט, תמונות וקבצים ✍️");
    return;
  }

  const text = msg.text.body.trim();

  // "מקליד..." בוואטסאפ מיד, כדי שתדע שההודעה התקבלה
  waPost({ messaging_product: "whatsapp", status: "read", message_id: msg.id, typing_indicator: { type: "text" } })
    .catch((e) => console.error("typing", e));

  // אישור / ביטול של מייל שממתין לשליחה
  if (/^(שלח|תשלח|כן|כן שלח|שלח אותו|אשר)[.!]?$/.test(text) || /^(בטל|לא|אל תשלח|ביטול)[.!]?$/.test(text)) {
    const pending = await getPendingEmail();
    if (pending) {
      if (/^(בטל|לא|אל תשלח|ביטול)/.test(text)) {
        await clearPendingEmail();
        await sendText(OWNER_PHONE, "🗑 המייל בוטל ולא נשלח.");
        return;
      }
      try {
        await gas("sendMail", { to: pending.to, subject: pending.subject, body: pending.body });
        await clearPendingEmail();
        await sendText(OWNER_PHONE, `✅ המייל נשלח ל-${pending.name || pending.to}.`);
        await autoTask({ title: `מייל ל${pending.name || pending.to}: ${pending.subject}`, area: "אחר", source: "email", status: "done", contact: pending.name || "" });
      } catch (e) {
        console.error(e);
        await sendText(OWNER_PHONE, "😕 לא הצלחתי לשלוח את המייל. הטיוטה נשמרה – נסה שוב לכתוב *שלח* בעוד רגע.");
      }
      return;
    }
  }

  // בחירה מתוך רשימת קבצים קודמת: "2" / "שלח 2"
  const pick = text.match(/^(?:שלח\s*)?(\d{1,2})$/);
  if (pick) {
    const { data: st } = await db.from("bot_state").select("value").eq("key", "last_drive").maybeSingle();
    const list = st?.value ? JSON.parse(st.value) : [];
    const f = list[Number(pick[1]) - 1];
    if (f) {
      try { await sendDriveFile(f); } catch (e) { console.error(e); await sendText(OWNER_PHONE, "😕 לא הצלחתי לשלוף את הקובץ."); }
      return;
    }
  }

  // שאלה כללית על חשבוניות ("מה עם חשבוניות?") – ישר לסטטוס, בלי לנחש
  if (/חשבוני/.test(text) && !/שלח|תמצא|מצא|קובץ|pdf|דרייב|מייל|משימ|להעביר|צריך|לא לשכוח|סיימתי/i.test(text)) {
    try {
      await sendText(OWNER_PHONE, formatInvoices(await gas("invoiceStats", { mark: false }), true));
    } catch (e) {
      console.error(e);
      await sendText(OWNER_PHONE, "😕 לא הצלחתי לבדוק את החשבוניות כרגע.");
    }
    return;
  }

  let intent: any;
  try {
    intent = await understand(text);
  } catch (e) {
    console.error(e);
    await sendText(OWNER_PHONE, geminiErrorReply(e));
    return;
  }

  try {
    switch (intent.action) {
      case "create": {
        if (!intent.start) break;
        const start = localToUtc(intent.start);
        if (start.getTime() <= Date.now()) {
          await sendText(OWNER_PHONE, `⚠️ הזמן שהבנתי (${heDateTime(start)}) כבר עבר. נסה לנסח שוב עם יום ושעה.`);
          return;
        }
        const dur = Number(intent.duration_minutes) || 15;
        const before = Math.max(0, Number(intent.remind_minutes_before) || 0);
        let eventId: string | null = null;
        try { eventId = await createEvent(intent.title, start, dur); } catch (e) { console.error(e); }
        let remindAt = new Date(start.getTime() - before * 60000);
        if (remindAt.getTime() < Date.now()) remindAt = new Date();
        const { error } = await db.from("reminders").insert({
          title: intent.title, event_start: start.toISOString(), remind_at: remindAt.toISOString(), calendar_event_id: eventId,
        });
        if (error) console.error("reminders insert", error);
        await autoTask({
          title: intent.title, area: normArea(intent.area), source: "calendar",
          due: localParts(start).date, ref: eventId ? "cal:" + eventId : "rem:" + start.toISOString() + intent.title,
        });
        let reply = `✅ ${intent.title}\n🗓 ${heDateTime(start)}`;
        reply += before ? `\n⏰ אזכיר לך ${before} דקות לפני` : `\n⏰ אזכיר לך בזמן`;
        if (!eventId) reply += "\n⚠️ לא הצלחתי להוסיף ליומן גוגל, אבל התזכורת נשמרה";
        await sendText(OWNER_PHONE, reply);
        return;
      }
      case "agenda": {
        const date = intent.date || localParts().date;
        await sendText(OWNER_PHONE, formatAgenda(await listEvents(date), `📅 ${heDay(date)}:`));
        return;
      }
      case "mail_unread": {
        const res = await gas("unread", { max: 8 });
        await sendText(OWNER_PHONE, formatMails(res.items, `📧 ${res.total} מיילים שלא נקראו (יומיים אחרונים):`));
        return;
      }
      case "mail_search": {
        const res = await gas("searchMail", { query: intent.query || text, max: 5 });
        await sendText(OWNER_PHONE, formatMails(res.items, `🔎 תוצאות חיפוש במייל:`, true));
        return;
      }
      case "drive_search": {
        const res = await gas("searchDrive", { query: intent.query || text, max: 5 });
        const items = res.items ?? [];
        if (!items.length) {
          await sendText(OWNER_PHONE, `📁 לא מצאתי בדרייב קובץ עם "${intent.query}". נסה מילים אחרות משם הקובץ.`);
        } else if (items.length === 1) {
          await sendDriveFile(items[0]);
        } else {
          await db.from("bot_state").upsert({ key: "last_drive", value: JSON.stringify(items) });
          const lines = items.map((f: any, i: number) => `${i + 1}. ${f.name}`);
          await sendText(OWNER_PHONE, `📁 מצאתי כמה קבצים. שלח מספר:\n${lines.join("\n")}`);
        }
        return;
      }
      case "invoices": {
        const st = await gas("invoiceStats", { mark: false });
        await sendText(OWNER_PHONE, formatInvoices(st, true));
        return;
      }
      case "wa_draft": {
        const phone = normPhone(intent.phone);
        if (phone.length < 11) {
          await sendText(OWNER_PHONE, "📱 צריך מספר טלפון כדי להכין את ההודעה. למשל: תכין הודעה ל-0501234567: אני מגיע ב-5");
          return;
        }
        const link = `https://wa.me/${phone}?text=${encodeURIComponent(intent.message ?? "")}`;
        await sendText(OWNER_PHONE, `✉️ ההודעה מוכנה:\n"${intent.message}"\n\nלחץ על הקישור, והיא תיפתח אצלך בוואטסאפ מוכנה לשליחה:\n${link}`);
        return;
      }
      case "web": {
        await sendText(OWNER_PHONE, await webAnswer(intent.query || text));
        return;
      }
      case "contact": {
        const found = await findContacts(intent.query || text);
        if (!found.length) {
          await sendText(OWNER_PHONE, `👤 לא מצאתי את "${intent.query}" באנשי הקשר של גוגל.\nאם זה עסק או מוסד, שאל למשל: מה הטלפון של מרפאת כללית בנהריה?`);
        } else {
          await sendText(OWNER_PHONE, formatContacts(found));
        }
        return;
      }
      case "call": {
        if (intent.phone && normPhone(intent.phone).length >= 11) {
          await sendText(OWNER_PHONE, `📞 לחץ על המספר כדי לחייג:\n${dialable(intent.phone)}`);
          return;
        }
        const found = (await findContacts(intent.query || text)).filter((c) => c.phones?.length);
        if (!found.length) {
          await sendText(OWNER_PHONE, `📞 לא מצאתי מספר של "${intent.query}" באנשי הקשר.\nאם זה עסק, שאל: מה הטלפון של ${intent.query}?`);
        } else {
          await sendText(OWNER_PHONE, `📞 לחץ על המספר כדי לחייג:\n\n${formatContacts(found.slice(0, 3), true)}`);
        }
        return;
      }
      case "email": {
        await draftEmail(intent);
        return;
      }
      case "task": {
        if (!intent.title) break;
        await addTaskFromIntent(intent);
        return;
      }
      case "tasks_list": {
        const res = await gas("listTasks", { query: intent.query || "", base: GAS_URL });
        const head = intent.query ? `📋 משימות פתוחות – ${intent.query} (${res.total}):` : `📋 משימות פתוחות (${res.total}):`;
        let txt = formatTasks(res.items ?? [], head);
        if (res.total > (res.items ?? []).length) txt += `\n…ועוד ${res.total - (res.items ?? []).length}`;
        await sendText(OWNER_PHONE, `${txt}\n\n✏️ לעדכון וסימון:\n${res.url}`);
        return;
      }
      case "task_done": {
        const res = await gas("doneTask", { query: intent.query || text });
        if (res.task) {
          await sendText(OWNER_PHONE, `✅ סומן כבוצע: ${res.task.title}`);
        } else if (res.matches?.length) {
          await sendText(OWNER_PHONE, `🤔 מצאתי כמה משימות מתאימות. כתוב שוב עם מילה מדויקת יותר:\n${res.matches.map(taskLine).join("\n")}`);
        } else {
          await sendText(OWNER_PHONE, `🤔 לא מצאתי משימה פתוחה עם "${intent.query}". אפשר לכתוב "מה המשימות שלי" לרשימה.`);
        }
        return;
      }
    }
  } catch (e) {
    console.error(e);
    const s = String(e);
    const why = s.includes("GAS not configured") ? " (ג'ימייל/דרייב עוד לא מחוברים)"
      : s.includes("unknown action") ? " (פעולה שעוד לא הותקנה ב-Apps Script)"
      : s.includes("Gemini") ? " (Gemini לא זמין כרגע)" : "";
    await sendText(OWNER_PHONE, `😕 משהו השתבש${why}. נסה שוב בעוד רגע.`);
    return;
  }

  await sendText(OWNER_PHONE, intent.reply || HELP);
}

// ---------- משימה מתוזמנת (כל דקה) ----------
async function tick() {
  const { data } = await db.from("reminders").select("*")
    .eq("status", "pending").lte("remind_at", new Date().toISOString()).limit(20);

  for (const r of data ?? []) {
    const { data: claimed } = await db.from("reminders").update({ status: "sending" })
      .eq("id", r.id).eq("status", "pending").select();
    if (!claimed?.length) continue;
    const when = r.event_start ? heTime(new Date(r.event_start)) : "";
    try {
      await notify(`⏰ תזכורת: ${r.title}${when ? `\n🕒 ${when}` : ""}`, `${r.title}${when ? ` (${when})` : ""}`);
      await db.from("reminders").update({ status: "sent" }).eq("id", r.id);
    } catch (e) {
      console.error(e);
      await db.from("reminders").update({ status: "failed", error: String(e).slice(0, 500) }).eq("id", r.id);
    }
  }

  // דוח משימות שבועי — יום שישי, פעם אחת, בין WEEKLY_TIME ל-20:00
  {
    const { date, time } = localParts();
    if (dowIL() === 5 && time >= WEEKLY_TIME && time < "20:00" && GAS_URL) {
      const { data: st } = await db.from("bot_state").select("value").eq("key", "last_weekly").maybeSingle();
      if (st?.value !== date) {
        await db.from("bot_state").upsert({ key: "last_weekly", value: date });
        try { await weeklyReport(date); } catch (e) { console.error("weekly", e); }
      }
    }
  }

  // סיכום בוקר — פעם ביום, בין SUMMARY_TIME ל-12:00
  const { date, time } = localParts();
  if (time >= SUMMARY_TIME && time < "12:00") {
    const { data: st } = await db.from("bot_state").select("value").eq("key", "last_summary").maybeSingle();
    if (st?.value !== date) {
      await db.from("bot_state").upsert({ key: "last_summary", value: date });
      try {
        const items = await listEvents(date);
        let txt = formatAgenda(items, `☀️ בוקר טוב דניאל! ${heDay(date)}:`);
        if (GAS_URL) {
          try {
            const mail = await gas("unread", { max: 5 });
            if (mail.total) txt += "\n\n" + formatMails(mail.items, `📧 ${mail.total} מיילים שלא נקראו:`);
          } catch (e) { console.error(e); }
          try {
            const inv = formatInvoices(await gas("invoiceStats", { mark: true }), false);
            if (inv) txt += "\n\n" + inv;
          } catch (e) { console.error(e); }
          try {
            const mt = await morningTasks(date);
            if (mt) txt += "\n\n" + mt;
          } catch (e) { console.error("morning tasks", e); }
        }
        await notify(txt, `סיכום הבוקר – ${items.length} אירועים היום`);
      } catch (e) {
        console.error(e);
      }
    }
  }
}

// ---------- משימות בסיכום הבוקר ----------
// ראשון: סגירת השבוע הקודם (מה שלא הסתיים עובר הלאה) + רשימה מלאה. שאר הימים: מספר + מה שמתוכנן להיום.
async function morningTasks(date: string) {
  if (dowIL() === 0) {
    const { data: st } = await db.from("bot_state").select("value").eq("key", "last_rollover").maybeSingle();
    if (st?.value !== date) {
      await db.from("bot_state").upsert({ key: "last_rollover", value: date });
      await gas("closeWeek");
    }
    const res = await gas("listTasks", { base: GAS_URL });
    return formatTasks(res.items ?? [], `📋 משימות פתוחות לשבוע (${res.total}):`) + `\n✏️ ${res.url}`;
  }
  const res = await gas("listTasks", { base: GAS_URL });
  const items = (res.items ?? []) as any[];
  if (!res.total) return "";
  const today = items.filter((t) => String(t.due).slice(0, 10) === date);
  let txt = `📋 ${res.total} משימות פתוחות`;
  if (today.length) txt += ` · להיום:\n` + today.map(taskLine).join("\n");
  return txt;
}

// ---------- דוח שבועי (שישי) ----------
async function weeklyReport(date: string) {
  // אירועי היומן של השבוע (ראשון–שבת) נכנסים כמשימות, בלי כפילויות
  const start = addDays(date, -dowIL());
  for (let i = 0; i < 7; i++) {
    const day = addDays(start, i);
    try {
      for (const ev of await listEvents(day)) {
        if (!ev.id || !ev.summary) continue;
        await autoTask({ title: ev.summary, area: "אחר", source: "calendar", due: day, ref: "cal:" + ev.id });
      }
    } catch (e) { console.error("weekly events", day, e); }
  }
  const { counts: c, url } = await gas("weeklySummary", { base: GAS_URL });
  const lines = [
    `🗓 סגירת שבוע – ${heDay(date)}`,
    `📋 ${c.total} משימות: ✅ ${c.done} בוצעו · 🔄 ${c.progress} בתהליך · ⬜ ${c.open + c.postponed} פתוחות`,
  ];
  if (c.waiting) lines.push(`⏳ ${c.waiting} מחכות לאחרים`);
  if (c.stuck) lines.push(`⚠️ ${c.stuck} תקועות שבועיים ומעלה`);
  lines.push(`\n✏️ לסימון ועדכון (צ'קבוקס, סטטוס והערות):\n${url}`);
  await notify(lines.join("\n"), `סגירת שבוע: ${c.total} משימות, ${c.done} בוצעו. לסימון: ${url}`);
}

// ---------- נקודת הכניסה ----------
Deno.serve(async (req) => {
  const url = new URL(req.url);

  if (req.method === "GET") {
    if (url.searchParams.get("hub.mode") === "subscribe" && url.searchParams.get("hub.verify_token") === VERIFY_TOKEN) {
      return new Response(url.searchParams.get("hub.challenge") ?? "", { status: 200 });
    }
    return new Response("Forbidden", { status: 403 });
  }
  if (req.method !== "POST") return new Response("Method not allowed", { status: 405 });

  if (url.searchParams.get("task") === "tick") {
    if (!CRON_SECRET || req.headers.get("x-cron-secret") !== CRON_SECRET) {
      return new Response("Unauthorized", { status: 401 });
    }
    await tick();
    return new Response("ok");
  }

  const raw = await req.text();
  if (APP_SECRET && !(await verifySignature(raw, req.headers.get("x-hub-signature-256")))) {
    return new Response("Bad signature", { status: 401 });
  }
  let payload: any;
  try {
    payload = JSON.parse(raw);
  } catch {
    return new Response("ok");
  }
  const msgs = (payload.entry ?? []).flatMap((e: any) =>
    (e.changes ?? []).flatMap((c: any) => c.value?.messages ?? [])
  );
  const work = Promise.all(msgs.map((m: any) => handleIncoming(m).catch((e) => console.error("handle", e))));
  // @ts-ignore EdgeRuntime קיים ב-Supabase
  if (typeof EdgeRuntime !== "undefined") EdgeRuntime.waitUntil(work);
  else await work;
  return new Response("ok");
});
