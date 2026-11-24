/**
 * VESC közös edzések naptár – Google Apps Script backend
 * Munkalapok:
 *   "Edzések" – egy sor = egy edzés (automatikusan jön létre)
 *   "Tagok"   – Név | E-mail (a rendszergazda tölti ki, ide lehet meghívót küldeni)
 *   "Versenyek" – Dátum | Verseny neve | Sportág | Táv | Helyszín | Link | Egyéb
 *                 (előre felvett versenyek a legördülő listához; a Sportág lehet:
 *                 Triatlon, Duatlon, Futóverseny, Kerékpárverseny, Egyéb)
 * Rendszergazda jelszó: a táblázat "VESC naptár" menüjében állítható be.
 */
const SHEET_NAME = 'Edzések';
const MEMBERS_SHEET = 'Tagok';
const RACES_SHEET = 'Versenyek';
const RACES_HEAD = ['Dátum', 'Verseny neve', 'Sportág', 'Táv', 'Helyszín', 'Link', 'Egyéb'];
const HEAD = ['id', 'sport', 'title', 'date', 'start', 'end', 'place', 'distance', 'desc', 'organizer', 'attendees', 'createdAt', 'updatedAt', 'ownerHash', 'link', 'seriesId', 'raceType', 'endDate', 'gcalId', 'gcalSig'];
const COL = {}; HEAD.forEach((h, i) => (COL[h] = i));
const SPORTS = ['run', 'bike', 'swim', 'other', 'race', 'social'];
const SPORT_NAMES = { run: 'Futás', bike: 'Kerékpározás', swim: 'Úszás', other: 'Erőnlét / Egyéb', race: 'Verseny', social: 'Közösségi program' };
const SPORT_EMOJI = { run: '🏃‍♂️', bike: '🚴‍♂️', swim: '🏊‍♂️', other: '🏋️‍♂️', race: '🏆', social: '🎉' };
const RACE_TYPES = { tri: 'Triatlon', du: 'Duatlon', run: 'Futóverseny', bike: 'Kerékpárverseny', other: 'Egyéb verseny' };
const TZ = 'Europe/Budapest';
const SENDER_NAME = 'VESC Triatlon';
const MAX_REPEAT = 53; // legfeljebb ~1 év heti ismétlés

/* ---------- nyilvános API (a külső weboldal ezen keresztül éri el a naptárt) ---------- */
const API = {
  getData: getData, listMembers: listMembers, listRaces: listRaces, checkAdmin: checkAdmin,
  saveEvent: saveEvent, deleteEvent: deleteEvent, setAttendance: setAttendance, sendInvite: sendInvite,
  getCalendarInfo: getCalendarInfo,
};
function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  return api_(body.fn, body.args);
}
function api_(fn, args) {
  let out;
  try {
    if (typeof fn !== 'string' || !Object.prototype.hasOwnProperty.call(API, fn)) throw new Error('Ismeretlen művelet.');
    out = { ok: true, result: API[fn].apply(null, Array.isArray(args) ? args : []) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (e && e.parameter && e.parameter.ping) {
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('VESC közös edzések naptár')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- menü a táblázatban ---------- */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('VESC naptár')
    .addItem('Munkalapok előkészítése', 'setup')
    .addItem('Rendszergazda jelszó beállítása', 'setAdminPassword')
    .addItem('Naptár nyilvános címének beállítása', 'setPublicUrl')
    .addSeparator()
    .addItem('Google Naptár beállítása / linkek', 'setupCalendar')
    .addItem('Google Naptár szinkronizálása most', 'syncAllMenu')
    .addToUi();
}

function setup() {
  const sh = sheet_();
  sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold').setBackground('#1b1f24').setFontColor('#ffffff');
  const ms = membersSheet_();
  ms.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1b1f24').setFontColor('#ffffff');
  const rs = racesSheet_();
  const n = Math.max(rs.getMaxRows() - 1, 1);
  rs.getRange(1, 1, 1, RACES_HEAD.length).setFontWeight('bold').setBackground('#1b1f24').setFontColor('#ffffff');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Triatlon', 'Duatlon', 'Futóverseny', 'Kerékpárverseny', 'Egyéb'], true).setAllowInvalid(true).build();
  rs.getRange(2, 2, n, 1).clearDataValidations();
  rs.getRange(2, 3, n, 1).setDataValidation(rule);
  rs.getRange(2, 1, n, 1).setNumberFormat('yyyy-mm-dd')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build());
  rs.autoResizeColumns(1, RACES_HEAD.length);
  SpreadsheetApp.getActiveSpreadsheet().toast('Kész. A "Tagok" lapra a tagokat, a "Versenyek" lapra az előre felvett versenyeket írd.');
}

function setAdminPassword() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Rendszergazda jelszó', 'Adj meg egy legalább 6 karakteres jelszót. Ezzel lehet a naptárban bármelyik edzést szerkeszteni vagy törölni.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const pw = r.getResponseText().trim();
  if (pw.length < 6) { ui.alert('Túl rövid, legalább 6 karakter kell.'); return; }
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', pw);
  ui.alert('A rendszergazda jelszó elmentve.');
}

/** A meghívókban szereplő naptár-link (pl. a GitHub Pages cím). */
function setPublicUrl() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Naptár nyilvános címe', 'Írd be a naptár weboldalának címét (https://…). Üresen hagyva a script.google.com-os linket használja.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const v = r.getResponseText().trim();
  const props = PropertiesService.getScriptProperties();
  if (!v) { props.deleteProperty('WEBAPP_URL'); ui.alert('Törölve, a meghívók a script.google.com linket használják.'); return; }
  if (!/^https:\/\/\S+$/.test(v)) { ui.alert('A címnek https://-sel kell kezdődnie.'); return; }
  props.setProperty('WEBAPP_URL', v);
  ui.alert('Elmentve. A meghívókban mostantól ez a cím szerepel.');
}

/* ---------- munkalapok ---------- */
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.setFrozenRows(1);
  }
  const head = sh.getRange(1, 1, 1, HEAD.length).getDisplayValues()[0];
  if (head.join('|') !== HEAD.join('|')) sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]);
  sh.getRange(1, 1, sh.getMaxRows(), HEAD.length).setNumberFormat('@'); // minden szöveg
  return sh;
}
function membersSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(MEMBERS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MEMBERS_SHEET);
    sh.getRange(1, 1, 1, 2).setValues([['Név', 'E-mail']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function racesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(RACES_SHEET);
  if (!sh) {
    sh = ss.insertSheet(RACES_SHEET);
    sh.getRange(1, 1, 1, RACES_HEAD.length).setValues([RACES_HEAD]);
    sh.setFrozenRows(1);
    return sh;
  }
  // régi elrendezés (Név | Típus | Helyszín | Link) átalakítása az újra
  if (sh.getRange(1, 1).getDisplayValue().trim() === 'Név') {
    const last = sh.getLastRow();
    const old = last >= 2 ? sh.getRange(2, 1, last - 1, 4).getDisplayValues() : [];
    sh.getRange(1, 1, Math.max(last, 1), Math.max(sh.getLastColumn(), RACES_HEAD.length)).clearContent().clearDataValidations();
    sh.getRange(1, 1, 1, RACES_HEAD.length).setValues([RACES_HEAD]);
    if (old.length) sh.getRange(2, 1, old.length, RACES_HEAD.length)
      .setValues(old.map(r => ['', r[0], r[1], '', r[2], r[3], '']));
  }
  return sh;
}
function raceTypeKey_(label) {
  const l = String(label || '').trim().toLowerCase();
  if (!l) return '';
  if (RACE_TYPES[l]) return l;
  if (l.indexOf('triatl') === 0) return 'tri';
  if (l.indexOf('duatl') === 0) return 'du';
  if (l.indexOf('fut') === 0) return 'run';
  if (l.indexOf('kerékp') === 0 || l.indexOf('kerekp') === 0 || l.indexOf('bring') === 0) return 'bike';
  return 'other';
}

/** Előre felvett versenyek a legördülő listához. */
function listRaces() {
  const sh = racesSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const n = RACES_HEAD.length;
  const raw = sh.getRange(2, 1, last - 1, n).getValues();
  const disp = sh.getRange(2, 1, last - 1, n).getDisplayValues();
  return disp.map((r, i) => {
    const d = raw[i][0];
    let date = '';
    if (d instanceof Date && !isNaN(d)) date = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
    else {
      const m = String(r[0]).match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
      if (m) date = m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
    }
    const link = String(r[5]).trim();
    return {
      date: date, name: String(r[1]).trim().slice(0, 80), type: raceTypeKey_(r[2]),
      distance: String(r[3]).trim().slice(0, 60), place: String(r[4]).trim().slice(0, 100),
      link: /^https?:\/\/\S+$/i.test(link) ? link.slice(0, 300) : '', note: String(r[6]).trim().slice(0, 1000),
    };
  }).filter(r => r.name)
    .sort((a, b) => (a.date || '9999').localeCompare(b.date || '9999'));
}

/* ---------- jogosultság ---------- */
function hash_(key) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'vesc:' + key, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(bytes);
}
function isAdmin_(auth) {
  const pw = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  return !!pw && !!auth && auth.adminPw === pw;
}
function validKey_(auth) { return !!auth && typeof auth.key === 'string' && /^[a-f0-9]{32,64}$/.test(auth.key); }
function canManage_(rowVals, auth) {
  if (isAdmin_(auth)) return true;
  const h = rowVals[COL.ownerHash];
  return !!h && validKey_(auth) && hash_(auth.key) === h;
}
function checkAdmin(pw) {
  return isAdmin_({ adminPw: String(pw || '') });
}

/* ---------- olvasás ---------- */
function getData(auth) {
  const sh = sheet_();
  const last = sh.getLastRow();
  const admin = isAdmin_(auth);
  const myHash = validKey_(auth) ? hash_(auth.key) : null;
  let events = [];
  if (last >= 2) {
    events = sh.getRange(2, 1, last - 1, HEAD.length).getDisplayValues()
      .filter(r => r[0])
      .map(r => {
        const o = {};
        HEAD.forEach((h, i) => { if (h !== 'ownerHash' && h !== 'gcalSig') o[h] = r[i]; });
        o.attendees = o.attendees ? o.attendees.split(';').map(s => s.trim()).filter(Boolean) : [];
        o.canManage = admin || (!!myHash && r[COL.ownerHash] === myHash);
        return o;
      });
  }
  return { events: events, isAdmin: admin };
}

/** Csak a nevek mennek ki a böngészőbe, az e-mail címek nem. */
function listMembers() {
  const sh = membersSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 2).getDisplayValues()
    .map((r, i) => ({ id: i + 2, name: r[0].trim(), hasEmail: /\S+@\S+\.\S+/.test(r[1]) }))
    .filter(m => m.name && m.hasEmail)
    .map(m => ({ id: m.id, name: m.name }));
}

/* ---------- írás ---------- */
/**
 * opts.repeatUntil: 'yyyy-mm-dd' – új edzésnél heti ismétlés eddig a napig
 * opts.scope: 'one' | 'future' – sorozat szerkesztésekor csak ez, vagy ez és a következők
 */
function saveEvent(ev, auth, opts) {
  opts = opts || {};
  return withLock_(() => {
    const clean = clean_(ev);
    const sh = sheet_();
    const now = new Date().toISOString();
    const row = ev.id ? findRow_(sh, String(ev.id)) : -1;
    if (ev.id && row < 1) throw new Error('Ez az edzés már nem létezik.');
    if (row > 0) {
      const cur = sh.getRange(row, 1, 1, HEAD.length).getDisplayValues()[0];
      if (!canManage_(cur, auth)) throw new Error('Ezt az edzést csak a létrehozója vagy a rendszergazda módosíthatja.');
      const targets = (opts.scope === 'future' && cur[COL.seriesId])
        ? seriesRows_(sh, cur[COL.seriesId], cur[COL.date]) : [{ row: row, vals: cur }];
      const shift = dayDiff_(cur[COL.date], clean.date);
      targets.forEach(t => {
        const v = t.vals;
        const date = t.row === row ? clean.date : addDays_(v[COL.date], shift);
        const out = HEAD.map(h => v[COL[h]]);
        ['sport', 'title', 'start', 'end', 'place', 'distance', 'desc', 'link', 'raceType'].forEach(h => (out[COL[h]] = clean[h]));
        out[COL.date] = date;
        out[COL.endDate] = clean.endDate ? addDays_(date, dayDiff_(clean.date, clean.endDate)) : '';
        out[COL.updatedAt] = now;
        if (opts.scope !== 'future' && t.row === row && targets.length === 1 && shift !== 0 && v[COL.seriesId]) {
          // egyetlen alkalom áthelyezése: kiválik a sorozatból
          out[COL.seriesId] = '';
        }
        sh.getRange(t.row, 1, 1, HEAD.length).setValues([out.map(safe_)]);
      });
      syncRows_(sh, targets.map(t => t.row));
    } else {
      if (!validKey_(auth)) throw new Error('Hiányzik a saját kulcs. Frissítsd az oldalt.');
      const organizer = str_(ev.organizer, 60);
      let dates = [clean.date];
      let seriesId = '';
      const until = String(opts.repeatUntil || '');
      if (clean.sport !== 'race' && /^\d{4}-\d{2}-\d{2}$/.test(until) && until > clean.date) {
        seriesId = Utilities.getUuid();
        dates = [];
        for (let d = clean.date; d <= until && dates.length < MAX_REPEAT; d = addDays_(d, 7)) dates.push(d);
      }
      const owner = hash_(auth.key);
      const rows = dates.map(d => {
        const out = HEAD.map(() => '');
        out[COL.id] = Utilities.getUuid();
        ['sport', 'title', 'start', 'end', 'place', 'distance', 'desc', 'link', 'raceType'].forEach(h => (out[COL[h]] = clean[h]));
        out[COL.date] = d;
        out[COL.endDate] = clean.endDate ? addDays_(d, dayDiff_(clean.date, clean.endDate)) : '';
        out[COL.organizer] = organizer;
        out[COL.attendees] = organizer;
        out[COL.createdAt] = now;
        out[COL.updatedAt] = now;
        out[COL.ownerHash] = owner;
        out[COL.seriesId] = seriesId;
        return out.map(safe_);
      });
      const first = sh.getLastRow() + 1;
      if (sh.getMaxRows() < first + rows.length) sh.insertRowsAfter(sh.getMaxRows(), rows.length + 10);
      sh.getRange(first, 1, rows.length, HEAD.length).setNumberFormat('@').setValues(rows);
      syncRows_(sh, rows.map((r, i) => first + i));
    }
  }, auth);
}

function deleteEvent(id, auth, scope) {
  return withLock_(() => {
    const sh = sheet_();
    const row = findRow_(sh, String(id));
    if (row < 1) return;
    const cur = sh.getRange(row, 1, 1, HEAD.length).getDisplayValues()[0];
    if (!canManage_(cur, auth)) throw new Error('Ezt az edzést csak a létrehozója vagy a rendszergazda törölheti.');
    const rows = (scope === 'future' && cur[COL.seriesId])
      ? seriesRows_(sh, cur[COL.seriesId], cur[COL.date]).map(t => t.row) : [row];
    rows.forEach(r => deleteCalCopy_(sh.getRange(r, COL.gcalId + 1).getDisplayValue()));
    rows.sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
  }, auth);
}

/** Jelentkezés / lemondás – ezt bárki megteheti a saját nevével. */
function setAttendance(id, name, going, auth) {
  name = str_(name, 60);
  if (!name) throw new Error('Hiányzik a név.');
  return withLock_(() => {
    const sh = sheet_();
    const row = findRow_(sh, String(id));
    if (row < 1) throw new Error('Ez az edzés már nem létezik.');
    const cell = sh.getRange(row, COL.attendees + 1);
    let list = cell.getDisplayValue().split(';').map(s => s.trim()).filter(Boolean);
    list = list.filter(n => n.toLowerCase() !== name.toLowerCase());
    if (going) list.push(name);
    cell.setValue(safe_(list.join('; ')));
  }, auth);
}

/* ---------- meghívó ---------- */
function sendInvite(id, memberIds, extraEmails, note, inviterName, auth) {
  const sh = sheet_();
  const row = findRow_(sh, String(id));
  if (row < 1) throw new Error('Ez az edzés már nem létezik.');
  const cur = sh.getRange(row, 1, 1, HEAD.length).getDisplayValues()[0];
  if (!canManage_(cur, auth)) throw new Error('Meghívót csak az edzés létrehozója vagy a rendszergazda küldhet.');

  const emails = {};
  const ms = membersSheet_();
  (memberIds || []).slice(0, 200).forEach(r => {
    r = Number(r);
    if (r >= 2 && r <= ms.getLastRow()) {
      const e = ms.getRange(r, 2).getDisplayValue().trim().toLowerCase();
      if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)) emails[e] = true;
    }
  });
  String(extraEmails || '').split(/[\s,;]+/).forEach(e => {
    e = e.trim().toLowerCase();
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)) emails[e] = true;
  });
  const list = Object.keys(emails);
  if (!list.length) throw new Error('Nincs érvényes e-mail cím kiválasztva.');
  if (list.length > 50) throw new Error('Egyszerre legfeljebb 50 címzett lehet.');
  if (MailApp.getRemainingDailyQuota() < list.length) throw new Error('Mára elfogyott a Gmail napi küldési kerete, próbáld holnap.');

  const ev = {};
  HEAD.forEach((h, i) => (ev[h] = cur[i]));
  const url = webAppUrl_();
  const kindText = ev.sport === 'race' ? 'egy versenyre' : ev.sport === 'social' ? 'egy közösségi programra' : 'egy közös edzésre';
  const sportName = ev.sport === 'race' ? (RACE_TYPES[ev.raceType] || 'Verseny') : (SPORT_NAMES[ev.sport] || SPORT_NAMES.other);
  const emoji = SPORT_EMOJI[ev.sport] || SPORT_EMOJI.other;
  const title = ev.title || sportName;
  const when = (ev.endDate ? dayLong_(ev.date) + ' – ' + dayLong_(ev.endDate) : dayLong_(ev.date)) +
    (ev.start ? ', ' + ev.start + (ev.end ? '–' + ev.end : '') : (ev.endDate ? '' : ', egész nap'));
  const inviter = str_(inviterName, 60) || ev.organizer || 'A VESC';
  note = String(note || '').trim().slice(0, 500);

  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const rows = [['📅', when], ev.place && ['📍', ev.place], ev.distance && ['📏', ev.distance], ev.link && ['🔗', ev.link]].filter(Boolean)
    .map(r => `<tr><td style="padding:4px 10px 4px 0;vertical-align:top">${r[0]}</td><td style="padding:4px 0">${esc(r[1])}</td></tr>`).join('');
  const html =
    `<div style="font-family:Arial,sans-serif;font-size:15px;color:#1b1f24;max-width:520px">
      <p>${esc(inviter)} meghívott ${kindText}:</p>
      <h2 style="margin:8px 0">${emoji} ${esc(title)}</h2>
      <p style="margin:0 0 8px;color:#6b7280">${esc(sportName)}</p>
      <table style="border-collapse:collapse">${rows}</table>
      ${ev.desc ? `<p style="white-space:pre-wrap;background:#f1f3f5;padding:10px 12px;border-radius:8px">${esc(ev.desc)}</p>` : ''}
      ${note ? `<p><em>${esc(note)}</em></p>` : ''}
      <p><a href="${url}" style="display:inline-block;background:#1b1f24;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none">Megnézem a naptárban és jelentkezem</a></p>
      <p style="color:#6b7280;font-size:12px">A csatolt meghívóval az edzést a saját naptáradba is felveheted.</p>
    </div>`;
  const plain = `${inviter} meghívott ${kindText}:\n\n${title} (${sportName})\n${when}` +
    (ev.place ? `\n${ev.place}` : '') + (ev.distance ? `\n${ev.distance}` : '') +
    (ev.desc ? `\n\n${ev.desc}` : '') + (note ? `\n\n${note}` : '') + `\n\nNaptár és jelentkezés: ${url}`;

  const ics = Utilities.newBlob(ics_(ev, title, url), 'text/calendar', 'edzes.ics');
  // mindenki külön levelet kap, csak a saját címére (a VESC fiók nem kap másolatot)
  let sent = 0;
  const failed = [];
  list.forEach(addr => {
    try {
      MailApp.sendEmail({
        to: addr,
        subject: `Meghívó: ${emoji} ${title} – ${when}`,
        body: plain,
        htmlBody: html,
        name: SENDER_NAME,
        attachments: [ics],
      });
      sent++;
    } catch (err) {
      failed.push(addr);
    }
  });
  if (!sent) throw new Error('A meghívót nem sikerült elküldeni.');
  return sent;
}

function ics_(ev, title, url) {
  const fmtUtc = d => Utilities.formatDate(d, 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  const e = s => String(s || '').replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n');
  let dt;
  if (ev.start) {
    const s = Utilities.parseDate(ev.date + ' ' + ev.start, TZ, 'yyyy-MM-dd HH:mm');
    const lastDay = ev.endDate || ev.date;
    const en = ev.end ? Utilities.parseDate(lastDay + ' ' + ev.end, TZ, 'yyyy-MM-dd HH:mm')
      : (ev.endDate ? Utilities.parseDate(lastDay + ' 23:59', TZ, 'yyyy-MM-dd HH:mm') : new Date(s.getTime() + 3600000));
    dt = ['DTSTART:' + fmtUtc(s), 'DTEND:' + fmtUtc(en)];
  } else {
    const d = ev.date.replace(/-/g, '');
    const next = addDays_(ev.endDate || ev.date, 1).replace(/-/g, '');
    dt = ['DTSTART;VALUE=DATE:' + d, 'DTEND;VALUE=DATE:' + next];
  }
  return ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//VESC//Kozos edzesek//HU', 'METHOD:PUBLISH',
    'BEGIN:VEVENT', 'UID:' + ev.id + '@vesc-naptar', 'DTSTAMP:' + fmtUtc(new Date())].concat(dt, [
    'SUMMARY:' + e(title), 'LOCATION:' + e(ev.place), 'DESCRIPTION:' + e([ev.distance, ev.desc, ev.link, url].filter(Boolean).join('\n')),
    'URL:' + url, 'END:VEVENT', 'END:VCALENDAR']).join('\r\n');
}

/* ---------- Google Naptár szinkron (egyirányú: VESC naptár → Google Naptár) ---------- */
const CAL_NAME = 'VESC közös edzések';

function cal_() {
  const id = PropertiesService.getScriptProperties().getProperty('CAL_ID');
  if (!id) return null;
  try { return CalendarApp.getCalendarById(id); } catch (e) { return null; }
}

/** Menü: naptár létrehozása (ha még nincs), teljes szinkron, óránkénti frissítés, linkek. */
function setupCalendar() {
  const ui = SpreadsheetApp.getUi();
  let cal = cal_();
  if (!cal) {
    cal = CalendarApp.createCalendar(CAL_NAME, { timeZone: TZ, summary: 'A VESC közös edzései, versenyei és programjai. Részletek és jelentkezés a VESC naptár oldalán.' });
    try { cal.setTimeZone(TZ); } catch (e) {}
    PropertiesService.getScriptProperties().setProperty('CAL_ID', cal.getId());
  }
  const n = syncAll();
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'syncAll')) {
    ScriptApp.newTrigger('syncAll').timeBased().everyHours(1).create();
  }
  const info = getCalendarInfo();
  ui.alert('Google Naptár kész',
    'Naptár: ' + CAL_NAME + '\nSzinkronizált események: ' + n + '\n\n' +
    'FONTOS: tedd nyilvánossá! Google Naptár → bal oldalt a VESC naptár ⋮ → Beállítások és megosztás → ' +
    'Hozzáférési engedélyek → „Nyilvánosan elérhetővé tétel” (Az összes eseményrészlet megtekintése).\n\n' +
    'Feliratkozás Google Naptárral:\n' + info.google + '\n\nOutlook / egyéb (iCal):\n' + info.ical,
    ui.ButtonSet.OK);
}
function syncAllMenu() {
  const n = syncAll();
  SpreadsheetApp.getActiveSpreadsheet().toast('Google Naptár szinkronizálva (' + n + ' esemény).');
}

/** Feliratkozási linkek a weboldalnak. */
function getCalendarInfo() {
  const id = PropertiesService.getScriptProperties().getProperty('CAL_ID');
  if (!id) return null;
  const ical = 'https://calendar.google.com/calendar/ical/' + encodeURIComponent(id) + '/public/basic.ics';
  return {
    name: CAL_NAME,
    google: 'https://calendar.google.com/calendar/r?cid=' + encodeURIComponent(id),
    ical: ical,
    webcal: ical.replace(/^https:/, 'webcal:'),
  };
}

/** Minden esemény átvezetése a Google Naptárba; a törölteket onnan is törli. Óránként is lefut. */
function syncAll() {
  const cal = cal_();
  if (!cal) return 0;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return 0;
  let n = 0;
  try {
    const sh = sheet_();
    const last = sh.getLastRow();
    const ids = {};
    if (last >= 2) {
      const all = sh.getRange(2, 1, last - 1, HEAD.length).getDisplayValues();
      all.forEach((v, i) => {
        if (!v[COL.id]) return;
        ids[v[COL.id]] = true;
        try { syncRow_(cal, sh, i + 2, v); n++; } catch (e) { console.warn('sync', v[COL.id], e); }
      });
    }
    // árva események törlése (a táblázatból már törölt sorok)
    const from = new Date(Date.now() - 180 * 86400000), to = new Date(Date.now() + 800 * 86400000);
    cal.getEvents(from, to).forEach(ev => {
      const tag = ev.getTag('vescId');
      if (tag && !ids[tag]) { try { ev.deleteEvent(); } catch (e) {} }
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return n;
}

function syncRows_(sh, rowNums) {
  const cal = cal_();
  if (!cal) return;
  rowNums.forEach(r => {
    try { syncRow_(cal, sh, r, sh.getRange(r, 1, 1, HEAD.length).getDisplayValues()[0]); }
    catch (e) { console.warn('syncRow', r, e); }
  });
}
function deleteCalCopy_(gcalId) {
  if (!gcalId) return;
  const cal = cal_();
  if (!cal) return;
  try { const ev = cal.getEventById(gcalId); if (ev) ev.deleteEvent(); } catch (e) {}
}

function calSig_(v) {
  const parts = ['sport', 'title', 'date', 'endDate', 'start', 'end', 'place', 'distance', 'desc', 'link', 'raceType'].map(h => v[COL[h]]);
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(parts) + '|v1', Utilities.Charset.UTF_8);
  return Utilities.base64Encode(bytes);
}

function syncRow_(cal, sh, row, v) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v[COL.date])) return;
  const sig = calSig_(v);
  let ev = null;
  if (v[COL.gcalId]) { try { ev = cal.getEventById(v[COL.gcalId]); } catch (e) { ev = null; } }
  if (ev && v[COL.gcalSig] === sig) return; // nem változott
  if (ev) { try { ev.deleteEvent(); } catch (e) {} }
  const ne = createCalEvent_(cal, v);
  sh.getRange(row, COL.gcalId + 1, 1, 2).setNumberFormat('@').setValues([[ne.getId(), sig]]);
}

function calTitle_(v) {
  const sp = v[COL.sport];
  const emoji = SPORT_EMOJI[sp] || SPORT_EMOJI.other;
  if (sp === 'race') return emoji + ' ' + (v[COL.title] || 'Verseny') + ' (' + (RACE_TYPES[v[COL.raceType]] || 'Verseny') + ')';
  if (sp === 'social') return emoji + ' ' + (v[COL.title] || 'Közösségi program');
  const name = SPORT_NAMES[sp] || SPORT_NAMES.other;
  return emoji + ' ' + name + (v[COL.title] ? ': ' + v[COL.title] : '');
}

function createCalEvent_(cal, v) {
  const date = v[COL.date];
  const lastDay = (v[COL.endDate] && v[COL.endDate] > date) ? v[COL.endDate] : date;
  const site = webAppUrl_();
  const desc = [v[COL.distance], v[COL.desc], v[COL.link], v[COL.organizer] ? 'Szervező: ' + v[COL.organizer] : '',
    site ? 'Részletek és jelentkezés: ' + site : ''].filter(Boolean).join('\n\n');
  const opts = { location: v[COL.place] || '', description: desc };
  const title = calTitle_(v);
  let ev;
  if (v[COL.start]) {
    const p = s => Utilities.parseDate(s, TZ, 'yyyy-MM-dd HH:mm');
    const s = p(date + ' ' + v[COL.start]);
    let en = v[COL.end] ? p(lastDay + ' ' + v[COL.end]) : (lastDay !== date ? p(lastDay + ' 23:59') : new Date(s.getTime() + 3600000));
    if (en <= s) en = new Date(s.getTime() + 3600000);
    ev = cal.createEvent(title, s, en, opts);
  } else {
    // egész napos: déli időpont a szkript időzónájában, hogy biztosan a jó napra essen
    const stz = Session.getScriptTimeZone();
    const d = s => Utilities.parseDate(s + ' 12:00', stz, 'yyyy-MM-dd HH:mm');
    ev = lastDay !== date ? cal.createAllDayEvent(title, d(date), d(addDays_(lastDay, 1)), opts)
                          : cal.createAllDayEvent(title, d(date), opts);
  }
  ev.setTag('vescId', v[COL.id]);
  return ev;
}

/* ---------- segédfüggvények ---------- */
/** A nyilvános (/exec) naptár-link, akkor is, ha a teszt (/dev) linkről küldenek meghívót. */
function webAppUrl_() {
  const fixed = PropertiesService.getScriptProperties().getProperty('WEBAPP_URL');
  if (fixed && /^https:\/\/\S+$/.test(fixed)) return fixed;
  return String(ScriptApp.getService().getUrl() || '').replace(/\/dev$/, '/exec');
}
function withLock_(fn, auth) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { fn(); SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
  return getData(auth);
}
/** A sorozat sorai a megadott dátumtól kezdve. */
function seriesRows_(sh, seriesId, fromDate) {
  const last = sh.getLastRow();
  if (last < 2 || !seriesId) return [];
  return sh.getRange(2, 1, last - 1, HEAD.length).getDisplayValues()
    .map((v, i) => ({ row: i + 2, vals: v }))
    .filter(t => t.vals[COL.seriesId] === seriesId && t.vals[COL.date] >= fromDate);
}
function addDays_(ymd, n) {
  const p = ymd.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + n));
  return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2);
}
function dayDiff_(a, b) {
  const pa = a.split('-').map(Number), pb = b.split('-').map(Number);
  return Math.round((Date.UTC(pb[0], pb[1] - 1, pb[2]) - Date.UTC(pa[0], pa[1] - 1, pa[2])) / 86400000);
}
function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return -1;
}
function dayLong_(s) {
  const M = ['január', 'február', 'március', 'április', 'május', 'június', 'július', 'augusztus', 'szeptember', 'október', 'november', 'december'];
  const D = ['vasárnap', 'hétfő', 'kedd', 'szerda', 'csütörtök', 'péntek', 'szombat'];
  const p = s.split('-').map(Number); const d = new Date(p[0], p[1] - 1, p[2]);
  return `${p[0]}. ${M[p[1] - 1]} ${p[2]}., ${D[d.getDay()]}`;
}
function str_(v, max) { return String(v == null ? '' : v).replace(/;/g, ',').trim().slice(0, max); }
function safe_(v) { v = String(v == null ? '' : v); return /^[=+\-@]/.test(v) ? "'" + v : v; }
function clean_(ev) {
  const sport = SPORTS.indexOf(ev.sport) >= 0 ? ev.sport : 'other';
  const date = String(ev.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Érvénytelen dátum.');
  const t = v => (/^\d{2}:\d{2}$/.test(String(v || '')) ? String(v) : '');
  return {
    sport: sport, date: date, start: t(ev.start), end: t(ev.end),
    title: String(ev.title || '').trim().slice(0, 80),
    place: String(ev.place || '').trim().slice(0, 100),
    distance: String(ev.distance || '').trim().slice(0, 60),
    desc: String(ev.desc || '').trim().slice(0, 1000),
    endDate: (/^\d{4}-\d{2}-\d{2}$/.test(String(ev.endDate || '')) && ev.endDate > date && dayDiff_(date, ev.endDate) <= 60) ? String(ev.endDate) : '',
    raceType: sport === 'race' ? (RACE_TYPES[ev.raceType] ? ev.raceType : 'other') : '',
    link: /^https?:\/\/\S+$/i.test(String(ev.link || '').trim()) ? String(ev.link).trim().slice(0, 300) : '',
  };
}
