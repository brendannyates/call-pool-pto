/* Call Pool PTO: pure logic, no Office.js. Used by the task pane and by the node tests.
   The PTO tab is a day grid: column A holds "Call Pool N" section heads and provider names, each section's
   second row holds the dates (Excel serials), and each provider row holds codes from the tab's legend:
   P = PTO, 1/2 Day P = half-day PTO, C = call out, L = leave (LOA), H = holiday. O / OFF (regular day off) is ignored. */
(function (root) {
  "use strict";
  const DAY = 86400000;
  const DOW = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

  /* ---- dates: Excel serials (1900 system) <-> ISO yyyy-mm-dd, all in UTC so time zones never shift a day ---- */
  const serialToUTC = s => Math.round((s - 25569) * DAY);
  const utcToSerial = ms => Math.round(ms / DAY) + 25569;
  const isoToSerial = iso => { const [y, m, d] = String(iso).split("-").map(Number); return utcToSerial(Date.UTC(y, m - 1, d)); };
  const serialToISO = s => new Date(serialToUTC(s)).toISOString().slice(0, 10);
  const dow = s => new Date(serialToUTC(s)).getUTCDay();
  const isWeekday = s => { const d = dow(s); return d >= 1 && d <= 5; };
  const md = s => { const d = new Date(serialToUTC(s)); return (d.getUTCMonth() + 1) + "/" + d.getUTCDate(); };
  const dayLabel = s => DOW[dow(s)] + " " + md(s);
  const mdy = s => md(s) + "/" + serialToISO(s).slice(0, 4);
  /* local calendar date of a JS Date as a serial (the pane's "today") */
  const localSerial = d => utcToSerial(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  /* coverage weeks run Saturday to Friday; Settings stores the Monday */
  const satOfMonday = mon => mon - 2;
  const satOnOrBefore = s => s - ((dow(s) + 1) % 7);

  /* ---- codes ---- */
  const TYPES = {
    pto: { code: "P", label: "PTO", kind: "pto" },
    half: { code: "1/2 Day P", label: "½ day PTO", kind: "pto" },
    callout: { code: "C", label: "Call out", kind: "pto" },
    leave: { code: "L", label: "Leave", kind: "leave" }
  };
  const norm = v => String(v == null ? "" : v).trim().toUpperCase();
  function typeOf(v) {
    const n = norm(v);
    if (n === "") return "";
    if (n === "P") return "pto";
    if (n === "1/2 DAY P") return "half";
    if (n === "C") return "callout";
    if (n === "L") return "leave";
    if (n === "H") return "holiday";
    return "other";
  }
  const isOut = t => t === "pto" || t === "half" || t === "callout" || t === "leave";

  /* Excel TRIM: trims and collapses inner runs of spaces */
  const xtrim = v => String(v == null ? "" : v).replace(/ +/g, " ").trim();
  const lastName = s => { const t = xtrim(s); const i = t.lastIndexOf(" "); return (i < 0 ? t : t.slice(i + 1)).toLowerCase(); };
  const sameLast = (a, b) => !!xtrim(a) && !!xtrim(b) && lastName(a) === lastName(b);

  /* ---- PTO tab layout ---- */
  const isPoolHead = v => /^call pool \d+$/i.test(xtrim(v));
  /* colA: column A values for rows 1..n. Returns sections with 1-based rows. */
  function parseSections(colA) {
    const a = colA.map(xtrim), heads = [];
    a.forEach((t, i) => { if (isPoolHead(t)) heads.push(i + 1); });
    let lastUsed = 0; a.forEach((t, i) => { if (t) lastUsed = i + 1; });
    return heads.map((head, k) => {
      const end = k + 1 < heads.length ? heads[k + 1] - 1 : lastUsed;
      let first = head + 2;
      for (let r = head + 1; r <= end; r++) if (/^providers:?$/i.test(a[r - 1])) { first = r + 1; break; }
      const rows = [];
      for (let r = first; r <= end; r++) if (a[r - 1]) rows.push({ row: r, name: a[r - 1], placeholder: /^employee$/i.test(a[r - 1]) });
      const t = a[head - 1];
      return { pool: "Call Pool " + t.match(/\d+/)[0], num: +t.match(/\d+/)[0], head, dateRow: head + 1, first, last: Math.max(first, end), rows };
    });
  }
  const providers = sections => sections.flatMap(s => s.rows.filter(r => !r.placeholder).map(r => ({ ...r, pool: s.pool, num: s.num })));

  /* ---- week status (same rules as the sheet formulas) ---- */
  function runs(flags, days) {
    const out = [];
    for (let i = 0; i < flags.length; i++) {
      if (!flags[i] || (i > 0 && flags[i - 1])) continue;
      let j = i; while (j + 1 < flags.length && flags[j + 1]) j++;
      out.push(i === j ? md(days[i]) : md(days[i]) + " to " + md(days[j]));
    }
    return out.join(", ");
  }
  /* codes7: the 7 cells Saturday..Friday. Returns the PTO-column text, the Leave-column text and per-type day lists. */
  function weekStatus(codes7, sat) {
    const days = Array.from({ length: 7 }, (_, i) => sat + i), t = codes7.map(typeOf);
    const part = key => { const f = t.map(x => x === key); return f.some(Boolean) ? TYPES[key].label + " " + runs(f, days) : ""; };
    const lf = t.map(x => x === "leave");
    const items = [];
    ["pto", "half", "callout", "leave"].forEach(key => { const f = t.map(x => x === key); if (f.some(Boolean)) items.push({ type: key, label: TYPES[key].label, dates: runs(f, days) }); });
    return { pto: [part("pto"), part("half"), part("callout")].filter(Boolean).join("; "), leave: lf.some(Boolean) ? runs(lf, days) : "", items };
  }
  /* Dashboard "Covering" text for one pool: entries [{name, pto, leave}] in sheet order */
  function coveringText(entries, onCall) {
    return entries.map(e => {
      const info = e.pto + (e.pto && e.leave ? "; " : "") + (e.leave ? "Leave " + e.leave : "");
      return info ? (sameLast(e.name, onCall) ? "⚠ ON CALL: " : "") + xtrim(e.name) + " " + info : "";
    }).filter(Boolean).join(" & ");
  }

  /* ---- writing time off ---- */
  /* cells: [{serial, value}] for each day from..to. typeKey: pto | half | callout | leave | clear.
     Never overwrites holidays or O/OFF days; weekends are skipped unless weekdaysOnly is false. */
  function planWrite(cells, typeKey, opts) {
    const o = Object.assign({ weekdaysOnly: true }, opts), changes = [], skipped = [];
    const code = typeKey === "clear" ? "" : TYPES[typeKey].code;
    const values = cells.map(c => {
      const cur = c.value == null ? "" : c.value, t = typeOf(cur);
      if (typeKey === "clear") { if (isOut(t)) { changes.push({ serial: c.serial, from: cur, to: "" }); return ""; } return cur; }
      if (t === "holiday") { skipped.push({ serial: c.serial, why: "holiday" }); return cur; }
      if (t === "other") { skipped.push({ serial: c.serial, why: "day off (" + xtrim(cur) + ")" }); return cur; }
      if (o.weekdaysOnly && !isWeekday(c.serial)) { if (!t) skipped.push({ serial: c.serial, why: "weekend" }); return cur; }
      if (norm(cur) === norm(code)) return cur;
      changes.push({ serial: c.serial, from: cur, to: code });
      return code;
    });
    return { values, changes, skipped };
  }

  /* runs of time off in a row of cells [{serial, value}]: same type, joined across weekends, holidays and days off */
  function timeOffRuns(cells) {
    const out = []; let cur = null, gapWork = false;
    for (const c of cells) {
      const t = typeOf(c.value);
      if (isOut(t)) {
        if (cur && cur.type === t && !gapWork) { cur.end = c.serial; cur.days++; }
        else { cur = { type: t, label: TYPES[t].label, start: c.serial, end: c.serial, days: 1 }; out.push(cur); }
        gapWork = false;
      } else if (cur && t === "" && isWeekday(c.serial)) gapWork = true;
    }
    return out.map(r => ({ ...r, text: r.start === r.end ? dayLabel(r.start) : dayLabel(r.start) + " to " + dayLabel(r.end) }));
  }

  /* weeks this provider is on call in [from, to]. weeks [{label, mon}], assign {label: [name pool1..8]} */
  function onCallWeeks(name, num, from, to, weeks, assign) {
    return weeks.filter(w => { const sat = satOfMonday(w.mon); return sat <= to && sat + 6 >= from && sameLast((assign[w.label] || [])[num - 1], name); });
  }
  function weekForSerial(weeks, s) { return weeks.find(w => s >= satOfMonday(w.mon) && s <= satOfMonday(w.mon) + 6) || null; }

  /* ---- sheet formulas (tested in Excel; the pane installs them, Setup shows them for copying) ---- */
  const F_PTO = `=LET(
  wk, $B$2,
  mon, XLOOKUP(wk, Cov_Wk_Range[Coverage Week Range], Wk_Start_Date[Week Start Date (Mon)], ""),
  c, IFERROR(XMATCH(mon - 2, $D$4:$XFD$4), 0),
  IF(OR(TRIM($A{R}) = "", c = 0), "",
  LET(
    days, mon - 2 + SEQUENCE(1, 7, 0),
    code, UPPER(TRIM(INDEX($D{R}:$XFD{R}, 1, c):INDEX($D{R}:$XFD{R}, 1, c + 6) & "")),
    runs, LAMBDA(f, LET(
      s, FILTER(days, f * NOT(HSTACK(FALSE, TAKE(f, , 6)))),
      e, FILTER(days, f * NOT(HSTACK(DROP(f, , 1), FALSE))),
      TEXTJOIN(", ", TRUE, IF(s = e, TEXT(s, "m/d"), TEXT(s, "m/d") & " to " & TEXT(e, "m/d"))))),
    pick, LAMBDA(label, f, IF(OR(f), label & " " & runs(f), "")),
    TEXTJOIN("; ", TRUE,
      pick("PTO", code = "P"),
      pick("½ day PTO", code = "1/2 DAY P"),
      pick("Call out", code = "C"))
  )))`;
  const F_LEAVE = `=LET(
  wk, $B$2,
  mon, XLOOKUP(wk, Cov_Wk_Range[Coverage Week Range], Wk_Start_Date[Week Start Date (Mon)], ""),
  c, IFERROR(XMATCH(mon - 2, $D$4:$XFD$4), 0),
  IF(OR(TRIM($A{R}) = "", c = 0), "",
  LET(
    days, mon - 2 + SEQUENCE(1, 7, 0),
    f, UPPER(TRIM(INDEX($D{R}:$XFD{R}, 1, c):INDEX($D{R}:$XFD{R}, 1, c + 6) & "")) = "L",
    IF(NOT(OR(f)), "",
    LET(
      s, FILTER(days, f * NOT(HSTACK(FALSE, TAKE(f, , 6)))),
      e, FILTER(days, f * NOT(HSTACK(DROP(f, , 1), FALSE))),
      TEXTJOIN(", ", TRUE, IF(s = e, TEXT(s, "m/d"), TEXT(s, "m/d") & " to " & TEXT(e, "m/d")))))
  )))`;
  const F_COVERING = `=LET(
  pool, "Call Pool " & $B{R},
  nm, TRIM(PTO!$A$1:$A$300 & ""),
  sec, SCAN("", nm, LAMBDA(a, x, IF(LEFT(x, 10) = "Call Pool ", x, a))),
  pto, PTO!$B$1:$B$300 & "",
  lv, PTO!$C$1:$C$300 & "",
  last, LAMBDA(t, IFERROR(TEXTAFTER(TRIM(t), " ", -1), TRIM(t))),
  info, pto & IF((pto <> "") * (lv <> ""), "; ", "") & IF(lv <> "", "Leave " & lv, ""),
  txt, IF(info = "", "", IF(last(nm) = last($C{R}), "⚠ ON CALL: ", "") & nm & " " & info),
  keep, (sec = pool) * (nm <> pool) * (nm <> "Providers:") * (nm <> "") * (txt <> ""),
  TEXTJOIN(" & ", TRUE, FILTER(txt, keep, ""))
)`;
  /* dateRow: the PTO tab row that holds the dates (row 4 in the Call Pool workbook) */
  const withDateRow = (f, dateRow) => f.split("$D$4:$XFD$4").join("$D$" + dateRow + ":$XFD$" + dateRow);
  const FORMULAS = {
    pto: (row, dateRow = 4) => withDateRow(F_PTO, dateRow).split("{R}").join(String(row)),
    leave: (row, dateRow = 4) => withDateRow(F_LEAVE, dateRow).split("{R}").join(String(row)),
    covering: row => F_COVERING.split("{R}").join(String(row)),
    week: "=Dashboard!$B$3"
  };
  /* a formula cell "looks installed" when it carries our signature */
  const isOurs = (f, kind) => typeof f === "string" && (kind === "covering" ? f.includes("ON CALL") && f.includes("PTO!$B$1") : kind === "leave" ? f.includes(`= "L"`) && f.includes("Cov_Wk_Range") : f.includes(`pick("PTO"`));

  /* rough row height for wrapped text in the merged Covering cell */
  const rowHeightFor = (text, charsPerLine, min) => Math.max(min, Math.ceil(String(text || "").length / Math.max(10, charsPerLine)) * 16 + 8);

  /* ---- who's using the pane (the same model as the IVK Tracker) ----
     The people list lives in the workbook's add-in settings (cpUsers: [{ name, added }]), never on a sheet. This computer remembers,
     per workbook, who picked today and who has used it here before: { day: "2026-11-07", name: "Sam", used: ["Sam", "Ana"] }. */
  const txt = v => String(v == null ? "" : v).trim();
  const userKey = n => txt(n).toLowerCase();
  /* the name typed to join the list: usually a first name, a last initial when two people share one */
  const cleanName = s => txt(s).replace(/\s+/g, " ").slice(0, 30).trim();
  function userList(v) {
    const seen = new Set(), out = [];
    for (const u of Array.isArray(v) ? v : []) {
      const name = cleanName(u && typeof u === "object" ? u.name : u), k = userKey(name);
      if (!name || seen.has(k)) continue; seen.add(k); out.push({ name, added: txt(u && u.added) });
    }
    return out;
  }
  const addUser = (list, name, addedIso) => { const L = userList(list), n = cleanName(name); return !n || L.some(u => userKey(u.name) === userKey(n)) ? L : L.concat({ name: n, added: addedIso || "" }); };
  const removeUser = (list, name) => userList(list).filter(u => userKey(u.name) !== userKey(name));
  /* the picker's order: people who have used this computer on this workbook first (newest first), then everyone else A-Z */
  function pickOrder(users, here) {
    const H = (here || []).map(userKey), rank = u => { const i = H.indexOf(userKey(u.name)); return i < 0 ? Infinity : i; };
    return userList(users).map(u => Object.assign({}, u, { here: rank(u) !== Infinity }))
      .sort((a, b) => rank(a) - rank(b) || a.name.localeCompare(b.name, undefined, { sensitivity: "base" }));
  }
  /* a short key from the workbook's URL (or the id saved in it), so one workbook's pick never carries to another */
  function wbKeyOf(s) {
    let h = 0x811c9dc5; const t = String(s || "").trim().toLowerCase().replace(/[?#].*$/, "");
    for (let i = 0; i < t.length; i++) { h ^= t.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; }
    return t ? "w" + h.toString(36) : "";
  }
  function whoRec(json) { try { const o = JSON.parse(json || "{}"); return { day: txt(o.day), name: txt(o.name), used: Array.isArray(o.used) ? o.used.map(txt).filter(Boolean) : [] }; } catch (e) { return { day: "", name: "", used: [] }; } }
  const whoToday = (rec, dayIso) => rec && rec.day === dayIso ? rec.name : "";
  const whoPick = (rec, name, dayIso) => ({ day: dayIso, name, used: [name].concat(((rec && rec.used) || []).filter(n => userKey(n) !== userKey(name))).slice(0, 20) });

  const CP = { DAY, DOW, mdy, serialToISO, isoToSerial, dow, isWeekday, md, dayLabel, localSerial, satOfMonday, satOnOrBefore, TYPES, typeOf, isOut, norm, xtrim, lastName, sameLast,
    isPoolHead, parseSections, providers, runs, weekStatus, coveringText, planWrite, timeOffRuns, onCallWeeks, weekForSerial, FORMULAS, isOurs, rowHeightFor,
    userKey, cleanName, userList, addUser, removeUser, pickOrder, wbKeyOf, whoRec, whoToday, whoPick };
  if (typeof module === "object" && module.exports) module.exports = CP; else root.CP = CP;
})(typeof window !== "undefined" ? window : globalThis);
