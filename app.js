/* Call Pool PTO: task pane UI. Three tabs: This week (who's out per call pool, on-call conflicts, send to Dashboard),
   Log time off (write codes into the PTO grid) and Setup (checks, install formulas, copy formulas). */
(function () {
  "use strict";
  const $ = (s, el = document) => el.querySelector(s);
  const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const S = { view: "week", model: null, weekIdx: -1, weekRows: null, form: { row: "", type: "pto", from: "", to: "", weekdays: true }, preview: null, runs: null, pw: null, busy: false, error: "" };
  window.APP = S;

  /* ---- helpers ---- */
  let toastT;
  function toast(text) { const t = $("#toast"); t.textContent = text; t.classList.add("on"); clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove("on"), 3800); }
  const time12 = d => d.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
  const todaySerial = () => CP.localSerial(new Date());
  const weekOf = i => S.model.weeks[i];
  const satOf = w => CP.satOfMonday(w.mon);
  const weekText = w => `${CP.dayLabel(satOf(w))} to ${CP.dayLabel(satOf(w) + 6)}`;
  const provList = () => CP.providers(S.model.sections);
  const provByRow = row => provList().find(p => p.row === +row);
  const chip = (type, label) => `<span class="chip ${type}">${esc(label)}</span>`;

  /* password prompt (blurs the pane). Resolves to the password or null. Never stored beyond this pane session. */
  function askPassword(sheets, errText) {
    return new Promise(resolve => {
      const box = document.createElement("div"); box.className = "scrim";
      box.innerHTML = `<div class="modal" role="dialog" aria-modal="true" aria-labelledby="pwh"><h3 id="pwh">Unlock to save</h3>
        <p>${esc(sheets.join(" and "))} ${sheets.length > 1 ? "are" : "is"} protected. Type the sheet password. The add-in unlocks, saves, then locks again with the same settings. The password isn't saved or sent anywhere.</p>
        <input type="password" id="pwi" autocomplete="off" aria-label="Sheet password">
        <label class="check"><input type="checkbox" id="pwr"> Remember until I close this pane</label>
        ${errText ? `<div class="err">${esc(errText)}</div>` : ""}
        <div class="act"><button class="btn" data-x="cancel">Cancel</button><button class="btn primary" data-x="ok">Unlock and save</button></div></div>`;
      document.body.appendChild(box);
      const input = $("#pwi", box); setTimeout(() => input.focus(), 30);
      const done = v => { box.remove(); resolve(v); };
      box.addEventListener("click", e => { const x = e.target.closest("[data-x]"); if (!x) return; if (x.dataset.x === "cancel") done(null); else { if ($("#pwr", box).checked) S.remember = true; done(input.value); } });
      box.addEventListener("keydown", e => { if (e.key === "Escape") done(null); if (e.key === "Enter") { if ($("#pwr", box).checked) S.remember = true; done(input.value); } });
    });
  }
  /* runs a write; asks for the password when a sheet is protected, retries once per wrong password */
  async function guarded(fn) {
    let pw = S.pw, err = "", sheets = ["The sheet"];
    for (let i = 0; i < 4; i++) {
      try { const out = await fn(pw); if (S.remember && pw) S.pw = pw; return { ok: true, out }; }
      catch (e) {
        if (e instanceof XL.NeedPassword || e instanceof XL.BadPassword) {
          if (e.sheets) sheets = e.sheets;
          if (e instanceof XL.BadPassword) { S.pw = null; err = "That password didn't unlock the sheet. Try again."; }
          pw = await askPassword(sheets, err);
          if (pw == null) return { ok: false, cancelled: true };
          continue;
        }
        throw e;
      }
    }
    return { ok: false };
  }

  /* ---- data ---- */
  async function load(keepWeek) {
    S.error = "";
    try {
      S.model = await XL.readModel();
      const m = S.model;
      if (!keepWeek || S.weekIdx < 0) {
        let i = m.dash ? m.weeks.findIndex(w => w.label === m.dash.week) : -1;
        if (i < 0) { const w = CP.weekForSerial(m.weeks, todaySerial()); i = w ? m.weeks.indexOf(w) : 0; }
        S.weekIdx = i;
      }
      await loadWeek();
    } catch (e) { S.error = e.message || String(e); }
    render();
  }
  async function loadWeek() {
    const w = weekOf(S.weekIdx); S.weekRows = w ? await XL.readWeek(S.model, satOf(w)) : null;
  }
  function weekData() {
    const m = S.model, w = weekOf(S.weekIdx); if (!w || !S.weekRows) return null;
    const sat = satOf(w), onCall = m.assign[w.label] || [];
    return m.sections.map(s => {
      const out = [], entries = [];
      s.rows.forEach(p => {
        const st = CP.weekStatus(S.weekRows[p.row - 1] || [], sat);
        entries.push({ name: p.name, pto: st.pto, leave: st.leave });
        if (st.items.length) out.push({ ...p, items: st.items });
      });
      const oc = onCall[s.num - 1] || "";
      const conflict = out.find(p => CP.sameLast(p.name, oc));
      return { ...s, onCall: oc, out, conflict, covering: CP.coveringText(entries, oc) };
    });
  }

  /* ---- views ---- */
  function render() {
    const app = $("#app");
    if (S.error && !S.model) { app.innerHTML = `${frame()}<main><div class="msg bad">${esc(S.error)}</div><div class="act" style="margin-top:12px"><button class="btn" data-a="reload">Try again</button></div></main>`; return; }
    if (!S.model) { app.innerHTML = `<div class="landing"><h2>Call Pool PTO</h2><p class="muted">Loading…</p></div>`; return; }
    const body = S.view === "log" ? viewLog() : S.view === "setup" ? viewSetup() : viewWeek();
    app.innerHTML = frame() + `<main class="stack">${S.error ? `<div class="msg bad">${esc(S.error)}</div>` : ""}${body}</main>`;
    if (S.view === "log") bindLog();
  }
  function frame() {
    const tab = (v, t) => `<button data-view="${v}" ${S.view === v ? 'aria-current="page"' : ""}>${t}</button>`;
    return `<header class="top"><div class="brand">Call Pool PTO <small>${S.model && S.model.dash && S.model.dash.week ? "Dashboard week " + esc(S.model.dash.week) : ""}</small></div><nav class="tabs">${tab("week", "This week")}${tab("log", "Log time off")}${tab("setup", "Setup")}</nav></header>`;
  }

  function viewWeek() {
    const m = S.model, w = weekOf(S.weekIdx);
    if (!m.weeks.length) return `<div class="msg warn">Couldn't read the coverage weeks (Settings tables <b>Wk_Start_Date</b> and <b>Cov_Wk_Range</b>).</div>`;
    const opts = m.weeks.map((x, i) => `<option value="${i}" ${i === S.weekIdx ? "selected" : ""}>${esc(x.label)}</option>`).join("");
    const isDash = m.dash && w && w.label === m.dash.week;
    const data = weekData();
    let html = `<div><div class="weekbar"><button class="arrow" data-a="prev" aria-label="Previous week" ${S.weekIdx <= 0 ? "disabled" : ""}>‹</button>
      <select id="wk" aria-label="Coverage week">${opts}</select>
      <button class="arrow" data-a="next" aria-label="Next week" ${S.weekIdx >= m.weeks.length - 1 ? "disabled" : ""}>›</button></div>
      <div class="weeksub">${w ? esc(weekText(w)) : ""}${isDash ? '<span class="tag">Dashboard week</span>' : ""}</div></div>`;
    if (!data) return html + `<div class="msg warn">The PTO sheet has no dates for this week yet.</div>`;
    if (!m.installed.dash || !m.installed.pto) html += `<div class="msg warn"><b>Dashboard › Covering isn't linked yet.</b> Send to Dashboard installs live formulas, so Covering fills itself from the PTO tab every week.<div class="act"><button class="btn primary" data-a="install">Send to Dashboard</button></div></div>`;
    const total = data.reduce((n, s) => n + s.out.length, 0), conflicts = data.filter(s => s.conflict);
    if (conflicts.length) html += `<div class="msg bad"><b>${conflicts.length} on-call conflict${conflicts.length > 1 ? "s" : ""}.</b> ${conflicts.map(s => esc(s.pool)).join(", ")}: the provider on call is also out. Add an override on the Overrides tab.</div>`;
    html += `<p class="small muted" style="margin:10px 2px 0">${total ? `${total} provider${total > 1 ? "s" : ""} out this week.` : "No one is out this week."} Holidays and regular days off aren't listed.</p>`;
    html += data.map(s => `<section class="card pool"><div class="ph"><h2>${esc(s.pool)}</h2><span class="count" title="Out">${s.out.length}</span></div>
      <div class="oncall">On call: <b>${esc(s.onCall || "—")}</b></div>
      ${s.out.length ? `<ul>${s.out.map(p => `<li><span class="nm">${esc(p.name)}</span><span class="chips">${[...new Set(p.items.map(i => i.type))].map(t => chip(t, CP.TYPES[t].label)).join("")}</span>
        <span class="dt">${p.items.map(i => esc((i.type === "leave" ? "Leave " : i.label + " ") + i.dates)).join(" · ")}</span></li>`).join("")}</ul>` : `<p class="none">No one out</p>`}
      ${s.conflict ? `<div class="conflict"><b>On call but out:</b> ${esc(s.conflict.name)} (${esc(s.conflict.items.map(i => i.label + " " + i.dates).join("; "))}). Pick a cover on the Overrides tab.</div>` : ""}
    </section>`).join("");
    if (m.installed.dash && m.installed.pto) html += `<div class="msg ok">Dashboard › Covering is live: it reads the PTO tab for the Dashboard week (${esc(m.dash ? m.dash.week : "")}).<div class="act"><button class="btn" data-a="fit">Resize Dashboard rows</button><button class="btn quiet" data-a="gotoDash">Open Dashboard</button></div></div>`;
    return html;
  }

  /* ---- log time off ---- */
  function viewLog() {
    const f = S.form, m = S.model;
    const groups = m.sections.map(s => `<optgroup label="${esc(s.pool)}">${s.rows.filter(r => !r.placeholder).map(r => `<option value="${r.row}" ${+f.row === r.row ? "selected" : ""}>${esc(r.name)}</option>`).join("")}</optgroup>`).join("");
    const types = [["pto", "PTO"], ["half", "½ day PTO"], ["leave", "Leave"], ["callout", "Call out"], ["clear", "Clear"]];
    const p = S.preview;
    let prev = "";
    if (p && p.error) prev = `<div class="msg bad" style="margin-top:12px">${esc(p.error)}</div>`;
    else if (p) {
      const code = f.type === "clear" ? "" : CP.TYPES[f.type].code;
      prev = `<div class="preview" aria-live="polite">${p.changes.length ? `<b>${f.type === "clear" ? "Clears" : "Marks"} ${p.changes.length} day${p.changes.length > 1 ? "s" : ""}</b>${code ? ` as <b>${esc(code)}</b>` : ""}:
        <div class="days">${p.changes.map(c => `<span class="d ${CP.typeOf(c.from) ? "rep" : ""}" title="${c.from ? "was " + esc(c.from) : ""}">${esc(CP.dayLabel(c.serial))}${CP.typeOf(c.from) && f.type !== "clear" ? " (was " + esc(c.from) + ")" : ""}</span>`).join("")}</div>` : `<b>Nothing to change</b> in these dates.`}
        ${p.skipped.filter(x => x.why !== "weekend").length ? `<p class="muted">Left alone: ${p.skipped.filter(x => x.why !== "weekend").map(x => esc(CP.dayLabel(x.serial) + " " + x.why)).join(", ")}.</p>` : ""}
        ${p.onCall.length ? `<p><b style="color:var(--danger)">On call:</b> this provider is on call the week of ${p.onCall.map(w => esc(w.label)).join(", ")}. Pick a cover on the Overrides tab.</p>` : ""}</div>`;
    }
    const runs = S.runs;
    return `<section class="card"><form id="lf" autocomplete="off">
      <label class="f" for="fp">Provider</label><select id="fp" required><option value="">Choose a provider…</option>${groups}</select>
      <label class="f">Type</label><div class="seg" role="group" aria-label="Type">${types.map(([k, t]) => `<button type="button" class="${k === "clear" ? "clear" : ""}" data-type="${k}" aria-pressed="${f.type === k}">${t}</button>`).join("")}</div>
      <div class="two"><div><label class="f" for="ff">From</label><input type="date" id="ff" value="${esc(f.from)}" required></div>
      <div><label class="f" for="ft">To</label><input type="date" id="ft" value="${esc(f.to)}" required></div></div>
      <label class="check"><input type="checkbox" id="fw" ${f.weekdays ? "checked" : ""}> Weekdays only</label>
      ${prev}
      <button class="btn primary wide" style="margin-top:14px" id="fs" ${p && !p.error && p.changes.length && !S.busy ? "" : "disabled"}>${f.type === "clear" ? "Clear these days" : "Save time off"}</button>
    </form></section>
    ${f.row ? `<section class="card"><h3>${esc((provByRow(f.row) || {}).name || "")}: upcoming</h3>${runs == null ? `<p class="muted small">Loading…</p>` : runs.length ? `<ul class="runs">${runs.map(r => `<li><span>${chip(r.type, r.label)} ${esc(r.text)}</span><button class="btn quiet" data-a="clearRun" data-s="${r.start}" data-e="${r.end}">Clear</button></li>`).join("")}</ul>` : `<p class="muted small">Nothing logged in the next 6 months.</p>`}</section>` : ""}
    <p class="small muted" style="margin:0 2px">Writes the PTO tab's own codes (P, 1/2 Day P, L, C). Holidays (H) and days off (O, OFF) are never overwritten.</p>`;
  }
  function bindLog() {
    const f = S.form;
    $("#fp").addEventListener("change", e => { f.row = e.target.value; S.runs = null; refreshPreview(); loadRuns(); });
    $("#ff").addEventListener("change", e => { f.from = e.target.value; if (!f.to || f.to < f.from) f.to = f.from; refreshPreview(); });
    $("#ft").addEventListener("change", e => { f.to = e.target.value; refreshPreview(); });
    $("#fw").addEventListener("change", e => { f.weekdays = e.target.checked; refreshPreview(); });
    $("#lf").addEventListener("submit", e => { e.preventDefault(); save(); });
  }
  let previewSeq = 0;
  async function refreshPreview() {
    const f = S.form, seq = ++previewSeq;
    if (!f.row || !f.from || !f.to) { S.preview = null; return render(); }
    const from = CP.isoToSerial(f.from), to = CP.isoToSerial(f.to);
    if (to < from) { S.preview = { error: "“To” is before “From”." }; return render(); }
    if (to - from > 366) { S.preview = { error: "Pick a range of a year or less." }; return render(); }
    try {
      const cells = await XL.readRow(S.model, +f.row, from, to);
      if (seq !== previewSeq) return;
      const plan = CP.planWrite(cells, f.type, { weekdaysOnly: f.weekdays }), p = provByRow(f.row);
      plan.onCall = f.type === "clear" ? [] : CP.onCallWeeks(p.name, p.num, from, to, S.model.weeks, S.model.assign);
      plan.from = from; S.preview = plan;
    } catch (e) { if (seq === previewSeq) S.preview = { error: e.message || String(e) }; }
    render();
  }
  async function loadRuns() {
    const row = +S.form.row; if (!row) return;
    const from = Math.max(todaySerial(), S.model.minDate), to = Math.min(from + 182, S.model.maxDate);
    try { const cells = await XL.readRow(S.model, row, from, to); if (+S.form.row === row) { S.runs = CP.timeOffRuns(cells); render(); } }
    catch (e) { S.runs = []; render(); }
  }
  async function writePlan(row, from, values, doneText) {
    S.busy = true; render();
    try {
      const r = await guarded(pw => XL.writeRow(S.model, row, from, values, pw));
      if (r.ok) { toast(doneText + " · " + time12(new Date())); await loadWeek(); }
      return r.ok;
    } catch (e) { S.error = e.message || String(e); return false; }
    finally { S.busy = false; }
  }
  async function save() {
    const f = S.form, p = S.preview, prov = provByRow(f.row); if (!p || p.error || !p.changes.length) return;
    const n = p.changes.length, what = f.type === "clear" ? "Cleared" : "Saved " + CP.TYPES[f.type].label;
    const ok = await writePlan(+f.row, p.from, p.values, `${what}: ${n} day${n > 1 ? "s" : ""} for ${prov.name}`);
    if (ok) { S.preview = null; await refreshPreview(); await loadRuns(); } else render();
  }
  async function clearRun(start, end) {
    const prov = provByRow(S.form.row);
    const cells = await XL.readRow(S.model, prov.row, start, end), plan = CP.planWrite(cells, "clear", {});
    if (!plan.changes.length) return;
    if (await writePlan(prov.row, start, plan.values, `Cleared ${plan.changes.length} day${plan.changes.length > 1 ? "s" : ""} for ${prov.name}`)) { await loadRuns(); await refreshPreview(); } else render();
  }

  /* ---- setup ---- */
  function viewSetup() {
    const m = S.model, row = (m.sections[0] || {}).first || 6;
    const ok = (b, t, no) => `<li><span class="i ${b ? "ok" : "no"}">${b ? "✓" : "!"}</span><span>${b ? t : no || t}</span></li>`;
    const provs = provList().length;
    const typed = m.dash ? m.dash.rows.filter(r => r.covering && !r.formula.startsWith("=")) : [];
    return `<section class="card"><h3>Workbook</h3><ul class="checks">
      ${ok(true, `PTO sheet: ${m.sections.length} call pools, ${provs} providers`)}
      ${ok(isFinite(m.minDate), `Dates ${esc(CP.mdy(m.minDate))} to ${esc(CP.mdy(m.maxDate))} (row ${m.dateRow})`, "No dates found on the PTO sheet")}
      ${ok(m.tables.weeks, `${m.weeks.length} coverage weeks (Settings)`, "Settings tables Wk_Start_Date / Cov_Wk_Range not found")}
      ${ok(m.tables.assign, "Weekly assignments table (Dashboard)", "Tbl_Yr_Assignments not found")}
      ${ok(!!m.dash, "Dashboard sheet", "No Dashboard sheet: only the PTO columns will be installed")}
      ${ok(m.installed.pto, "PTO tab: PTO and Leave columns (B:C) installed", "PTO tab: PTO and Leave columns (B:C) not installed yet")}
      ${ok(m.installed.dash, "Dashboard › Covering (D7:D14) is live", "Dashboard › Covering (D7:D14) not linked yet")}
      ${m.ptoProtected || (m.dash && m.dash.protected) ? ok(false, "", `Protected: ${[m.ptoProtected && "PTO", m.dash && m.dash.protected && "Dashboard"].filter(Boolean).join(" and ")}. You'll be asked for the sheet password when saving.`) : ""}
    </ul>
    ${typed.length ? `<div class="msg warn" style="margin-top:10px">Installing replaces the typed Covering notes in ${typed.map(r => "D" + r.row).join(", ")}. A copy is kept in this workbook's add-in settings.</div>` : ""}
    <button class="btn primary wide" style="margin-top:12px" data-a="install">${m.installed.pto && m.installed.dash ? "Repair formulas" : "Install formulas"}</button>
    <p class="small muted">Shows a PTO and a Leave column next to each name on the PTO tab for the Dashboard week, and fills Dashboard › Covering from them. Safe to run again.</p></section>
    <section class="card"><h3>Formulas</h3><p class="small muted" style="margin:0">The same formulas the add-in installs, for copying by hand. Row ${row} shown; change the row numbers for other rows.</p>
      ${fx("PTO!B2 (week shown)", CP.FORMULAS.week)}
      ${fx(`PTO!B${row} (PTO column)`, CP.FORMULAS.pto(row, m.dateRow))}
      ${fx(`PTO!C${row} (Leave column)`, CP.FORMULAS.leave(row, m.dateRow))}
      ${fx("Dashboard!D7 (Covering, copy down to D14)", CP.FORMULAS.covering(7))}
      ${fx("Dashboard!C7 (On call, already in the workbook)", '=INDEX(Tbl_Yr_Assignments[], MATCH($B$3, Tbl_Yr_Assignments[Week], 0), MATCH("Call Pool " & $B7, Tbl_Yr_Assignments[#Headers], 0))')}
    </section>
    <section class="card"><h3>How it reads the PTO tab</h3><p class="small" style="margin:0">Weeks run Saturday to Friday, like the rotation. P = PTO, 1/2 Day P = half-day PTO, C = call out, L = leave. Holidays (H) and regular days off (O, OFF) aren't listed. A provider on call who is also out is flagged ⚠.</p>
    <div class="act" style="margin-top:10px;display:flex;gap:8px"><button class="btn" data-a="reload">Reload workbook</button><button class="btn quiet" data-a="gotoPto">Open PTO tab</button></div></section>`;
  }
  const fx = (title, f) => `<details class="fx"><summary>${esc(title)}</summary><pre>${esc(f)}</pre><button class="btn quiet" data-a="copy" data-f="${esc(f)}">Copy</button></details>`;

  async function install() {
    const m = S.model;
    if (m.dash) {
      const typed = m.dash.rows.filter(r => r.covering && !r.formula.startsWith("="));
      if (typed.length) {
        try { const st = Office.context.document.settings; st.set("cpCoveringBackup", { at: new Date().toISOString(), week: m.dash.week, cells: typed.map(r => ({ cell: "D" + r.row, text: r.covering })) }); await new Promise(res => st.saveAsync(res)); } catch (e) { /* settings unavailable: carry on */ }
      }
    }
    S.busy = true;
    try {
      const r = await guarded(pw => XL.install(m, pw));
      if (r.ok) { toast("Formulas installed. Covering is live · " + time12(new Date())); await load(true); return; }
    } catch (e) { S.error = "Couldn't install the formulas: " + (e.message || e); }
    finally { S.busy = false; }
    render();
  }
  async function copyText(t) {
    try { await navigator.clipboard.writeText(t); }
    catch (e) { const ta = document.createElement("textarea"); ta.value = t; document.body.appendChild(ta); ta.select(); try { document.execCommand("copy"); } catch (x) { /* ignore */ } ta.remove(); }
    toast("Copied");
  }

  /* ---- events ---- */
  document.addEventListener("click", async e => {
    const v = e.target.closest("[data-view]"); if (v) { S.view = v.dataset.view; if (S.view === "log" && !S.form.from) { const t = CP.serialToISO(todaySerial()); S.form.from = t; S.form.to = t; } render(); return; }
    const t = e.target.closest("[data-type]"); if (t) { S.form.type = t.dataset.type; refreshPreview(); return; }
    const a = e.target.closest("[data-a]"); if (!a || a.disabled) return;
    const act = a.dataset.a;
    if (act === "prev" || act === "next") { S.weekIdx += act === "prev" ? -1 : 1; await loadWeek(); render(); }
    else if (act === "install") install();
    else if (act === "reload") { S.model = null; render(); load(true); }
    else if (act === "fit") { const r = await guarded(pw => XL.refreshDashboard(pw)); if (r.ok) toast("Dashboard rows resized"); }
    else if (act === "gotoDash") XL.goTo(XL.DASH, "D7");
    else if (act === "gotoPto") XL.goTo(XL.PTO, "B1");
    else if (act === "copy") copyText(a.dataset.f);
    else if (act === "clearRun") clearRun(+a.dataset.s, +a.dataset.e);
  });
  document.addEventListener("change", async e => { if (e.target.id === "wk") { S.weekIdx = +e.target.value; await loadWeek(); render(); } });

  window.reset = () => { S.view = "week"; S.form = { row: "", type: "pto", from: "", to: "", weekdays: true }; S.preview = null; S.runs = null; S.pw = null; S.remember = false; };
  window.show = v => { S.view = v; render(); };
  window.reload = () => load(false);
  Office.onReady(() => load(false));
})();
