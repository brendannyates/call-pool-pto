/* Call Pool PTO: every workbook read and write (Office.js). Reads the PTO tab, the Dashboard and the Settings week tables;
   writes time-off codes into the PTO grid and installs the status/Covering formulas. Protected sheets are unlocked with the
   password the user types into the pane, then locked again with the same options. The password is never stored. */
const XL = (() => {
  "use strict";
  const PTO = "PTO", DASH = "Dashboard", MAXROW = 300;
  class UserError extends Error {}
  class NeedPassword extends Error { constructor(sheets) { super("Sheet password needed"); this.sheets = sheets; } }
  class BadPassword extends Error { constructor() { super("That password didn't unlock the sheet."); } }

  const colL = i => { let s = ""; for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; };
  const colIdx = l => l.split("").reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
  const local = a => String(a || "").replace(/^.*!/, "").replace(/\$/g, "");

  async function readModel() {
    return Excel.run(async ctx => {
      const wb = ctx.workbook, wss = wb.worksheets;
      const pto = wss.getItemOrNullObject(PTO), dash = wss.getItemOrNullObject(DASH);
      const tA = wb.tables.getItemOrNullObject("Tbl_Yr_Assignments");
      const tW = wb.tables.getItemOrNullObject("Wk_Start_Date"), tR = wb.tables.getItemOrNullObject("Cov_Wk_Range");
      await ctx.sync();
      if (pto.isNullObject) throw new UserError("This workbook has no “PTO” sheet. Open the Call Pool workbook and try again.");
      const colA = pto.getRange("A1:A" + MAXROW).load("values");
      const used = pto.getUsedRange(true).load("columnIndex,columnCount");
      const head = pto.getRange("B2").load("formulas");
      pto.protection.load("protected");
      let dr = null;
      if (!dash.isNullObject) { dr = dash.getRange("B3:D14").load("values,formulas"); dash.protection.load("protected"); }
      const av = tA.isNullObject ? null : tA.getRange().load("values");
      const wv = tW.isNullObject ? null : tW.getDataBodyRange().load("values");
      const rv = tR.isNullObject ? null : tR.getDataBodyRange().load("values");
      await ctx.sync();

      const sections = CP.parseSections(colA.values.map(r => r[0]));
      if (!sections.length) throw new UserError("Couldn't find the “Call Pool 1”, “Call Pool 2” … sections in column A of the PTO sheet.");
      const lastCol = used.columnIndex + used.columnCount - 1, dateRow = sections[0].dateRow;
      const dates = pto.getRangeByIndexes(dateRow - 1, 0, 1, lastCol + 1).load("values");
      const first = sections[0].first;
      const probe = pto.getRange(`B${first}:C${first}`).load("formulas");
      await ctx.sync();

      const dateCol = new Map(); let minDate = Infinity, maxDate = -Infinity;
      dates.values[0].forEach((v, i) => { if (i >= 3 && typeof v === "number" && v > 30000) { if (!dateCol.has(v)) dateCol.set(v, i); minDate = Math.min(minDate, v); maxDate = Math.max(maxDate, v); } });
      const weeks = [];
      if (wv && rv) wv.values.forEach((r, i) => { const mon = r[0], label = String((rv.values[i] || [])[0] || "").trim(); if (typeof mon === "number" && label) weeks.push({ label, mon }); });
      const assign = {};
      if (av) av.values.slice(1).forEach(r => { const k = String(r[0] || "").trim(); if (k) assign[k] = r.slice(1, 9).map(x => CP.xtrim(x)); });
      const d = dr ? {
        week: String(dr.values[0][0] || "").trim(),
        rows: dr.values.slice(4).map((r, i) => ({ row: 7 + i, num: +r[0], onCall: CP.xtrim(r[1]), covering: String(r[2] == null ? "" : r[2]), formula: String(dr.formulas[4 + i][2] || "") })),
        protected: dash.protection.protected
      } : null;
      return {
        sections, dateRow, dateCol, minDate, maxDate, lastCol, weeks, assign, dash: d,
        ptoProtected: pto.protection.protected,
        installed: { pto: CP.isOurs(probe.formulas[0][0], "pto") && CP.isOurs(probe.formulas[0][1], "leave") && String(head.formulas[0][0]).startsWith("="),
          dash: !!d && d.rows.every(r => !(r.num >= 1 && r.num <= 8) || CP.isOurs(r.formula, "covering")) },
        tables: { assign: !!av, weeks: !!(wv && rv) }
      };
    });
  }

  /* the 7 columns Saturday..Friday for every PTO row: [[sat..fri] per row 1..MAXROW] */
  async function readWeek(model, sat) {
    const c = model.dateCol.get(sat);
    if (c == null || model.dateCol.get(sat + 6) !== c + 6) return null;
    return Excel.run(async ctx => {
      const r = ctx.workbook.worksheets.getItem(PTO).getRangeByIndexes(0, c, MAXROW, 7).load("values");
      await ctx.sync();
      return r.values;
    });
  }

  /* one provider row from..to as [{serial, value}]; dates must be consecutive columns */
  function span(model, from, to) {
    const c0 = model.dateCol.get(from), c1 = model.dateCol.get(to);
    if (c0 == null || c1 == null) throw new UserError(`Dates must be between ${CP.mdy(model.minDate)} and ${CP.mdy(model.maxDate)} (the PTO sheet's date range).`);
    if (c1 - c0 !== to - from) throw new UserError("The PTO sheet's date columns aren't one per day in this range, so the add-in won't write here. Check row " + model.dateRow + ".");
    return { c0, n: c1 - c0 + 1 };
  }
  async function readRow(model, row, from, to) {
    const { c0, n } = span(model, from, to);
    return Excel.run(async ctx => {
      const r = ctx.workbook.worksheets.getItem(PTO).getRangeByIndexes(row - 1, c0, 1, n).load("values");
      await ctx.sync();
      return r.values[0].map((v, i) => ({ serial: from + i, value: v }));
    });
  }

  /* run fn with the named sheets unprotected; locks them again with their own options afterwards */
  async function withUnlocked(names, password, fn) {
    return Excel.run(async ctx => {
      const sheets = names.map(n => ctx.workbook.worksheets.getItemOrNullObject(n));
      await ctx.sync();
      const real = sheets.map((s, i) => ({ s, name: names[i] })).filter(x => !x.s.isNullObject);
      real.forEach(x => x.s.protection.load("protected,options"));
      await ctx.sync();
      const locked = real.filter(x => x.s.protection.protected).map(x => ({ ...x, opts: Object.assign({}, x.s.protection.options) }));
      if (locked.length && !password) throw new NeedPassword(locked.map(x => x.name));
      locked.forEach(x => x.s.protection.unprotect(password));
      try { await ctx.sync(); }
      catch (e) {
        locked.forEach(x => x.s.protection.load("protected")); await ctx.sync();
        locked.filter(x => !x.s.protection.protected).forEach(x => x.s.protection.protect(x.opts, password));
        await ctx.sync();
        throw new BadPassword();
      }
      try { const out = await fn(ctx); await ctx.sync(); return out; }
      finally { locked.forEach(x => x.s.protection.protect(x.opts, password)); await ctx.sync(); }
    });
  }

  async function writeRow(model, row, from, values, password) {
    const { c0 } = span(model, from, from + values.length - 1);
    return withUnlocked([PTO], password, async ctx => {
      ctx.workbook.worksheets.getItem(PTO).getRangeByIndexes(row - 1, c0, 1, values.length).values = [values];
    });
  }

  /* puts a merged header that starts in B (e.g. "January" over B1:AF2) back over D.. so B:C are free */
  async function freeHeader(ctx, ws, cell) {
    const ma = ws.getRange(cell).getMergedAreasOrNullObject(); ma.load("address");
    const v = ws.getRange(cell).load("values");
    await ctx.sync();
    if (ma.isNullObject) return;
    const m = local(ma.address).match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/); if (!m) return;
    ws.getRange(local(ma.address)).unmerge();
    if (colIdx(m[3]) > 2 && m[1] === "B") {
      const label = v.values[0][0];
      ws.getRange(`B${m[2]}:C${m[4]}`).values = Array.from({ length: +m[4] - +m[2] + 1 }, () => ["", ""]);
      const rest = ws.getRange(`D${m[2]}:${m[3]}${m[4]}`); rest.merge(false);
      ws.getRange(`D${m[2]}`).values = [[label]];
    }
  }

  const CF_PTO = '=AND(B3<>"",$A3<>"",LEFT(TRIM($A3),10)<>"Call Pool ")';
  const CF_LEAVE = '=AND(C3<>"",$A3<>"",LEFT(TRIM($A3),10)<>"Call Pool ")';

  /* PTO status columns B:C + Dashboard Covering formulas; safe to run again */
  async function install(model, password) {
    return withUnlocked([PTO, DASH], password, async ctx => {
      const pto = ctx.workbook.worksheets.getItem(PTO);
      const bc = pto.getRange("B:C"); bc.columnHidden = false; bc.format.columnWidth = 150;
      await freeHeader(ctx, pto, "B1"); await freeHeader(ctx, pto, "B2");
      pto.getRange("B1:C1").merge(false); pto.getRange("B2:C2").merge(false);
      pto.getRange("B1").values = [["Out this coverage week"]];
      pto.getRange("B2").formulas = [[model.dash ? CP.FORMULAS.week : ""]];
      const top = pto.getRange("B1:C2"); top.format.font.bold = true; top.format.horizontalAlignment = "Center"; top.format.fill.color = "#F2F2F2";
      for (const s of model.sections) {
        const h = pto.getRange(`B${s.head}:C${s.head}`); h.values = [["PTO", "Leave"]]; h.format.font.bold = true; h.format.horizontalAlignment = "Center";
        const rows = []; for (let r = s.first; r <= s.last; r++) rows.push([CP.FORMULAS.pto(r, model.dateRow), CP.FORMULAS.leave(r, model.dateRow)]);
        const body = pto.getRange(`B${s.first}:C${s.last}`);
        body.formulas = rows; body.format.horizontalAlignment = "Left"; body.format.wrapText = false; body.format.font.bold = false;
      }
      // light versions of the grid's own colours (PTO pink, Leave green), added once
      const cfs = pto.getRange(`B3:C${MAXROW}`).conditionalFormats; cfs.load("items/type");
      await ctx.sync();
      const customs = cfs.items.filter(c => c.type === "Custom"); customs.forEach(c => c.custom.rule.load("formula"));
      await ctx.sync();
      const have = new Set(customs.map(c => c.custom.rule.formula));
      [[`B3:B${MAXROW}`, CF_PTO, "#FFD6DE", "#7A1029"], [`C3:C${MAXROW}`, CF_LEAVE, "#C8F5E1", "#0B5B3A"]].forEach(([addr, f, fill, font]) => {
        if (have.has(f)) return;
        const cf = pto.getRange(addr).conditionalFormats.add(Excel.ConditionalFormatType.custom);
        cf.custom.rule.formula = f; cf.custom.format.fill.color = fill; cf.custom.format.font.color = font;
      });
      pto.freezePanes.freezeAt(pto.getRange("A1:C2"));

      if (model.dash) {
        const dash = ctx.workbook.worksheets.getItem(DASH);
        for (const r of model.dash.rows) if (r.num >= 1 && r.num <= 8) dash.getRange(`D${r.row}`).formulas = [[CP.FORMULAS.covering(r.row)]];
        dash.getRange("D7:E14").format.wrapText = true;
      }
      await ctx.sync();
      if (model.dash) await fitDashboard(ctx);
    });
  }

  /* merged cells don't auto-fit, so size rows 7-14 to the Covering text */
  async function fitDashboard(ctx) {
    const dash = ctx.workbook.worksheets.getItem(DASH);
    const vals = dash.getRange("D7:D14").load("values"), d = dash.getRange("D7").load("format/columnWidth"), e = dash.getRange("E7").load("format/columnWidth");
    await ctx.sync();
    const cpl = (d.format.columnWidth + e.format.columnWidth) / 6.2;
    vals.values.forEach((r, i) => { dash.getRange(`A${7 + i}`).format.rowHeight = CP.rowHeightFor(r[0], cpl, 39.75); });
  }
  async function refreshDashboard(password) { return withUnlocked([DASH], password, async ctx => { await fitDashboard(ctx); }); }

  async function goTo(sheet, address) {
    return Excel.run(async ctx => { const ws = ctx.workbook.worksheets.getItem(sheet); ws.activate(); ws.getRange(address).select(); await ctx.sync(); });
  }

  return { PTO, DASH, MAXROW, UserError, NeedPassword, BadPassword, readModel, readWeek, readRow, writeRow, install, refreshDashboard, goTo, colL };
})();
if (typeof module === "object" && module.exports) module.exports = XL;
