// node --test tests/*.test.js   (invented names only)
const test = require("node:test"), assert = require("node:assert/strict");
const CP = require("../core.js");
const S = iso => CP.isoToSerial(iso);

test("dates round-trip and weekday labels", () => {
  assert.equal(CP.serialToISO(S("2026-10-24")), "2026-10-24");
  assert.equal(S("2026-08-01"), 46235);
  assert.equal(CP.dayLabel(S("2026-10-26")), "Mon 10/26");
  assert.equal(CP.satOfMonday(S("2026-10-26")), S("2026-10-24"));
  assert.equal(CP.satOnOrBefore(S("2026-10-29")), S("2026-10-24"));
  assert.equal(CP.satOnOrBefore(S("2026-10-24")), S("2026-10-24"));
  assert.equal(CP.localSerial(new Date(2026, 10, 3, 23, 30)), S("2026-11-03"));
});

test("codes", () => {
  assert.equal(CP.typeOf("P"), "pto"); assert.equal(CP.typeOf(" p "), "pto");
  assert.equal(CP.typeOf("1/2 Day P"), "half"); assert.equal(CP.typeOf("L"), "leave");
  assert.equal(CP.typeOf("C"), "callout"); assert.equal(CP.typeOf("H"), "holiday");
  assert.equal(CP.typeOf("OFF"), "other"); assert.equal(CP.typeOf("O"), "other"); assert.equal(CP.typeOf(""), "");
});

const colA = ["PTO Log", "", "Call Pool 1", "", "Providers:", "Dr. Ada  Lark", "NP Bo Wren ", "Employee ", "", "Call Pool 2", "", "Providers:", "Dr. Cy Moss", "Employee", "", ""];
test("parseSections finds pools, provider rows and placeholders", () => {
  const s = CP.parseSections(colA);
  assert.equal(s.length, 2);
  assert.deepEqual([s[0].pool, s[0].head, s[0].dateRow, s[0].first, s[0].last], ["Call Pool 1", 3, 4, 6, 9]);
  assert.deepEqual(s[0].rows.map(r => [r.row, r.name, r.placeholder]), [[6, "Dr. Ada Lark", false], [7, "NP Bo Wren", false], [8, "Employee", true]]);
  assert.deepEqual([s[1].first, s[1].last], [13, 14]);
  assert.deepEqual(CP.providers(s).map(p => p.name), ["Dr. Ada Lark", "NP Bo Wren", "Dr. Cy Moss"]);
});

test("weekStatus: runs, half days, call outs, leave; holidays and days off ignored", () => {
  const sat = S("2026-10-24");
  let st = CP.weekStatus(["", "", "P", "P", "P", "P", "P"], sat);
  assert.equal(st.pto, "PTO 10/26 to 10/30"); assert.equal(st.leave, "");
  st = CP.weekStatus(["", "", "", "P", "", "P", "P"], sat);
  assert.equal(st.pto, "PTO 10/27, 10/29 to 10/30");
  st = CP.weekStatus(["", "", "P", "1/2 Day P", "C", "H", "O"], sat);
  assert.equal(st.pto, "PTO 10/26; ½ day PTO 10/27; Call out 10/28");
  st = CP.weekStatus(["", "", "L", "L", "OFF", "L", "L"], sat);
  assert.equal(st.leave, "10/26 to 10/27, 10/29 to 10/30"); assert.equal(st.pto, "");
  assert.deepEqual(st.items.map(i => i.type), ["leave"]);
  assert.equal(CP.weekStatus([], sat).pto, "");
});

test("coveringText matches the Dashboard formula, flags the on-call provider", () => {
  const e = [{ name: "Dr. Ada Lark", pto: "PTO 10/26", leave: "" }, { name: "NP Bo Wren", pto: "", leave: "" }, { name: "Dr. Cy Moss", pto: "PTO 10/26", leave: "10/27 to 10/30" }];
  assert.equal(CP.coveringText(e, "Dr. Bo Wren"), "Dr. Ada Lark PTO 10/26 & Dr. Cy Moss PTO 10/26; Leave 10/27 to 10/30");
  assert.equal(CP.coveringText(e, "Dr. Cy MOSS"), "Dr. Ada Lark PTO 10/26 & ⚠ ON CALL: Dr. Cy Moss PTO 10/26; Leave 10/27 to 10/30");
  assert.equal(CP.coveringText([], "x"), "");
  assert.ok(CP.sameLast("Dr. Josh Lark", "Dr. Joshua Lark ")); assert.ok(!CP.sameLast("", ""));
});

test("planWrite: weekdays only, never over holidays or days off, replaces other codes", () => {
  const from = S("2026-11-21"); // Sat
  const cells = ["", "", "", "L", "", "H", "H", "", ""].map((v, i) => ({ serial: from + i, value: v }));
  const p = CP.planWrite(cells, "pto", { weekdaysOnly: true });
  assert.deepEqual(p.values, ["", "", "P", "P", "P", "H", "H", "", ""]);
  assert.deepEqual(p.changes.map(c => [CP.dayLabel(c.serial), c.from, c.to]), [["Mon 11/23", "", "P"], ["Tue 11/24", "L", "P"], ["Wed 11/25", "", "P"]]);
  assert.deepEqual(p.skipped.map(s => s.why), ["weekend", "weekend", "holiday", "holiday", "weekend", "weekend"]);
  const all = CP.planWrite(cells, "leave", { weekdaysOnly: false });
  assert.equal(all.values[0], "L"); assert.equal(all.values[5], "H");
  const off = CP.planWrite([{ serial: S("2026-11-23"), value: "OFF" }], "pto", {});
  assert.deepEqual([off.values[0], off.skipped[0].why], ["OFF", "day off (OFF)"]);
  const half = CP.planWrite([{ serial: S("2026-11-23"), value: "" }], "half", {});
  assert.equal(half.values[0], "1/2 Day P");
  const same = CP.planWrite([{ serial: S("2026-11-23"), value: "p" }], "pto", {});
  assert.equal(same.changes.length, 0);
});

test("planWrite clear removes time off only", () => {
  const from = S("2026-11-23");
  const p = CP.planWrite(["P", "1/2 Day P", "H", "L", "OFF", "C"].map((v, i) => ({ serial: from + i, value: v })), "clear", {});
  assert.deepEqual(p.values, ["", "", "H", "", "OFF", ""]);
  assert.equal(p.changes.length, 4);
});

test("timeOffRuns joins across weekends and holidays, splits on a worked weekday or a new type", () => {
  const from = S("2026-11-19"); // Thu
  const v = ["P", "P", "", "", "P", "P", "P", "H", "H", "", "", "L", "L", "", "P"];
  const r = CP.timeOffRuns(v.map((x, i) => ({ serial: from + i, value: x })));
  assert.deepEqual(r.map(x => [x.type, x.text, x.days]), [["pto", "Thu 11/19 to Wed 11/25", 5], ["leave", "Mon 11/30 to Tue 12/1", 2], ["pto", "Thu 12/3", 1]]);
});

test("onCallWeeks and weekForSerial", () => {
  const weeks = [{ label: "10/24-10/30", mon: S("2026-10-26") }, { label: "10/31-11/06", mon: S("2026-11-02") }];
  const assign = { "10/24-10/30": ["NP Bo Wren", "x"], "10/31-11/06": ["Dr. Ada Lark"] };
  assert.deepEqual(CP.onCallWeeks("Dr. Ada Lark", 1, S("2026-10-30"), S("2026-11-03"), weeks, assign).map(w => w.label), ["10/31-11/06"]);
  assert.deepEqual(CP.onCallWeeks("Dr. Ada Lark", 2, S("2026-10-30"), S("2026-11-03"), weeks, assign), []);
  assert.equal(CP.weekForSerial(weeks, S("2026-10-31")).label, "10/31-11/06");
  assert.equal(CP.weekForSerial(weeks, S("2026-12-31")), null);
});

test("formulas: row numbers, date row, signatures", () => {
  const f = CP.FORMULAS.pto(19, 4);
  assert.ok(f.startsWith("=LET(") && f.includes("$D19:$XFD19") && f.includes("$D$4:$XFD$4") && !f.includes("{R}"));
  assert.ok(CP.FORMULAS.leave(6, 7).includes("$D$7:$XFD$7"));
  assert.ok(CP.FORMULAS.covering(9).includes('"Call Pool " & $B9') && CP.FORMULAS.covering(9).includes("last($C9)"));
  assert.ok(CP.isOurs(CP.FORMULAS.pto(6), "pto") && CP.isOurs(CP.FORMULAS.leave(6), "leave") && CP.isOurs(CP.FORMULAS.covering(7), "covering"));
  assert.ok(!CP.isOurs("=INDEX(x)", "covering") && !CP.isOurs("", "pto"));
  for (const x of [CP.FORMULAS.pto(6), CP.FORMULAS.leave(6), CP.FORMULAS.covering(7)]) assert.equal((x.match(/\(/g) || []).length, (x.match(/\)/g) || []).length);
});

test("people list: clean, dedupe, add, remove", () => {
  assert.deepEqual(CP.userList([{ name: "  Avery  " }, { name: "avery" }, "Jordan", null, { name: "" }]).map(u => u.name), ["Avery", "Jordan"]);
  assert.deepEqual(CP.userList("not a list"), []);
  assert.equal(CP.cleanName("  Sam   R  "), "Sam R");
  assert.equal(CP.cleanName("x".repeat(40)).length, 30);
  const L = CP.addUser([{ name: "Avery" }], "Jordan", "2026-11-07");
  assert.deepEqual(L, [{ name: "Avery", added: "" }, { name: "Jordan", added: "2026-11-07" }]);
  assert.equal(CP.addUser(L, "JORDAN").length, 2);
  assert.deepEqual(CP.removeUser(L, "avery").map(u => u.name), ["Jordan"]);
});

test("today's pick and picker order", () => {
  const rec = CP.whoRec(JSON.stringify(CP.whoPick(CP.whoRec(""), "Sam", "2026-11-07")));
  assert.equal(CP.whoToday(rec, "2026-11-07"), "Sam");
  assert.equal(CP.whoToday(rec, "2026-11-08"), "");
  assert.deepEqual(CP.whoRec("{bad json").used, []);
  const r2 = CP.whoPick(CP.whoPick(rec, "Ana", "2026-11-08"), "sam", "2026-11-09");
  assert.deepEqual(r2.used, ["sam", "Ana"]);
  const order = CP.pickOrder([{ name: "Zoe" }, { name: "Ana" }, { name: "Sam" }, { name: "beth" }], ["Sam"]);
  assert.deepEqual(order.map(u => u.name), ["Sam", "Ana", "beth", "Zoe"]);
  assert.deepEqual(order.map(u => u.here), [true, false, false, false]);
  assert.equal(CP.wbKeyOf("https://x/Call Pool.xlsx?web=1"), CP.wbKeyOf("HTTPS://X/call pool.xlsx"));
  assert.notEqual(CP.wbKeyOf("a.xlsx"), CP.wbKeyOf("b.xlsx"));
  assert.equal(CP.wbKeyOf(""), "");
});

test("rowHeightFor", () => {
  assert.equal(CP.rowHeightFor("", 50, 39.75), 39.75);
  assert.equal(CP.rowHeightFor("x".repeat(160), 50, 39.75), 72);
});
