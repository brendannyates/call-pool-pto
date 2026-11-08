# Call Pool PTO

Excel task-pane add-in for the provider Call Pool workbook. It logs provider time off on the **PTO** tab, shows who is out in each call pool for a coverage week (Saturday to Friday), flags on-call providers who are also out, and fills **Dashboard › Covering** with live formulas.

The hosted files hold no PHI and no provider data. Everything is read from and written to the open workbook.

## What it does

- **This week**: one card per call pool with the on-call provider (from `Tbl_Yr_Assignments`) and everyone out that week, with PTO, ½ day PTO, call out and leave dates. A provider on call who is also out is flagged, so you can add an override.
- **Log time off**: pick a provider, a type (PTO, ½ day PTO, Leave, Call out, Clear) and dates. The preview lists the days it will mark. It skips weekends unless you untick *Weekdays only*. It never overwrites holidays (H) or days off (O, OFF), and it warns if the provider is on call that week. It writes the PTO tab's own codes: `P`, `1/2 Day P`, `L`, `C`.
- **Setup › Install formulas** (also **Send to Dashboard**):
  - PTO tab: unhides columns B:C, which were empty and hidden (Jan–Jul 2026, no dates), and adds a **PTO** and a **Leave** column next to every name for the Dashboard week. `B2 = Dashboard!$B$3`. It moves the January header to D1, adds light pink and green highlights, and freezes A1:C2.
  - Dashboard: `D7:D14` (Covering) reads those two columns for each pool, for example `Dr. A PTO 10/26 to 10/30 & NP B Leave 10/26 to 10/30`. The on-call provider shows with `⚠ ON CALL:` when out.
  - Typed Covering notes it replaces are copied into the workbook's add-in settings (`cpCoveringBackup`).
  - Safe to run again.

## Name screen and Dev Update bell

On open, a loading screen ticks off each read. Then, if nobody has picked a name on this computer today, the pane asks "Enter your name". The list of names is kept in the workbook's add-in settings (`cpUsers`), so no sheet is added; today's pick is kept in this computer's browser storage, per workbook. **Switch user** at the top changes it; Setup lists the names and can remove one.

The bell at the right end of the header (`dev-updates.js`, `.css`, `.json`, shared with the IVK Tracker) shows what changed. Add an entry to the top of `dev-updates.json` with each release. On the loading and name screens, which have no toolbar, it floats top right and can be dragged.

## Protected sheets

The PTO and Dashboard sheets are password-protected. When a save needs them, the pane asks for the sheet password, unprotects, writes, then protects again with the same options. The password is kept only in memory: for one save, or until the pane closes if you tick *Remember*. It is never stored or sent.

## Formulas

`core.js` (`FORMULAS`) holds the exact formulas; Setup shows them with Copy buttons. They need Excel 365 (LET, LAMBDA, XLOOKUP, HSTACK; the workbook already uses these). They were checked in desktop Excel against the add-in's own logic on 7 weeks of a real copy (623 matching cells).

## Files

`core.js` rules (node tests in `tests/`), `excel.js` workbook I/O, `app.js` pane UI, `style.css` MHS add-in tokens, `manifest.xml` (ExcelApi 1.13).

```
npm run check
npm test
```
