/**
 * P16 — Google Sheets Automation Pack (SAMPLE)
 * DailyNeeds Mart, Lahore — weekly sales report that builds itself.
 *
 * INSTALL (2 minutes, in Google Sheets):
 *   1. Paste the client's daily sales into the `RawData` sheet with the
 *      headers: date | week_no | week_start | product | category |
 *      qty | unit_price_pkr | amount_pkr
 *   2. Extensions > Apps Script, paste this file, Save.
 *   3. Reload the sheet: a "Weekly Report" menu appears.
 *   4. Weekly Report > Build weekly report  — the report rebuilds.
 *   5. Weekly Report > Install Monday 08:00 trigger — from then on
 *      the report rebuilds itself every Monday morning, no clicks.
 *   6. Set OWNER_EMAIL below, then Weekly Report > Email this week's
 *      summary (first run asks for the one-time authorisation).
 *
 * WHAT IT BUILDS (sheets are created if missing):
 *   Weekly Report — one row per week (Mon start), one column per
 *     category, Total and week-on-week growth as live SUMIFS formulas.
 *   Top Products  — top 5 products by revenue, live SUMIFS formulas.
 *
 * HONEST NOTE: this script reads only the RawData sheet in the same
 * spreadsheet. It never invents numbers — every total is a formula
 * over the client's own rows. The Excel workbook in this repo
 * (outputs/weekly_report_SAMPLE.xlsx) is the offline demo of the
 * exact same output, so the buyer can see the report before ordering.
 */

// ── CONFIG: edit these two lines for the client ──────────────────────────
const OWNER_EMAIL = 'owner@example.com';   // where the weekly email goes
const CATEGORIES = ['Groceries', 'Dairy', 'Bakery',
                    'Household', 'Beverages', 'Personal Care'];
const RAW_SHEET = 'RawData';
const REPORT_SHEET = 'Weekly Report';
const TOP_SHEET = 'Top Products';
// ─────────────────────────────────────────────────────────────────────────

/** Adds the "Weekly Report" menu when the spreadsheet opens. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Weekly Report')
    .addItem('Build weekly report', 'buildWeeklyReport')
    .addItem('Email this week\'s summary', 'emailWeeklySummary')
    .addSeparator()
    .addItem('Install Monday 08:00 trigger', 'createWeeklyTrigger')
    .addItem('Remove weekly trigger', 'deleteWeeklyTrigger')
    .addToUi();
}

/** Creates the three sheets (with headers) if they are missing. */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let raw = ss.getSheetByName(RAW_SHEET);
  if (!raw) {
    raw = ss.insertSheet(RAW_SHEET);
    raw.getRange(1, 1, 1, 8).setValues([['date', 'week_no', 'week_start',
      'product', 'category', 'qty', 'unit_price_pkr', 'amount_pkr']]);
    raw.getRange(1, 1, 1, 8).setFontWeight('bold');
    raw.setFrozenRows(1);
  }
  if (!ss.getSheetByName(REPORT_SHEET)) ss.insertSheet(REPORT_SHEET);
  if (!ss.getSheetByName(TOP_SHEET)) ss.insertSheet(TOP_SHEET);
  return ss;
}

/**
 * Rebuilds the Weekly Report sheet from RawData.
 * One row per week, one SUMIFS column per category, a Total column
 * and a week-on-week growth column — all live formulas, so pasting
 * new rows into RawData and re-running is all the upkeep there is.
 * Returns the number of weeks written.
 */
function buildWeeklyReport() {
  const ss = setupSheets();
  const raw = ss.getSheetByName(RAW_SHEET);
  const lastRow = raw.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert(
      'RawData is empty — paste the daily sales first.');
    return 0;
  }

  // Distinct week_starts, in order, straight from the data.
  const starts = raw.getRange(2, 3, lastRow - 1, 1).getValues()
    .map(r => r[0])
    .filter((v, i, a) => a.indexOf(v) === i)
    .sort();
  const nWeeks = starts.length;

  const sheet = ss.getSheetByName(REPORT_SHEET);
  sheet.clear();
  const header = ['Week starting', 'Lines'].concat(CATEGORIES)
    .concat(['Total (PKR)', 'WoW growth', 'Top category']);
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  const rawRef = "'" + RAW_SHEET + "'";
  for (let i = 0; i < nWeeks; i++) {
    const r = i + 2; // sheet row
    const startCell = "A" + r;
    // A: week start (as a real date, formatted), B: line count
    sheet.getRange(r, 1).setValue(new Date(starts[i]));
    sheet.getRange(r, 1).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(r, 2).setFormula(
      '=COUNTIF(' + rawRef + '!$C$2:$C$' + lastRow + ',' + startCell + ')');
    // C..H: one SUMIFS per category
    CATEGORIES.forEach((cat, j) => {
      const col = String.fromCharCode(67 + j); // C..H
      sheet.getRange(r, 3 + j).setFormula(
        '=SUMIFS(' + rawRef + '!$H$2:$H$' + lastRow + ',' +
        rawRef + '!$C$2:$C$' + lastRow + ',' + startCell + ',' +
        rawRef + '!$E$2:$E$' + lastRow + ',"' + cat + '")');
      sheet.getRange(r, 3 + j).setNumberFormat('#,##0');
    });
    // I: total = sum of the category columns
    sheet.getRange(r, 9).setFormula('=SUM(C' + r + ':H' + r + ')');
    sheet.getRange(r, 9).setNumberFormat('#,##0');
    sheet.getRange(r, 9).setFontWeight('bold');
    // J: week-on-week growth (blank for week 1)
    if (i === 0) {
      sheet.getRange(r, 10).setValue('—');
    } else {
      sheet.getRange(r, 10).setFormula('=IF(I' + (r - 1) + '=0,"—",I' + r +
        '/I' + (r - 1) + '-1)');
      sheet.getRange(r, 10).setNumberFormat('0.00%');
    }
    // K: top category this week (INDEX+MATCH over the category row)
    sheet.getRange(r, 11).setFormula(
      '=INDEX($C$1:$H$1,1,MATCH(MAX(C' + r + ':H' + r + '),C' + r +
      ':H' + r + ',0))');
  }

  // Grand total row
  const t = nWeeks + 2;
  sheet.getRange(t, 1).setValue('TOTAL (all weeks)');
  sheet.getRange(t, 1).setFontWeight('bold');
  sheet.getRange(t, 2).setFormula('=SUM(B2:B' + (nWeeks + 1) + ')');
  for (let j = 0; j < CATEGORIES.length; j++) {
    const col = String.fromCharCode(67 + j);
    sheet.getRange(t, 3 + j).setFormula(
      '=SUM(' + col + '2:' + col + (nWeeks + 1) + ')');
    sheet.getRange(t, 3 + j).setNumberFormat('#,##0');
  }
  sheet.getRange(t, 9).setFormula('=SUM(I2:I' + (nWeeks + 1) + ')');
  sheet.getRange(t, 9).setNumberFormat('#,##0');
  sheet.getRange(t, 9).setFontWeight('bold');

  buildTopProducts(ss, raw, lastRow);
  sheet.autoResizeColumns(1, header.length);
  return nWeeks;
}

/** Top 5 products by revenue, as live SUMIFS formulas. */
function buildTopProducts(ss, raw, lastRow) {
  const sheet = ss.getSheetByName(TOP_SHEET);
  sheet.clear();
  const products = raw.getRange(2, 4, lastRow - 1, 1).getValues()
    .map(r => r[0])
    .filter((v, i, a) => a.indexOf(v) === i)
    .sort();
  const rawRef = "'" + RAW_SHEET + "'";
  sheet.getRange(1, 1, 1, 4)
    .setValues([['Product', 'Category', 'Units sold', 'Revenue (PKR)']]);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  sheet.setFrozenRows(1);
  products.forEach((p, i) => {
    const r = i + 2;
    sheet.getRange(r, 1).setValue(p);
    // category: first matching row's category (products never change)
    sheet.getRange(r, 2).setFormula(
      '=INDEX(' + rawRef + '!$E$2:$E$' + lastRow +
      ',MATCH(A' + r + ',' + rawRef + '!$D$2:$D$' + lastRow + ',0))');
    sheet.getRange(r, 3).setFormula(
      '=SUMIF(' + rawRef + '!$D$2:$D$' + lastRow + ',A' + r + ',' +
      rawRef + '!$F$2:$F$' + lastRow + ')');
    sheet.getRange(r, 4).setFormula(
      '=SUMIF(' + rawRef + '!$D$2:$D$' + lastRow + ',A' + r + ',' +
      rawRef + '!$H$2:$H$' + lastRow + ')');
    sheet.getRange(r, 4).setNumberFormat('#,##0');
  });
  // Sort the table by revenue, highest first (keeps formulas intact).
  sheet.getRange(2, 1, products.length, 4).sort({ column: 4, ascending: false });
  // Keep only the top 5 rows, so the sheet answers the question at a glance.
  if (products.length > 5) {
    sheet.deleteRows(7, products.length - 5);
  }
  sheet.autoResizeColumns(1, 4);
}

/** Emails the latest week's headline numbers to the owner. */
function emailWeeklySummary() {
  const ss = setupSheets();
  const sheet = ss.getSheetByName(REPORT_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('Build the weekly report first.');
    return;
  }
  const week = sheet.getRange(lastRow - 1, 1, 1, 11).getValues()[0];
  const totalRow = sheet.getRange(lastRow, 1, 1, 11).getValues()[0];
  const body = [
    'Weekly sales summary — DailyNeeds Mart',
    '',
    'Latest week starting ' + Utilities.formatDate(
      new Date(week[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd') +
      ': PKR ' + Number(week[8]).toLocaleString('en-PK'),
    'Week-on-week growth: ' +
      (typeof week[9] === 'number'
        ? (week[9] * 100).toFixed(2) + '%'
        : week[9]),
    'Top category this week: ' + week[10],
    '',
    'Total so far (all weeks): PKR ' +
      Number(totalRow[8]).toLocaleString('en-PK'),
    '',
    'Built automatically by the Weekly Report script — open the sheet ' +
      'for the full breakdown.',
  ].join('\n');
  MailApp.sendEmail(OWNER_EMAIL,
    'Weekly sales summary — week of ' +
    Utilities.formatDate(new Date(week[0]),
      Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    body);
}

/** Installs a Monday 08:00 trigger that rebuilds the report weekly. */
function createWeeklyTrigger() {
  deleteWeeklyTrigger();
  ScriptApp.newTrigger('buildWeeklyReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
  SpreadsheetApp.getUi().alert(
    'Weekly trigger installed — the report rebuilds every Monday 08:00.');
}

/** Removes any weekly rebuild trigger created by this script. */
function deleteWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'buildWeeklyReport') {
      ScriptApp.deleteTrigger(t);
    }
  });
}
