/**
 * Indian Mutual Fund Screener: Google Sheets engine.
 *
 * The weekly pipeline (GitHub: pipeline/mf_pipeline.py) publishes CSVs to the
 * data folder named in Settings > Data feed URL. This script:
 *   1. builds the workbook (setupScreener, run once),
 *   2. checks hourly for a new pipeline run and imports it,
 *   3. applies the persistence rules (NEW > QUALIFIED > WATCH > REVIEW > REMOVED),
 *   4. logs changes, keeps monthly history and rebuilds the archive.
 * Screening thresholds and weights live on the Settings tab; the screen itself is
 * live formulas on All Funds, so editing Settings re-screens instantly.
 */

const TAB = {
  dash: 'Dashboard', all: 'All Funds', qual: 'Qualified Funds', rank: 'Rankings', watch: 'Watchlist',
  port: 'Portfolio', chg: 'Changes', hist: 'Historical Data', arch: 'Archive', qnav: 'Quarterly NAV',
  qaum: 'Quarterly AUM', set: 'Settings', src: 'Data Sources', state: '_State'
};
const TAB_ORDER = ['dash', 'all', 'qual', 'rank', 'watch', 'port', 'chg', 'hist', 'arch', 'qnav', 'qaum', 'set', 'src', 'state'];
const SHEET_ID = '1snjuSUqgrIlSX1o2CtEqV-Wgp5gB3mPffiCEZ0yKrH0';
/** Works whether the script is bound to the Sheet or created standalone at script.google.com. */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SHEET_ID); }
const DEFAULT_DATA_URL = 'https://raw.githubusercontent.com/AlanP13/mf-screener/main/data/';

// ---------------------------------------------------------------- column model
// [key, header, csv field, type]   type: int|num|num2|num1|pct|date|str
const FUND_COLS = [
  ['code', 'Scheme Code', 'code', 'int'], ['name', 'Scheme Name', 'name', 'str'], ['amc', 'AMC', 'amc', 'str'],
  ['broad', 'Broad Category', 'broad', 'str'], ['cat', 'Category', 'category', 'str'],
  ['plan', 'Plan', 'plan', 'str'], ['opt', 'Option', 'option', 'str'],
  ['incep', 'Inception Date', 'inception', 'date'], ['age', 'Age (yrs)', 'age_years', 'num1'],
  ['nav', 'NAV', 'nav', 'num'], ['navdate', 'NAV Date', 'nav_date', 'date'],
  ['hi', '52W High NAV', 'high_52w', 'num'], ['lo', '52W Low NAV', 'low_52w', 'num'],
  ['offhi', 'vs 52W High', 'pct_from_high', 'pct'],
  ['aum', 'AUM (Rs Cr)', 'aum_cr', 'int'], ['aumsrc', 'AUM Source', 'aum_source', 'str'],
  ['aaum', 'Avg AUM Last Qtr (Rs Cr)', 'aaum_latest_q_cr', 'int'],
  ['ter', 'Expense Ratio (Regular)', 'expense_ratio', 'pct2'], ['mgr', 'Fund Manager', 'fund_manager', 'str'],
  ['r3m', '3M', 'ret_3m', 'pct'], ['r6m', '6M', 'ret_6m', 'pct'], ['r1', '1Y', 'ret_1y', 'pct'],
  ['r2', '2Y', 'ret_2y', 'pct'], ['r3', '3Y', 'ret_3y', 'pct'], ['r5', '5Y', 'ret_5y', 'pct'],
  ['r10', '10Y', 'ret_10y', 'pct'], ['rsi', 'Since Inception', 'ret_si', 'pct'], ['sibasis', 'SI Basis', 'si_source', 'str'],
  ['r1c', '1Y vs Category Median', 'ret_1y_vs_cat', 'pct'], ['r3c', '3Y vs Category Median', 'ret_3y_vs_cat', 'pct'],
  ['vol', 'Volatility 3Y', 'vol_3y', 'pct'], ['mdd', 'Max Drawdown 3Y', 'mdd_3y', 'pct'],
  ['sharpe', 'Sharpe 3Y', 'sharpe_3y', 'num2'], ['sortino', 'Sortino 3Y', 'sortino_3y', 'num2'],
  ['cons', 'Consistency 5Y', 'consistency_5y', 'pct'],
  ['psharpe', 'Sharpe Pctile', 'pct_sharpe', 'num0'], ['pmdd', 'Drawdown Pctile', 'pct_mdd', 'num0'],
  ['pcons', 'Consistency Pctile', 'pct_consistency', 'num0'], ['peers', 'Peers', 'peer_count', 'int'],
  ['flags', 'Data Flags', 'data_flags', 'str'], ['ddate', 'Data Date', 'data_date', 'date']
];
const STATUS_COLS = [['status', 'Status', 'str'], ['streak', 'Consecutive Months Qualified', 'int'],
  ['since', 'Status Since', 'date'], ['months', 'Months Qualified (total)', 'int']];
const FORMULA_COLS = [['appl', 'Checks Applicable', 'int'], ['fail', 'Checks Failed', 'int'],
  ['failb', 'Checks Failed (near-miss band)', 'int'], ['aumok', 'AUM OK', 'str'], ['elig', 'In Screened Categories', 'str'],
  ['result', 'Screen Result', 'str'], ['perf', 'Performance Score', 'pct'], ['pperf', 'Performance Pctile', 'num0'],
  ['quality', 'Quality Score', 'num1'], ['rank', 'Rank', 'int'], ['signal', 'Allocation Signal', 'str']];

const COL = {};   // key -> 1-based column index on All Funds
(function () {
  let i = 1;
  FUND_COLS.forEach(c => { COL[c[0]] = i++; });
  STATUS_COLS.forEach(c => { COL[c[0]] = i++; });
  FORMULA_COLS.forEach(c => { COL[c[0]] = i++; });
})();
const N_DATA = FUND_COLS.length, N_STATUS = STATUS_COLS.length, N_ALL = N_DATA + N_STATUS + FORMULA_COLS.length;

function L(key) { return colLetter(COL[key]); }
function R(key) { return L(key) + '2:' + L(key); }
function colLetter(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }

// ---------------------------------------------------------------- settings model
// [named range, label, default, format, note]; a string in the first slot alone is a section heading
const SETTINGS = [
  ['Stage 1 screen: minimum returns and size'],
  ['MIN_1Y', 'Minimum 1Y return', 0.20, 'pct', 'Point-to-point for 1Y; CAGR for 2Y and longer. Clear any threshold cell to switch that check off.'],
  ['MIN_2Y', 'Minimum 2Y return (CAGR)', 0.20, 'pct', ''],
  ['MIN_3Y', 'Minimum 3Y return (CAGR)', 0.20, 'pct', ''],
  ['MIN_5Y', 'Minimum 5Y return (CAGR)', 0.20, 'pct', ''],
  ['MIN_10Y', 'Minimum 10Y return (CAGR)', 0.20, 'pct', ''],
  ['MIN_SI', 'Minimum since-inception return (CAGR)', 0.20, 'pct', ''],
  ['MIN_AUM', 'Minimum AUM (Rs crore)', 10000, 'int', 'Fund-level AUM across all plans. A scale filter, not part of the score.'],
  ['MIN_HIST_YEARS', 'Emerging qualifier: minimum fund age (years)', 3, 'num1', 'Younger funds are screened only on the periods they have; missing periods are N/A, not failures.'],
  ['SCREEN_CATEGORIES', 'Categories to screen', 'All', 'text', 'All, or a comma list of broad categories (Equity, Hybrid, Debt, Index Funds, FoF, Solution Oriented) or categories (Flexi Cap, Small Cap ...).'],
  ['Performance score weights (renormalised over the periods a fund has)'],
  ['W_1Y', '1Y weight', 0.10, 'pct', ''], ['W_2Y', '2Y weight', 0.15, 'pct', ''], ['W_3Y', '3Y weight', 0.20, 'pct', ''],
  ['W_5Y', '5Y weight', 0.25, 'pct', ''], ['W_10Y', '10Y weight', 0.20, 'pct', ''], ['W_SI', 'Since-inception weight', 0.10, 'pct', ''],
  ['Quality score weights (category percentiles, 0-100)'],
  ['Q_W_PERF', 'Performance score percentile', 0.40, 'pct', 'Stage 1: returns'],
  ['Q_W_CONS', 'Consistency percentile', 0.20, 'pct', 'Stage 2: share of monthly rolling 1Y windows (last 5Y) at or above the category median'],
  ['Q_W_SHARPE', 'Sharpe percentile', 0.20, 'pct', 'Stage 3: risk-adjusted return, 3Y, risk-free 6.5%'],
  ['Q_W_DD', 'Max drawdown percentile', 0.20, 'pct', 'Stage 3: smaller 3Y drawdown ranks higher'],
  ['Watchlist'],
  ['NEAR_BAND', 'Near-miss band below each threshold', 0.02, 'pct', 'A fund within this many points of every threshold is a Near miss.'],
  ['WATCH_AUM', 'Watchlist minimum AUM (Rs crore)', 5000, 'int', ''],
  ['RECENT_1Y', 'Recent outperformer: minimum 1Y return', 0.30, 'pct', 'Strong recent returns without the long-term record.'],
  ['MOM_3M', 'Momentum: minimum 3M return', 0.08, 'pct', ''],
  ['MOM_6M', 'Momentum: minimum 6M return', 0.12, 'pct', ''],
  ['Persistence (evaluated on monthly observations; the last weekly run in a month is that month\'s observation)'],
  ['CONFIRM_MONTHS', 'Months qualified before NEW becomes QUALIFIED', 3, 'int', ''],
  ['REVIEW_MONTHS', 'Months failing before WATCH becomes REVIEW', 2, 'int', 'One failing month = WATCH (a warning, not a sell signal).'],
  ['REMOVE_MONTHS', 'Months failing before REMOVED', 3, 'int', ''],
  ['AUM_CHANGE_ALERT', 'Flag AUM change between runs larger than', 0.15, 'pct', ''],
  ['Allocation layer (separate from fund quality)'],
  ['LUMP_TRIGGER', 'Opportunistic lump sum when NAV is this far below 52W high', 0.10, 'pct', 'Qualified funds below this get "SIP + Opportunistic Lump Sum"; others "SIP".'],
  ['Data feed'],
  ['DATA_URL', 'Data feed URL', DEFAULT_DATA_URL, 'text', 'Folder holding funds.csv and manifest.json, written by the weekly pipeline.']
];

// ---------------------------------------------------------------- style
const NAVY = '#1F3864', BURGUNDY = '#7B2C3B', CREAM = '#F7F3EA', GRID = '#C9C2B2', FONT = 'Arial';
const FILL = {
  strong: '#9FD39B', pass: '#DDEFD6', near: '#FFF0C2', fail: '#F6C9C0',
  full: '#9FD39B', emerging: '#DDEFD6', recent: '#D6E4F5', nearmiss: '#FFF0C2', momentum: '#E6DAF0',
  NEW: '#D6E4F5', QUALIFIED: '#9FD39B', WATCH: '#FFF0C2', REVIEW: '#F9D3A5', REMOVED: '#F6C9C0'
};

// ================================================================ menu & triggers
function onOpen() {
  try { SpreadsheetApp.getUi(); } catch (e) { return; }
  SpreadsheetApp.getUi().createMenu('MF Screener')
    .addItem('Refresh now (import latest data)', 'refreshNow')
    .addItem('Check for new data', 'checkForNewData')
    .addSeparator()
    .addItem('Set up / repair workbook', 'setupScreener')
    .addItem('Re-seed history from backfill', 'reseedHistory')
    .addToUi();
}

function installTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['checkForNewData'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkForNewData').timeBased().everyHours(1).create();
}

function checkForNewData() {
  const props = PropertiesService.getDocumentProperties();
  const m = fetchJson_(dataUrl_() + 'manifest.json');
  if (m.run_id && m.run_id !== props.getProperty('last_run_id')) refresh_(m);
}

function refreshNow() {
  refresh_(fetchJson_(dataUrl_() + 'manifest.json'));
}

// ================================================================ setup
function setupScreener() {
  const ss = ss_();
  ss.setSpreadsheetTimeZone('Asia/Kolkata');
  TAB_ORDER.forEach(k => { if (!ss.getSheetByName(TAB[k])) ss.insertSheet(TAB[k]); });
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > TAB_ORDER.length) ss.deleteSheet(blank);
  TAB_ORDER.forEach((k, i) => { ss.setActiveSheet(ss.getSheetByName(TAB[k])); ss.moveActiveSheet(i + 1); });

  buildSettings_(ss);
  buildAllFunds_(ss);
  buildQueryTabs_(ss);
  buildPortfolio_(ss);
  buildLogTabs_(ss);
  buildDashboard_(ss);
  ss.getSheetByName(TAB.state).hideSheet();
  installTriggers_();
  ss.setActiveSheet(ss.getSheetByName(TAB.dash));
  SpreadsheetApp.flush();
  try { refreshNow(); } catch (e) { ss.toast('Workbook built. Data import failed: ' + e.message, 'MF Screener', 15); }
}

function styleHeader_(range) {
  range.setFontWeight('bold').setFontColor('#FFFFFF').setBackground(NAVY).setFontFamily(FONT)
    .setVerticalAlignment('middle').setWrap(true);
}

function title_(sh, text, sub) {
  sh.getRange('A1').setValue(text).setFontFamily(FONT).setFontSize(15).setFontWeight('bold').setFontColor(BURGUNDY);
  if (sub) sh.getRange('A2').setValue(sub).setFontFamily(FONT).setFontColor('#555555').setFontStyle('italic');
}

function fmtFor_(t, textAsGeneral) {
  if (textAsGeneral && (t === 'str' || t === 'text')) return 'General';
  return { int: '#,##0', num: '#,##0.0000', num2: '0.00', num1: '0.0', num0: '0', pct: '0.0%', pct2: '0.00%',
    date: 'dd-mmm-yyyy', str: '@', text: '@' }[t] || 'General';
}

function buildSettings_(ss) {
  const sh = ss.getSheetByName(TAB.set);
  const existing = {};
  ss.getNamedRanges().forEach(n => { existing[n.getName()] = n.getRange().getValue(); });
  sh.clear();
  title_(sh, 'Settings', 'Blue cells are inputs. Every screen, score and status reads from here; change a value and the screen recalculates.');
  sh.getRange('A4:C4').setValues([['Parameter', 'Value', 'Notes']]);
  styleHeader_(sh.getRange('A4:C4'));
  let row = 5;
  SETTINGS.forEach(s => {
    if (s.length === 1) {
      sh.getRange(row, 1, 1, 3).merge().setValue(s[0]).setFontWeight('bold').setBackground(CREAM).setFontColor(BURGUNDY);
      row++;
      return;
    }
    const [name, label, def, fmt, note] = s;
    const val = (existing[name] !== undefined && existing[name] !== '') ? existing[name] : def;
    sh.getRange(row, 1, 1, 3).setValues([[label, val, note]]);
    const cell = sh.getRange(row, 2);
    cell.setNumberFormat(fmtFor_(fmt)).setFontColor('#0000FF').setBackground('#FFFDE7');
    const nr = ss.getNamedRanges().filter(x => x.getName() === name)[0];
    if (nr) nr.setRange(cell); else ss.setNamedRange(name, cell);
    row++;
  });
  row++;
  sh.getRange(row, 1, 1, 3).merge().setValue('Review calendar').setFontWeight('bold').setBackground(CREAM).setFontColor(BURGUNDY);
  const cal = [['Weekly (Saturday)', 'Data refresh', 'Pipeline runs Saturday morning US Eastern, after Friday NAVs are published in India.'],
    ['Monthly', 'Status observation', 'The last weekly run of each month is that month\'s observation for the persistence rules.'],
    ['Quarterly', 'Formal review of qualifiers', 'January, April, July, October. Quarterly average AUM from AMFI arrives in the first weeks after quarter end.'],
    ['Semiannually', 'Allocation and rebalancing review', 'April and October.'],
    ['Annually', 'Methodology review', 'Thresholds, weights and categories on this tab.']];
  sh.getRange(row + 1, 1, cal.length, 3).setValues(cal);
  sh.setColumnWidth(1, 330); sh.setColumnWidth(2, 340); sh.setColumnWidth(3, 620);
  sh.getRange('A:C').setFontFamily(FONT).setVerticalAlignment('middle');
  sh.getRange('C:C').setWrap(true);
  sh.setFrozenRows(4);
}

function setting_(ss, name) { const r = ss.getRangeByName(name); return r ? r.getValue() : null; }
function settings_(ss) {
  const s = {};
  SETTINGS.forEach(x => { if (x.length > 1) s[x[0]] = setting_(ss, x[0]); });
  return s;
}
function dataUrl_() {
  let u = setting_(ss_(), 'DATA_URL') || DEFAULT_DATA_URL;
  return u.slice(-1) === '/' ? u : u + '/';
}

// ---------------------------------------------------------------- All Funds formulas
function screenFormulas_() {
  const v = k => R(k);
  const periods = [['r1', 'MIN_1Y', 'W_1Y'], ['r2', 'MIN_2Y', 'W_2Y'], ['r3', 'MIN_3Y', 'W_3Y'],
    ['r5', 'MIN_5Y', 'W_5Y'], ['r10', 'MIN_10Y', 'W_10Y'], ['rsi', 'MIN_SI', 'W_SI']];
  const A = v('code');
  const wrap = (hdr, expr) => '={"' + hdr + '";ARRAYFORMULA(IF(' + A + '="",,' + expr + '))}';
  const appl = periods.map(p => '(' + v(p[0]) + '<>"")').join('+');
  // a blank threshold on Settings switches that check off
  const fail = periods.map(p => '(' + v(p[0]) + '<>"")*(' + p[1] + '<>"")*(' + v(p[0]) + '<' + p[1] + ')').join('+');
  const failb = periods.map(p => '(' + v(p[0]) + '<>"")*(' + p[1] + '<>"")*(' + v(p[0]) + '<' + p[1] + '-NEAR_BAND)').join('+');
  const num = periods.map(p => 'IF(' + v(p[0]) + '="",0,' + v(p[0]) + '*' + p[2] + ')').join('+');
  const den = periods.map(p => 'IF(' + v(p[0]) + '="",0,' + p[2] + ')').join('+');
  const qparts = [['pperf', 'Q_W_PERF'], ['pcons', 'Q_W_CONS'], ['psharpe', 'Q_W_SHARPE'], ['pmdd', 'Q_W_DD']];
  const qnum = qparts.map(p => 'IF(' + v(p[0]) + '="",0,' + v(p[0]) + '*' + p[1] + ')').join('+');
  const qden = qparts.map(p => 'IF(' + v(p[0]) + '="",0,' + p[1] + ')').join('+');
  const catList = '","&SUBSTITUTE(SUBSTITUTE(SCREEN_CATEGORIES,", ",",")," ,",",")&","';
  const elig = 'IF(TRIM(SCREEN_CATEGORIES)="All",' + A + '<>"",ISNUMBER(SEARCH(","&' + v('broad') + '&",",' + catList +
    '))+ISNUMBER(SEARCH(","&' + v('cat') + '&",",' + catList + '))>0)';
  const E = v('elig'), AOK = v('aumok'), F = v('fail'), FB = v('failb');
  const result = 'IF((' + E + ')*(' + AOK + ')*(' + F + '=0)*(' + v('r10') + '<>""),"Full-history qualifier",' +
    'IF((' + E + ')*(' + AOK + ')*(' + F + '=0)*(' + v('age') + '>=MIN_HIST_YEARS)*(' + v('r3') + '<>""),"Emerging qualifier",' +
    'IF((' + E + ')*(' + v('aum') + '>=WATCH_AUM)*(' + v('r1') + '<>"")*(' + v('r1') + '>=RECENT_1Y)*(((' + v('r2') + '="")+(' + v('r2') + '>=MIN_2Y))>0),"Recent outperformer",' +
    'IF((' + E + ')*(' + v('aum') + '>=WATCH_AUM)*(' + FB + '=0)*(' + v('r3') + '<>""),"Near miss",' +
    'IF((' + E + ')*(' + v('aum') + '>=WATCH_AUM)*(' + v('r3m') + '<>"")*(' + v('r3m') + '>=MOM_3M)*(' + v('r6m') + '>=MOM_6M),"Momentum","Not qualified")))))';
  const P = v('perf'), C = v('cat');
  const pperf = 'IF(' + P + '="",,IFERROR(100*COUNTIFS(' + C + ',' + C + ',' + P + ',"<"&' + P + ')/(COUNTIFS(' + C + ',' + C + ',' + P + ',">-100")-1),))';
  const Q = v('quality'), RS = v('result');
  const rank = 'IF(REGEXMATCH(' + RS + ',"qualifier$"),COUNTIFS(' + RS + ',"*qualifier",' + Q + ',">"&' + Q + ')+1,)';
  const signal = 'IF(REGEXMATCH(' + RS + ',"qualifier$"),IF(' + v('offhi') + '<=-LUMP_TRIGGER,"QUALIFIED - SIP + Opportunistic Lump Sum","QUALIFIED - SIP"),' +
    'IF(REGEXMATCH(' + v('status') + '&"","^(WATCH|REVIEW|REMOVED)$"),"HOLD / REVIEW",))';
  return {
    appl: wrap('Checks Applicable', appl),
    fail: wrap('Checks Failed', fail),
    failb: wrap('Checks Failed (near-miss band)', failb),
    aumok: wrap('AUM OK', v('aum') + '>=MIN_AUM'),
    elig: wrap('In Screened Categories', elig),
    result: wrap('Screen Result', result),
    perf: wrap('Performance Score', 'IF((' + den + ')=0,,(' + num + ')/(' + den + '))'),
    pperf: wrap('Performance Pctile', pperf),
    quality: wrap('Quality Score', 'IF((' + qden + ')=0,,(' + qnum + ')/(' + qden + '))'),
    rank: wrap('Rank', rank),
    signal: wrap('Allocation Signal', signal)
  };
}

function buildAllFunds_(ss) {
  const sh = ss.getSheetByName(TAB.all);
  ensureCols_(sh, N_ALL);
  ensureRows_(sh, 2000);
  sh.setConditionalFormatRules([]);
  const headers = FUND_COLS.map(c => c[1]).concat(STATUS_COLS.map(c => c[1]));
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  const f = screenFormulas_();
  FORMULA_COLS.forEach(c => sh.getRange(1, COL[c[0]]).setFormula(f[c[0]]));
  styleHeader_(sh.getRange(1, 1, 1, N_ALL));
  sh.getRange(1, N_DATA + 1, 1, N_STATUS).setBackground(BURGUNDY);
  sh.getRange(1, N_DATA + N_STATUS + 1, 1, FORMULA_COLS.length).setBackground('#2E5E4E');
  sh.setRowHeight(1, 42);
  sh.setFrozenRows(1); sh.setFrozenColumns(2);
  FUND_COLS.concat(STATUS_COLS.map(c => [c[0], c[1], null, c[2]])).forEach(c => {
    sh.getRange(2, COL[c[0]], sh.getMaxRows() - 1, 1).setNumberFormat(fmtFor_(c[3]));
  });
  FORMULA_COLS.forEach(c => sh.getRange(2, COL[c[0]], sh.getMaxRows() - 1, 1).setNumberFormat(fmtFor_(c[2], true)));
  sh.getRange(1, 1, sh.getMaxRows(), N_ALL).setFontFamily(FONT).setFontSize(9);
  sh.setColumnWidths(1, N_ALL, 92);
  sh.setColumnWidth(COL.name, 300); sh.setColumnWidth(COL.amc, 170); sh.setColumnWidth(COL.mgr, 200);
  sh.setColumnWidth(COL.sibasis, 200); sh.setColumnWidth(COL.flags, 160); sh.setColumnWidth(COL.result, 150);
  sh.setColumnWidth(COL.signal, 220); sh.setColumnWidth(COL.cat, 140);
  applyReturnColours_(sh, 2, { r1: COL.r1, r2: COL.r2, r3: COL.r3, r5: COL.r5, r10: COL.r10, rsi: COL.rsi });
  applyTextColours_(sh, [sh.getRange(2, COL.result, sh.getMaxRows() - 1, 1), sh.getRange(2, COL.status, sh.getMaxRows() - 1, 1)]);
  const filter = sh.getFilter();
  if (filter) filter.remove();
}

/** Return columns: strong green >= threshold+5pts, green >= threshold, amber within band, red below. */
function applyReturnColours_(sh, firstRow, cols) {
  const th = { r1: 'MIN_1Y', r2: 'MIN_2Y', r3: 'MIN_3Y', r5: 'MIN_5Y', r10: 'MIN_10Y', rsi: 'MIN_SI' };
  const rules = sh.getConditionalFormatRules().filter(r => !(r.getBooleanCondition() &&
    String(r.getBooleanCondition().getCriteriaValues()[0]).indexOf('INDIRECT("Settings!') >= 0));
  Object.keys(cols).forEach(k => {
    const a1 = colLetter(cols[k]) + firstRow;
    const rng = sh.getRange(firstRow, cols[k], sh.getMaxRows() - firstRow + 1, 1);
    const ref = 'INDIRECT("Settings!' + ss_settingCell_(th[k]) + '")';
    const band = 'INDIRECT("Settings!' + ss_settingCell_('NEAR_BAND') + '")';
    const isNum = 'ISNUMBER(' + a1 + ')';
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(' + isNum + ',' + a1 + '>=' + ref + '+0.05)').setBackground(FILL.strong).setRanges([rng]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(' + isNum + ',' + a1 + '>=' + ref + ')').setBackground(FILL.pass).setRanges([rng]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(' + isNum + ',' + a1 + '>=' + ref + '-' + band + ')').setBackground(FILL.near).setRanges([rng]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(' + isNum + ',' + a1 + '<' + ref + '-' + band + ')').setBackground(FILL.fail).setRanges([rng]).build());
  });
  sh.setConditionalFormatRules(rules);
}

function ss_settingCell_(name) {
  const r = ss_().getRangeByName(name);
  return '$B$' + r.getRow();
}

function applyTextColours_(sh, ranges) {
  const map = [['Full-history qualifier', FILL.full], ['Emerging qualifier', FILL.emerging], ['Recent outperformer', FILL.recent],
    ['Near miss', FILL.nearmiss], ['Momentum', FILL.momentum], ['NEW', FILL.NEW], ['QUALIFIED', FILL.QUALIFIED],
    ['WATCH', FILL.WATCH], ['REVIEW', FILL.REVIEW], ['REMOVED', FILL.REMOVED],
    ['QUALIFIED - SIP', FILL.pass], ['QUALIFIED - SIP + Opportunistic Lump Sum', FILL.strong], ['HOLD / REVIEW', FILL.REVIEW]];
  const rules = sh.getConditionalFormatRules();
  map.forEach(m => rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(m[0]).setBackground(m[1]).setRanges(ranges).build()));
  sh.setConditionalFormatRules(rules);
}

// ---------------------------------------------------------------- QUERY tabs
function q_(keys) { return keys.map(k => L(k)).join(', '); }
const ALLREF = "'All Funds'!A1:";

function buildQueryTabs_(ss) {
  const end = colLetter(N_ALL);
  const rs = L('result'), st = L('status');

  const qs = ss.getSheetByName(TAB.qual);
  qs.clear();
  title_(qs, 'Qualified Funds', 'Funds passing every applicable Stage 1 threshold and the AUM minimum, ordered by quality score. Live: edits on Settings re-screen immediately.');
  const qCols = ['rank', 'name', 'cat', 'result', 'status', 'streak', 'quality', 'perf', 'r1', 'r2', 'r3', 'r5', 'r10', 'rsi', 'aum', 'ter', 'sharpe', 'mdd', 'cons', 'signal', 'code'];
  qs.getRange('A4').setFormula('=IFERROR(QUERY(' + ALLREF + end + ',"select ' + q_(qCols) + ' where ' + rs + " ends with 'qualifier' order by " + L('quality') + ' desc",1),"No fund meets every threshold on the Settings tab today.")');
  formatQueryTab_(qs, qCols, 4);
  applyReturnColours_(qs, 5, { r1: 9, r2: 10, r3: 11, r5: 12, r10: 13, rsi: 14 });

  const rk = ss.getSheetByName(TAB.rank);
  rk.clear();
  title_(rk, 'Rankings', 'Left: qualifiers ranked by quality score. Right: top 50 by performance score among funds meeting the AUM minimum with 5+ years of history, whatever the return threshold. Quality score is a percentile within each fund\'s own category, so compare it across funds in the same category.');
  const rCols = ['rank', 'name', 'cat', 'result', 'quality', 'perf', 'pperf', 'pcons', 'psharpe', 'pmdd', 'status', 'streak', 'signal'];
  rk.getRange('A4').setFormula('=IFERROR(QUERY(' + ALLREF + end + ',"select ' + q_(rCols) + ' where ' + rs + " ends with 'qualifier' order by " + L('rank') + ' asc",1),"No qualifiers today.")');
  const lCols = ['name', 'cat', 'perf', 'r5', 'r10', 'rsi', 'sharpe', 'mdd', 'cons', 'quality', 'result'];
  const lStart = rCols.length + 2;
  rk.getRange(4, lStart).setFormula('=IFERROR(QUERY(' + ALLREF + end + ',"select ' + q_(lCols) + ' where ' + L('aumok') + ' = true and ' + L('elig') + ' = true and ' + L('r5') + ' is not null order by ' + L('perf') + ' desc limit 50",1),"")');
  formatQueryTab_(rk, rCols, 4, 1);
  formatQueryTab_(rk, lCols, 4, lStart);

  const wl = ss.getSheetByName(TAB.watch);
  wl.clear();
  title_(wl, 'Watchlist', 'Near misses, recent outperformers, momentum funds and former qualifiers on WATCH or REVIEW. For monitoring, not buying.');
  const wCols = ['name', 'cat', 'result', 'status', 'aum', 'r3m', 'r6m', 'r1', 'r2', 'r3', 'r5', 'r10', 'rsi', 'fail', 'perf', 'quality', 'offhi', 'code'];
  wl.getRange('A4').setFormula('=IFERROR(QUERY(' + ALLREF + end + ',"select ' + q_(wCols) + ' where ' + rs + " = 'Near miss' or " + rs + " = 'Recent outperformer' or " + rs + " = 'Momentum' or " + st + " = 'WATCH' or " + st + " = 'REVIEW' order by " + rs + ', ' + L('perf') + ' desc",1),"Nothing on the watchlist today.")');
  formatQueryTab_(wl, wCols, 4);
  applyReturnColours_(wl, 5, { r1: 8, r2: 9, r3: 10, r5: 11, r10: 12, rsi: 13 });
}

function typeOf_(key) {
  const all = FUND_COLS.map(c => [c[0], c[3]]).concat(STATUS_COLS.map(c => [c[0], c[2]])).concat(FORMULA_COLS.map(c => [c[0], c[2]]));
  const hit = all.filter(x => x[0] === key)[0];
  return hit ? hit[1] : 'str';
}

function formatQueryTab_(sh, keys, headerRow, startCol) {
  startCol = startCol || 1;
  styleHeader_(sh.getRange(headerRow, startCol, 1, keys.length));
  sh.setRowHeight(headerRow, 40);
  keys.forEach((k, i) => {
    sh.getRange(headerRow + 1, startCol + i, Math.max(sh.getMaxRows() - headerRow, 1), 1).setNumberFormat(fmtFor_(typeOf_(k), true));
    sh.setColumnWidth(startCol + i, k === 'name' ? 290 : (k === 'signal' ? 230 : (k === 'cat' || k === 'result' ? 140 : 84)));
  });
  sh.getRange(1, startCol, sh.getMaxRows(), keys.length).setFontFamily(FONT).setFontSize(9);
  sh.getRange('A1').setFontSize(15);
  sh.setFrozenRows(headerRow);
  applyTextColours_(sh, [sh.getRange(headerRow + 1, startCol, Math.max(sh.getMaxRows() - headerRow, 1), keys.length)]);
}

// ---------------------------------------------------------------- Portfolio
function buildPortfolio_(ss) {
  const sh = ss.getSheetByName(TAB.port);
  const keep = sh.getLastRow() >= 5 ? sh.getRange(5, 1, sh.getLastRow() - 4, 5).getValues() : [];
  sh.clear();
  title_(sh, 'Portfolio', 'Enter holdings in the blue columns (Scheme Code from All Funds). Everything else fills in. Reviewed quarterly; rebalanced semiannually.');
  const inputs = ['Scheme Code', 'Units', 'Avg Cost NAV', 'Target Weight', 'Notes'];
  sh.getRange(4, 1, 1, 5).setValues([inputs]);
  const lk = (key, hdr) => '={"' + hdr + '";ARRAYFORMULA(IF(A5:A="",,IFERROR(VLOOKUP(A5:A,\'All Funds\'!$A$2:$' + colLetter(N_ALL) + ',' + COL[key] + ',FALSE),)))}';
  const f = [
    '={"Fund";ARRAYFORMULA(IF(A5:A="",,IFERROR(VLOOKUP(A5:A,\'All Funds\'!$A$2:$B,2,FALSE),"Not found")))}',
    lk('cat', 'Category'), lk('nav', 'NAV'),
    '={"Value (Rs)";ARRAYFORMULA(IF(A5:A="",,IFERROR(B5:B*H5:H,)))}',
    '={"Gain";ARRAYFORMULA(IF((A5:A="")+(C5:C=""),,IFERROR(H5:H/C5:C-1,)))}',
    '={"Weight";ARRAYFORMULA(IF(A5:A="",,IFERROR(I5:I/SUM(I5:I),)))}',
    lk('result', 'Screen Result'), lk('status', 'Status'), lk('streak', 'Months Qualified (consecutive)'),
    lk('quality', 'Quality Score'), lk('signal', 'Allocation Signal'),
    '={"Alert";ARRAYFORMULA(IF(A5:A="",,IF(F5:F="Not found","Not in universe: check merger or closure",IF(M5:M="REMOVED","Review: removed from screen",IF(M5:M="REVIEW","Review: failing 2+ months",IF(M5:M="WATCH","Watch: failed this month",IF(REGEXMATCH(L5:L&"","qualifier$"),"OK","Not a qualifier")))))))}',
    '={"Weight vs Target";ARRAYFORMULA(IF((A5:A="")+(D5:D=""),,K5:K-D5:D))}'
  ];
  f.forEach((x, i) => sh.getRange(4, 6 + i).setFormula(x));
  styleHeader_(sh.getRange(4, 1, 1, 5 + f.length));
  sh.getRange(4, 1, 1, 5).setBackground('#2F5597');
  sh.getRange('A5:E').setFontColor('#0000FF').setBackground('#FFFDE7');
  sh.getRange('A5:A').setNumberFormat('0');
  sh.getRange('B5:B').setNumberFormat('#,##0.000'); sh.getRange('C5:C').setNumberFormat('#,##0.0000');
  sh.getRange('D5:D').setNumberFormat('0.0%'); sh.getRange('H5:H').setNumberFormat('#,##0.0000');
  sh.getRange('I5:I').setNumberFormat('#,##0'); sh.getRange('J5:K').setNumberFormat('0.0%');
  sh.getRange('N5:N').setNumberFormat('0'); sh.getRange('O5:O').setNumberFormat('0.0'); sh.getRange('R5:R').setNumberFormat('0.0%');
  if (keep.length) sh.getRange(5, 1, keep.length, 5).setValues(keep);
  sh.setColumnWidth(6, 290); sh.setColumnWidth(16, 230); sh.setColumnWidth(17, 220); sh.setColumnWidth(5, 160);
  sh.getRange(1, 1, sh.getMaxRows(), 18).setFontFamily(FONT).setFontSize(9);
  sh.getRange('A1').setFontSize(15);
  sh.setFrozenRows(4);
  applyTextColours_(sh, [sh.getRange('L5:P')]);
  const rules = sh.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('OK').setBackground(FILL.pass).setRanges([sh.getRange('Q5:Q')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(Q5<>"",Q5<>"OK")').setBackground(FILL.fail).setRanges([sh.getRange('Q5:Q')]).build());
  sh.setConditionalFormatRules(rules);
}

// ---------------------------------------------------------------- log tabs
const CHG_HDR = ['Run Date', 'Scheme Code', 'Fund', 'Category', 'Change', 'From', 'To', 'Detail'];
const HIST_HDR = ['Month', 'Run Date', 'Scheme Code', 'Fund', 'Category', 'Screen Result', 'Qualified', 'Status',
  'Quality Score', 'Rank', '1Y', '3Y', '5Y', '10Y', 'Since Inception', 'AUM (Rs Cr)', 'Reason Not Qualified', 'Source'];
const ARCH_HDR = ['Fund', 'Scheme Code', 'Category', 'Date Entered (latest streak)', 'First Qualified', 'Date Removed',
  'Months Qualified', 'Best Rank', 'Latest Rank', 'Current Status', 'Previous Status', 'Reason for Removal',
  '1Y', '3Y', '5Y', '10Y', 'AUM at Entry (Rs Cr)', 'Latest AUM (Rs Cr)', 'Last Observed'];
const STATE_HDR = ['Scheme Code', 'Fund', 'Category', 'AUM', 'Screen Result', 'Status', 'Run Date'];
const RUNLOG_HDR = ['Refreshed', 'Pipeline Run', 'NAV Date', 'Funds', 'Full-history', 'Emerging', 'Recent', 'Near miss',
  'Momentum', 'WATCH', 'REVIEW', 'Changes Logged', 'Validation Flags'];

function buildLogTabs_(ss) {
  const spec = [[TAB.chg, CHG_HDR, 'Changes', 'Every run logs entries, exits, status changes, large AUM moves, and funds added to or dropped from the universe. Newest first.'],
    [TAB.hist, HIST_HDR, 'Historical Data', 'Monthly snapshots of every fund that qualified, was on the watchlist, or was held. The latest weekly run replaces its month\'s rows.'],
    [TAB.arch, ARCH_HDR, 'Archive', 'Every fund that has ever qualified, rebuilt from Historical Data on each run. Nothing is deleted.']];
  spec.forEach(s => {
    const sh = ss.getSheetByName(s[0]);
    title_(sh, s[2], s[3]);
    sh.getRange(4, 1, 1, s[1].length).setValues([s[1]]);
    styleHeader_(sh.getRange(4, 1, 1, s[1].length));
    sh.setFrozenRows(4);
    sh.getRange(1, 1, sh.getMaxRows(), s[1].length).setFontFamily(FONT).setFontSize(9);
    sh.getRange('A1').setFontSize(15);
  });
  const h = ss.getSheetByName(TAB.hist);
  h.getRange('A5:A').setNumberFormat('@'); h.getRange('B5:B').setNumberFormat('dd-mmm-yyyy');
  h.getRange('I5:I').setNumberFormat('0.0'); h.getRange('K5:O').setNumberFormat('0.0%'); h.getRange('P5:P').setNumberFormat('#,##0');
  h.setColumnWidth(4, 280); h.setColumnWidth(17, 220);
  const c = ss.getSheetByName(TAB.chg);
  c.getRange('A5:A').setNumberFormat('dd-mmm-yyyy'); c.setColumnWidth(3, 280); c.setColumnWidth(5, 170); c.setColumnWidth(8, 320);
  const a = ss.getSheetByName(TAB.arch);
  a.getRange('D5:F').setNumberFormat('dd-mmm-yyyy'); a.getRange('M5:P').setNumberFormat('0.0%');
  a.getRange('Q5:R').setNumberFormat('#,##0'); a.getRange('S5:S').setNumberFormat('dd-mmm-yyyy');
  a.setColumnWidth(1, 280); a.setColumnWidth(12, 240);
  applyTextColours_(a, [a.getRange('J5:K')]);
  applyTextColours_(h, [h.getRange('F5:H')]);
  const st = ss.getSheetByName(TAB.state);
  st.getRange(1, 1, 1, STATE_HDR.length).setValues([STATE_HDR]);
  const src = ss.getSheetByName(TAB.src);
  title_(src, 'Data Sources', 'Where each figure comes from, when it was pulled, and validation results for every run.');
}

// ---------------------------------------------------------------- Dashboard
const DASH = { kpiRow: 4, sections: 7, gap: 14, monthlyCol: 16 };
const K = {
  qual: ['rank', 'name', 'cat', 'result', 'status', 'streak', 'quality', 'r1', 'r3', 'r5', 'r10', 'aum'],
  best: ['name', 'cat', 'perf', 'r1', 'r3', 'r5', 'r10', 'rsi', 'sharpe', 'mdd', 'quality', 'result'],
  close: ['name', 'cat', 'fail', 'appl', 'perf', 'r1', 'r2', 'r3', 'r5', 'r10', 'rsi', 'aum'],
  streak: ['name', 'cat', 'streak', 'months', 'status', 'result', 'quality', 'r5', 'r10', 'aum'],
  fresh: ['name', 'cat', 'result', 'quality', 'r1', 'r3', 'r5', 'r10', 'aum'],
  recent: ['name', 'cat', 'r1', 'r2', 'r3', 'r5', 'aum', 'quality']
};

function buildDashboard_(ss) {
  const sh = ss.getSheetByName(TAB.dash);
  sh.clear();
  sh.getCharts().forEach(c => sh.removeChart(c));
  title_(sh, 'Indian Mutual Fund Screener', 'Waiting for first data import.');
  const end = colLetter(N_ALL);
  const AF = "'All Funds'!";
  const rs = AF + R('result'), st = AF + R('status'), sk = AF + R('streak');
  const kpis = [
    ['Funds screened', '=COUNTA(' + AF + R('code') + ')'],
    ['Full-history qualifiers', '=COUNTIF(' + rs + ',"Full-history qualifier")'],
    ['Emerging qualifiers', '=COUNTIF(' + rs + ',"Emerging qualifier")'],
    ['Watchlist', '=COUNTIF(' + rs + ',"Near miss")+COUNTIF(' + rs + ',"Recent outperformer")+COUNTIF(' + rs + ',"Momentum")'],
    ['New this month', '=COUNTIFS(' + st + ',"NEW",' + sk + ',1)'],
    ['At risk (WATCH / REVIEW)', '=COUNTIF(' + st + ',"WATCH")+COUNTIF(' + st + ',"REVIEW")'],
    ['Portfolio alerts', '=COUNTIFS(Portfolio!Q5:Q,"<>OK",Portfolio!A5:A,"<>")']
  ];
  kpis.forEach((k, i) => {
    const c = 1 + i * 2;
    sh.getRange(DASH.kpiRow, c, 1, 2).merge().setValue(k[0]).setFontColor('#FFFFFF').setBackground(NAVY).setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
    sh.getRange(DASH.kpiRow + 1, c, 1, 2).merge().setFormula(k[1]).setFontSize(20).setFontWeight('bold').setFontColor(BURGUNDY)
      .setBackground(CREAM).setHorizontalAlignment('center').setNumberFormat('#,##0');
  });
  sh.setRowHeight(DASH.kpiRow + 1, 44);
  const sel = keys => q_(keys);
  const Q = (where, keys, order, limit) => '=IFERROR(QUERY(' + ALLREF + end + ',"select ' + sel(keys) + ' where ' + where +
    (order ? ' order by ' + order : '') + ' limit ' + limit + '",1),"None.")';
  const rsc = L('result'), stc = L('status');
  const base = L('aumok') + ' = true and ' + L('elig') + ' = true';
  const dropped = '=IFERROR(QUERY(' + ALLREF + end + ',"select ' + sel(['name', 'cat', 'status', 'since', 'months', 'result', 'r1', 'r3', 'r5', 'aum']) +
    ' where ' + stc + " = 'WATCH' or " + stc + " = 'REVIEW' or (" + stc + " = 'REMOVED' and " + L('since') +
    " >= date '\"&TEXT(TODAY()-92,\"yyyy-mm-dd\")&\"') order by " + L('since') + ' desc limit 10",1),"None.")';
  // [title, formula, column keys (for number formats)]
  const sections = [
    ['Current qualifiers: which funds meet every requirement today?',
      Q(rsc + " ends with 'qualifier'", K.qual, L('quality') + ' desc', 10), K.qual],
    ['Best long-term performers: strongest persistent returns among funds meeting the AUM minimum (5+ years, any threshold)',
      Q(base + ' and ' + L('r5') + ' is not null', K.best, L('perf') + ' desc', 10), K.best],
    ['Closest to qualifying: fewest failed checks against the Settings thresholds',
      Q(base + ' and ' + L('r3') + ' is not null and not (' + rsc + " ends with 'qualifier')", K.close, L('fail') + ' asc, ' + L('perf') + ' desc', 10), K.close],
    ['Consecutive months qualified: who has stayed in the screen longest?',
      Q(L('streak') + ' > 0', K.streak, L('streak') + ' desc, ' + L('quality') + ' desc', 10), K.streak],
    ['New entrants this month', Q(stc + " = 'NEW' and " + L('streak') + ' = 1', K.fresh, L('quality') + ' desc', 10), K.fresh],
    ['Dropped or at risk: former qualifiers on WATCH or REVIEW, or removed in the last 3 months', dropped,
      ['name', 'cat', 'status', 'since', 'months', 'result', 'r1', 'r3', 'r5', 'aum']],
    ['Recent outperformers: strong 1-3 year results without the long record', Q(rsc + " = 'Recent outperformer'", K.recent, L('r1') + ' desc', 10), K.recent],
    ['Portfolio alerts: held funds needing review',
      '=IFERROR(QUERY(Portfolio!A4:Q,"select F, G, L, M, N, P, Q where A is not null and Q <> \'OK\'",1),"None.")',
      ['name', 'cat', 'result', 'status', 'streak', 'signal', 'flags']]
  ];
  sections.forEach((s, i) => {
    const r = DASH.sections + i * DASH.gap;
    sh.getRange(r, 1, 1, 13).merge().setValue(s[0]).setFontWeight('bold').setFontColor(BURGUNDY).setBackground(CREAM).setFontSize(11);
    sh.getRange(r + 1, 1).setFormula(s[1]);
    sh.getRange(r + 1, 1, 1, 13).setFontWeight('bold').setBorder(null, null, true, null, null, null, NAVY, SpreadsheetApp.BorderStyle.SOLID);
  });
  const widths = [190, 120, 120, 110, 90, 90, 90, 80, 80, 80, 80, 80, 90];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setColumnWidth(1, 260);
  ensureRows_(sh, DASH.sections + sections.length * DASH.gap + 5);
  sh.getRange(1, 1, sh.getMaxRows(), 20).setFontFamily(FONT).setFontSize(9);
  sh.getRange('A1').setFontSize(16); sh.getRange('A2').setFontSize(9);
  sh.getRange(DASH.sections, 1, sh.getMaxRows() - DASH.sections + 1, 13).setNumberFormat('General');
  sections.forEach((s, i) => {
    const r0 = DASH.sections + i * DASH.gap + 2;
    s[2].forEach((k, j) => {
      const t = typeOf_(k);
      if (t !== 'str') sh.getRange(r0, j + 1, 11, 1).setNumberFormat(fmtFor_(t));
    });
  });
  applyTextColours_(sh, [sh.getRange(DASH.sections, 1, sections.length * DASH.gap, 13)]);
  // monthly qualifier history (written by the script) + chart
  const mc = DASH.monthlyCol;
  sh.getRange(DASH.sections, mc, 1, 4).merge().setValue('Qualifiers by month').setFontWeight('bold').setFontColor(BURGUNDY).setBackground(CREAM).setFontSize(11);
  sh.getRange(DASH.sections + 1, mc, 1, 4).setValues([['Month', 'Full-history', 'Emerging', 'WATCH / REVIEW']]);
  styleHeader_(sh.getRange(DASH.sections + 1, mc, 1, 4));
  sh.setColumnWidth(14, 24); sh.setColumnWidth(15, 24);
  sh.setColumnWidths(mc, 4, 90);
  sh.setFrozenRows(5);
}

function drawMonthlyChart_(sh, nRows) {
  sh.getCharts().forEach(c => sh.removeChart(c));
  if (nRows < 1) return;
  const mc = DASH.monthlyCol;
  const rng = sh.getRange(DASH.sections + 1, mc, nRows + 1, 4);
  const chart = sh.newChart().asColumnChart().addRange(rng).setStacked()
    .setNumHeaders(1).setTitle('Qualifying funds per monthly observation')
    .setOption('colors', ['#2E7D32', '#9CCC65', '#F9A825']).setOption('legend', { position: 'bottom' })
    .setPosition(DASH.sections, mc + 5, 0, 0).setOption('width', 620).setOption('height', 320).build();
  sh.insertChart(chart);
}

// ================================================================ refresh pipeline
function refresh_(manifest) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return;
  try {
    const ss = ss_();
    if (!ss.getSheetByName(TAB.set) || !ss.getRangeByName('MIN_1Y')) setupScreener();
    const base = dataUrl_();
    const funds = fetchCsv_(base + 'funds.csv');
    importFunds_(ss, funds);
    importWide_(ss, TAB.qnav, fetchCsv_(base + 'quarterly_nav.csv'), 'Quarterly Average NAV',
      'Mean of daily NAVs in each calendar quarter (Jan-Mar, Apr-Jun, Jul-Sep, Oct-Dec). The last column is the current quarter to date.', '#,##0.0000');
    importWide_(ss, TAB.qaum, fetchCsv_(base + 'quarterly_aum.csv'), 'Quarterly Average AUM (Rs crore)',
      'AMFI quarterly average AUM, summed across every plan and option of the fund.', '#,##0');
    SpreadsheetApp.flush();
    if (historyRowCount_(ss) === 0) seedHistory_(ss, fetchCsv_(base + 'backfill.csv'));
    const res = runStatusEngine_(ss, manifest);
    writeSources_(ss, manifest, res);
    PropertiesService.getDocumentProperties().setProperty('last_run_id', manifest.run_id || '');
    ss.toast('Imported ' + funds.length + ' funds; ' + res.changes + ' changes logged.', 'MF Screener', 10);
  } finally {
    lock.releaseLock();
  }
}

function fetchJson_(url) {
  const r = UrlFetchApp.fetch(url + '?t=' + Date.now(), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('HTTP ' + r.getResponseCode() + ' for ' + url);
  return JSON.parse(r.getContentText());
}

/** CSV -> array of objects keyed by header. */
function fetchCsv_(url) {
  const r = UrlFetchApp.fetch(url + '?t=' + Date.now(), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('HTTP ' + r.getResponseCode() + ' for ' + url);
  const rows = Utilities.parseCsv(r.getContentText('UTF-8'));
  const hdr = rows.shift();
  return rows.filter(x => x.length === hdr.length).map(x => { const o = {}; hdr.forEach((h, i) => { o[h] = x[i]; }); o.__hdr = hdr; return o; });
}

function cast_(v, t) {
  if (v === undefined || v === null || v === '' || v === 'nan') return '';
  if (t === 'str') return String(v);
  if (t === 'date') { const p = String(v).split('-'); return p.length === 3 ? new Date(+p[0], +p[1] - 1, +p[2]) : ''; }
  const n = Number(v);
  return isFinite(n) ? n : '';
}

function ensureCols_(sh, n) {
  if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), n - sh.getMaxColumns());
}

function ensureRows_(sh, n) {
  if (sh.getMaxRows() < n) sh.insertRowsAfter(sh.getMaxRows(), n - sh.getMaxRows());
}

function importFunds_(ss, funds) {
  const sh = ss.getSheetByName(TAB.all);
  const values = funds.map(f => FUND_COLS.map(c => cast_(f[c[2]], c[3])));
  ensureRows_(sh, values.length + 1);
  sh.getRange(2, 1, sh.getMaxRows() - 1, N_DATA + N_STATUS).clearContent();
  if (values.length) sh.getRange(2, 1, values.length, N_DATA).setValues(values);
}

function importWide_(ss, tab, rows, title, sub, fmt) {
  const sh = ss.getSheetByName(tab);
  sh.clear();
  if (rows.length) ensureCols_(sh, rows[0].__hdr.length);
  title_(sh, title, sub);
  if (!rows.length) return;
  const hdr = rows[0].__hdr.map(h => ({ code: 'Scheme Code', name: 'Scheme Name', category: 'Category' }[h] || h));
  const vals = rows.map(r => r.__hdr.map((h, i) => i < 3 ? (i === 0 ? Number(r[h]) : r[h]) : cast_(r[h], 'num')));
  ensureRows_(sh, vals.length + 4);
  sh.getRange(4, 1, 1, hdr.length).setNumberFormat('@').setValues([hdr]);
  styleHeader_(sh.getRange(4, 1, 1, hdr.length));
  sh.getRange(5, 1, vals.length, hdr.length).setValues(vals);
  sh.getRange(5, 4, vals.length, hdr.length - 3).setNumberFormat(fmt);
  sh.getRange(5, 1, vals.length, 1).setNumberFormat('0');
  sh.setFrozenRows(4); sh.setFrozenColumns(2);
  sh.setColumnWidth(2, 300); sh.setColumnWidth(3, 140);
  sh.getRange(1, 1, sh.getMaxRows(), hdr.length).setFontFamily(FONT).setFontSize(9);
  sh.getRange('A1').setFontSize(15);
}

// ================================================================ status engine
function isQual_(result) { return /qualifier$/.test(String(result || '')); }

/** Walk a fund's monthly observations (oldest first) and return its status. Pure function. */
function statusFromSeq(seq, S) {
  let streakQ = 0, streakF = 0, ever = false, status = '', months = 0;
  seq.forEach(q => {
    if (q) { streakQ++; streakF = 0; ever = true; months++; status = streakQ < S.CONFIRM_MONTHS ? 'NEW' : 'QUALIFIED'; }
    else if (ever) {
      streakF++; streakQ = 0;
      status = streakF >= S.REMOVE_MONTHS ? 'REMOVED' : (streakF >= S.REVIEW_MONTHS ? 'REVIEW' : 'WATCH');
    } else { streakQ = 0; }
  });
  return { status: status, streak: streakQ, failStreak: streakF, months: months, ever: ever };
}

/** Same rules as the Screen Result formula, for back-filled month-end snapshots. Pure function. */
function qualifyRow(r, S, fund) {
  const per = [['ret_1y', 'MIN_1Y'], ['ret_2y', 'MIN_2Y'], ['ret_3y', 'MIN_3Y'], ['ret_5y', 'MIN_5Y'], ['ret_10y', 'MIN_10Y'], ['ret_si', 'MIN_SI']];
  const reasons = [];
  let fails = 0;
  per.forEach(p => {
    const v = r[p[0]];
    if (S[p[1]] !== '' && v !== '' && v !== null && v !== undefined && v < S[p[1]]) { fails++; reasons.push(p[0].replace('ret_', '').toUpperCase() + ' ' + (v * 100).toFixed(1) + '%'); }
  });
  const aum = r.aum_cr;
  const aumOk = aum !== '' && aum !== null && aum >= S.MIN_AUM;
  if (!aumOk) reasons.push('AUM ' + (aum === '' || aum === null ? 'n/a' : Math.round(aum).toLocaleString('en-IN')) + ' cr');
  const elig = categoryEligible(fund, S.SCREEN_CATEGORIES);
  if (!elig) reasons.push('category not screened');
  let result = 'Not qualified';
  if (elig && aumOk && fails === 0 && r.ret_10y !== '' && r.ret_10y !== null) result = 'Full-history qualifier';
  else if (elig && aumOk && fails === 0 && r.age_years >= S.MIN_HIST_YEARS && r.ret_3y !== '' && r.ret_3y !== null) result = 'Emerging qualifier';
  else if (fails === 0 && aumOk && elig) reasons.push('history under ' + S.MIN_HIST_YEARS + 'Y');
  return { result: result, reason: reasons.join('; ') };
}

function categoryEligible(fund, list) {
  if (!list || String(list).trim() === 'All') return true;
  const items = String(list).split(',').map(x => x.trim().toLowerCase());
  return items.indexOf(String(fund.broad || '').toLowerCase()) >= 0 || items.indexOf(String(fund.category || '').toLowerCase()) >= 0;
}

function monthKey_(d) { return Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM'); }

function obsMonths_() {
  try { return JSON.parse(PropertiesService.getDocumentProperties().getProperty('obs_months') || '[]'); } catch (e) { return []; }
}
function addObsMonths_(list) {
  const set = {};
  obsMonths_().concat(list).forEach(m => { set[m] = true; });
  PropertiesService.getDocumentProperties().setProperty('obs_months', JSON.stringify(Object.keys(set).sort().slice(-120)));
}

function historyRowCount_(ss) { return Math.max(ss.getSheetByName(TAB.hist).getLastRow() - 4, 0); }

function readHistory_(ss) {
  const sh = ss.getSheetByName(TAB.hist);
  const n = historyRowCount_(ss);
  return n ? sh.getRange(5, 1, n, HIST_HDR.length).getValues() : [];
}

/** Seed Historical Data from the pipeline's month-end snapshots so streaks are meaningful from day one. */
function seedHistory_(ss, rows) {
  if (!rows.length) return;
  const S = settings_(ss);
  const sh = ss.getSheetByName(TAB.all);
  const n = sh.getLastRow() - 1;
  const meta = {};
  sh.getRange(2, 1, n, COL.cat).getValues().forEach(r => { meta[r[0]] = { name: r[COL.name - 1], broad: r[COL.broad - 1], category: r[COL.cat - 1] }; });
  const by = {};
  rows.forEach(r => {
    const o = { month: r.month, date: cast_(r.date, 'date'), code: Number(r.code) };
    ['ret_1y', 'ret_2y', 'ret_3y', 'ret_5y', 'ret_10y', 'ret_si', 'aum_cr', 'age_years'].forEach(k => { o[k] = cast_(r[k], 'num'); });
    (by[o.code] = by[o.code] || []).push(o);
  });
  addObsMonths_(rows.map(r => r.month).filter((m, i, a) => m && a.indexOf(m) === i));
  const out = [];
  Object.keys(by).forEach(code => {
    const m = meta[code];
    if (!m) return;
    const seq = by[code].sort((a, b) => a.month < b.month ? -1 : 1);
    const flags = [];
    seq.forEach((o, i) => {
      const q = qualifyRow(o, S, m);
      flags.push(isQual_(q.result));
      const st = statusFromSeq(flags, S);
      const recent = flags.slice(Math.max(0, i - S.REMOVE_MONTHS), i + 1).some(x => x);
      if (!recent) return;
      out.push([o.month, o.date, Number(code), m.name, m.category, q.result, isQual_(q.result) ? 'Yes' : 'No', st.status,
        '', '', o.ret_1y, o.ret_3y, o.ret_5y, o.ret_10y, o.ret_si, o.aum_cr, isQual_(q.result) ? '' : q.reason, 'Backfill (month-end NAV)']);
    });
  });
  out.sort((a, b) => a[0] < b[0] ? -1 : (a[0] > b[0] ? 1 : a[2] - b[2]));
  if (out.length) {
    const h = ss.getSheetByName(TAB.hist);
    ensureRows_(h, out.length + 5);
    h.getRange(5, 1, out.length, 1).setNumberFormat('@');
    h.getRange(5, 1, out.length, HIST_HDR.length).setValues(out);
  }
}

function reseedHistory() {
  const ss = ss_();
  const h = ss.getSheetByName(TAB.hist);
  if (h.getLastRow() > 4) h.getRange(5, 1, h.getLastRow() - 4, HIST_HDR.length).clearContent();
  seedHistory_(ss, fetchCsv_(dataUrl_() + 'backfill.csv'));
  runStatusEngine_(ss, fetchJson_(dataUrl_() + 'manifest.json'));
}

function runStatusEngine_(ss, manifest) {
  const S = settings_(ss);
  const now = new Date();
  const runDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const curMonth = monthKey_(now);
  const sh = ss.getSheetByName(TAB.all);
  const n = sh.getLastRow() - 1;
  const all = sh.getRange(2, 1, n, N_ALL).getValues();
  const c = k => COL[k] - 1;

  // portfolio codes
  const ps = ss.getSheetByName(TAB.port);
  const held = {};
  if (ps.getLastRow() >= 5) ps.getRange(5, 1, ps.getLastRow() - 4, 1).getValues().forEach(r => { if (r[0] !== '') held[Number(r[0])] = true; });

  // previous run state
  const stSh = ss.getSheetByName(TAB.state);
  const prev = {};
  if (stSh.getLastRow() > 1) stSh.getRange(2, 1, stSh.getLastRow() - 1, STATE_HDR.length).getValues()
    .forEach(r => { prev[r[0]] = { name: r[1], cat: r[2], aum: r[3], result: r[4], status: r[5] }; });
  const firstRun = Object.keys(prev).length === 0;

  // history: observation months and per-fund last observation in each month
  const hist = readHistory_(ss).filter(r => String(r[0]) !== curMonth);
  const obsMonths = {};
  const fundMonths = {};
  hist.forEach(r => {
    const m = String(r[0]); obsMonths[m] = true;
    const f = (fundMonths[r[2]] = fundMonths[r[2]] || {});
    if (!f[m] || r[1] >= f[m].date) f[m] = { date: r[1], q: r[6] === 'Yes', status: r[7], rank: r[9] };
  });
  obsMonths_().forEach(m => { if (m !== curMonth) obsMonths[m] = true; });
  const months = Object.keys(obsMonths).sort();

  const statusOut = [], histOut = [], changes = [], stateOut = [];
  const counts = { full: 0, emerging: 0, recent: 0, near: 0, momentum: 0, WATCH: 0, REVIEW: 0 };
  const seen = {};
  all.forEach(r => {
    const code = r[c('code')];
    if (code === '') { statusOut.push(['', '', '', '']); return; }
    seen[code] = true;
    const result = r[c('result')];
    const q = isQual_(result);
    const fm = fundMonths[code] || {};
    const seq = months.map(m => !!(fm[m] && fm[m].q)).concat([q]);
    const st = statusFromSeq(seq, S);
    // status since: walk back while the status was the same
    let since = runDate;
    const statuses = months.map(m => fm[m] ? fm[m].status : '');
    for (let i = statuses.length - 1; i >= 0; i--) {
      if (!fm[months[i]] && st.status === 'REMOVED') continue;   // removed funds stop being tracked
      if (statuses[i] === st.status && fm[months[i]]) since = fm[months[i]].date; else break;
    }
    statusOut.push([st.status, st.ever ? st.streak : '', st.status ? since : '', st.ever ? st.months : '']);
    if (result === 'Full-history qualifier') counts.full++;
    else if (result === 'Emerging qualifier') counts.emerging++;
    else if (result === 'Recent outperformer') counts.recent++;
    else if (result === 'Near miss') counts.near++;
    else if (result === 'Momentum') counts.momentum++;
    if (st.status === 'WATCH' || st.status === 'REVIEW') counts[st.status]++;

    const tracked = q || result !== 'Not qualified' || held[code] || (st.ever && st.status !== 'REMOVED') ||
      (st.status === 'REMOVED' && st.failStreak <= S.REMOVE_MONTHS);
    if (tracked) {
      const reason = q ? '' : reasonText_(r, S, c);
      histOut.push([curMonth, runDate, code, r[c('name')], r[c('cat')], result, q ? 'Yes' : 'No', st.status, r[c('quality')], r[c('rank')],
        r[c('r1')], r[c('r3')], r[c('r5')], r[c('r10')], r[c('rsi')], r[c('aum')], reason, 'Weekly run ' + (manifest.run_id || '')]);
    }

    // changes against the previous run
    const p = prev[code];
    const name = r[c('name')], cat = r[c('cat')];
    const add = (type, from, to, detail) => changes.push([runDate, code, name, cat, type, from, to, detail || '']);
    if (!p) {
      if (!firstRun) add('New fund in universe', '', result, 'First appearance in the Regular Growth universe');
    } else {
      const pq = isQual_(p.result);
      let logged = true;
      if (!pq && q) add('Entered qualifying list', p.result, result, st.status);
      else if (pq && !q) add('Left qualifying list', p.result, result, reasonText_(r, S, c) + (st.status ? '. Status ' + st.status : ''));
      else if (p.result !== result) add('Screen result changed', p.result, result, '');
      else logged = false;
      if (!logged && p.status !== st.status && (p.status || st.status)) add('Status change', p.status, st.status, '');
      if (p.name && p.name !== name) add('Name change', p.name, name, '');
      if (p.cat && p.cat !== cat) add('Category change', p.cat, cat, '');
      const aum = r[c('aum')];
      if ((tracked) && p.aum && aum && Math.abs(aum / p.aum - 1) >= S.AUM_CHANGE_ALERT)
        add('AUM change', Math.round(p.aum), Math.round(aum), ((aum / p.aum - 1) * 100).toFixed(1) + '% since last run');
    }
    stateOut.push([code, r[c('name')], r[c('cat')], r[c('aum')], result, st.status, runDate]);
  });
  Object.keys(prev).forEach(code => {
    if (!seen[code]) changes.push([runDate, Number(code), prev[code].name, prev[code].cat, 'Removed from universe', prev[code].result, '',
      'No current Regular Growth NAV: merged, closed, renamed to a new code, or NAV stale over 30 days']);
  });

  // write status columns
  sh.getRange(2, COL.status, statusOut.length, N_STATUS).setValues(statusOut);

  // historical data: replace this month's rows
  const hsh = ss.getSheetByName(TAB.hist);
  const keepHist = readHistory_(ss).filter(r => String(r[0]) !== curMonth);
  const newHist = keepHist.concat(histOut);
  if (hsh.getLastRow() > 4) hsh.getRange(5, 1, hsh.getLastRow() - 4, HIST_HDR.length).clearContent();
  ensureRows_(hsh, newHist.length + 5);
  if (newHist.length) {
    hsh.getRange(5, 1, newHist.length, 1).setNumberFormat('@');
    hsh.getRange(5, 1, newHist.length, HIST_HDR.length).setValues(newHist);
  }

  // changes log (newest first)
  if (changes.length) {
    const csh = ss.getSheetByName(TAB.chg);
    csh.insertRowsAfter(4, changes.length);
    csh.getRange(5, 1, changes.length, CHG_HDR.length).setValues(changes);
    csh.getRange(5, 1, changes.length, 1).setNumberFormat('dd-mmm-yyyy');
  }

  // state
  stSh.clearContents();
  stSh.getRange(1, 1, 1, STATE_HDR.length).setValues([STATE_HDR]);
  if (stateOut.length) { ensureRows_(stSh, stateOut.length + 1); stSh.getRange(2, 1, stateOut.length, STATE_HDR.length).setValues(stateOut); }

  addObsMonths_([curMonth]);
  buildArchive_(ss, newHist, S);
  writeMonthly_(ss, newHist);
  return { counts: counts, changes: changes.length };
}

function reasonText_(r, S, c) {
  const per = [['r1', 'MIN_1Y', '1Y'], ['r2', 'MIN_2Y', '2Y'], ['r3', 'MIN_3Y', '3Y'], ['r5', 'MIN_5Y', '5Y'], ['r10', 'MIN_10Y', '10Y'], ['rsi', 'MIN_SI', 'SI']];
  const out = [];
  per.forEach(p => { const v = r[c(p[0])]; if (S[p[1]] !== '' && v !== '' && v < S[p[1]]) out.push(p[2] + ' ' + (v * 100).toFixed(1) + '%'); });
  const aum = r[c('aum')];
  if (aum === '' || aum < S.MIN_AUM) out.push('AUM ' + (aum === '' ? 'n/a' : Math.round(aum)) + ' cr');
  if (r[c('elig')] === false) out.push('category not screened');
  if (!out.length && r[c('r3')] === '') out.push('under 3Y history');
  return out.join('; ');
}

function buildArchive_(ss, hist, S) {
  const by = {};
  hist.forEach(r => { (by[r[2]] = by[r[2]] || []).push(r); });
  const out = [];
  Object.keys(by).forEach(code => {
    const rows = by[code].sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : a[1] - b[1]));
    // keep the last row of each month
    const monthly = [];
    rows.forEach(r => { if (monthly.length && monthly[monthly.length - 1][0] === r[0]) monthly[monthly.length - 1] = r; else monthly.push(r); });
    const qRows = monthly.filter(r => r[6] === 'Yes');
    if (!qRows.length) return;
    const last = monthly[monthly.length - 1];
    let streakStart = null;
    for (let i = monthly.length - 1; i >= 0; i--) { if (monthly[i][6] === 'Yes') streakStart = monthly[i]; else if (streakStart) break; }
    const lastQ = qRows[qRows.length - 1];
    let removed = '';
    if (last[7] === 'REMOVED') { for (let i = monthly.length - 1; i >= 0; i--) { if (monthly[i][7] === 'REMOVED') removed = monthly[i][1]; else break; } }
    let exitReason = '';
    const idx = monthly.indexOf(lastQ);
    if (idx < monthly.length - 1) exitReason = monthly[idx + 1][16];
    const ranks = monthly.map(r => r[9]).filter(x => x !== '' && x !== null);
    const prevStatus = monthly.length > 1 ? monthly[monthly.length - 2][7] : '';
    out.push([last[3], Number(code), last[4], streakStart ? streakStart[1] : '', qRows[0][1], removed, qRows.length,
      ranks.length ? Math.min.apply(null, ranks) : '', lastQ[9], last[7], prevStatus, exitReason,
      last[10], last[11], last[12], last[13], streakStart ? streakStart[15] : '', last[15], last[1]]);
  });
  out.sort((a, b) => (b[6] - a[6]) || String(a[0]).localeCompare(String(b[0])));
  const sh = ss.getSheetByName(TAB.arch);
  if (sh.getLastRow() > 4) sh.getRange(5, 1, sh.getLastRow() - 4, ARCH_HDR.length).clearContent();
  ensureRows_(sh, out.length + 5);
  if (out.length) sh.getRange(5, 1, out.length, ARCH_HDR.length).setValues(out);
}

function writeMonthly_(ss, hist) {
  const by = {};
  const lastRow = {};
  hist.forEach(r => { const k = r[0] + '|' + r[2]; if (!lastRow[k] || r[1] >= lastRow[k][1]) lastRow[k] = r; });
  Object.keys(lastRow).forEach(k => {
    const r = lastRow[k];
    const m = (by[r[0]] = by[r[0]] || [0, 0, 0]);
    if (r[5] === 'Full-history qualifier') m[0]++;
    else if (r[5] === 'Emerging qualifier') m[1]++;
    if (r[7] === 'WATCH' || r[7] === 'REVIEW') m[2]++;
  });
  obsMonths_().forEach(m => { if (!by[m]) by[m] = [0, 0, 0]; });
  const months = Object.keys(by).sort().slice(-36);
  const sh = ss.getSheetByName(TAB.dash);
  const mc = DASH.monthlyCol;
  sh.getRange(DASH.sections + 2, mc, 40, 4).clearContent();
  if (months.length) {
    const vals = months.map(m => ['\'' + m].concat(by[m]));
    sh.getRange(DASH.sections + 2, mc, vals.length, 4).setValues(vals);
  }
  drawMonthlyChart_(sh, months.length);
}

function writeSources_(ss, m, res) {
  const sh = ss.getSheetByName(TAB.src);
  sh.getRange(4, 1, 24, 6).clear();
  title_(sh, 'Data Sources', 'Where each figure comes from, when it was pulled, and validation results for every run.');
  const hdr = ['Source', 'Used for', 'Endpoint', 'Pulled (UTC)', 'Status'];
  sh.getRange(4, 1, 1, hdr.length).setValues([hdr]);
  styleHeader_(sh.getRange(4, 1, 1, hdr.length));
  const srcs = (m.sources || []).map(s => [s.name, s.used_for, s.url, m.run_utc, 'OK']);
  if (srcs.length) sh.getRange(5, 1, srcs.length, hdr.length).setValues(srcs);
  let r = 6 + srcs.length;
  const val = [['Universe rule', 'Open-ended schemes, Regular Plan, Growth option, NAV published in the last 30 days; ETFs excluded'],
    ['NAV date', m.latest_nav_date], ['Funds in universe', m.funds], ['NAV histories fetched', m.history_ok + ' ok, ' + m.history_failed + ' failed'],
    ['Kuvera AUM/manager matches', m.kuvera_ok], ['TER months loaded', m.ter_month || ''],
    ['AMFI average AUM quarters', (m.aaum_quarters || []).join(' | ')], ['Risk-free rate (Sharpe, Sortino)', m.risk_free],
    ['Return convention', '3M, 6M, 1Y point-to-point; 2Y and longer CAGR; NAV on or before each anniversary date'],
    ['Validation flags', JSON.stringify(m.flags || {})]];
  sh.getRange(r, 1, 1, 2).setValues([['Validation', '']]);
  styleHeader_(sh.getRange(r, 1, 1, 2));
  sh.getRange(r + 1, 1, val.length, 2).setValues(val);
  // run log (append, newest first) from row 30
  const logRow = 30;
  if (sh.getRange(logRow, 1).getValue() !== RUNLOG_HDR[0]) {
    sh.getRange(logRow - 1, 1).setValue('Run log').setFontWeight('bold').setFontColor(BURGUNDY);
    sh.getRange(logRow, 1, 1, RUNLOG_HDR.length).setValues([RUNLOG_HDR]);
    styleHeader_(sh.getRange(logRow, 1, 1, RUNLOG_HDR.length));
  }
  sh.insertRowsAfter(logRow, 1);
  const k = res.counts;
  sh.getRange(logRow + 1, 1, 1, RUNLOG_HDR.length).setValues([[new Date(), m.run_id, m.latest_nav_date, m.funds, k.full, k.emerging,
    k.recent, k.near, k.momentum, k.WATCH, k.REVIEW, res.changes, JSON.stringify(m.flags || {})]]).setFontWeight('normal').setFontColor('#000000').setBackground(null);
  sh.getRange(logRow + 1, 1).setNumberFormat('dd-mmm-yyyy hh:mm');
  sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 380); sh.setColumnWidth(3, 330); sh.setColumnWidth(4, 150);
  sh.getRange(1, 1, sh.getMaxRows(), 13).setFontFamily(FONT).setFontSize(9);
  sh.getRange('A1').setFontSize(15);
  // dashboard subtitle
  const nav = m.latest_nav_date ? Utilities.formatDate(new Date(m.latest_nav_date + 'T12:00:00Z'), 'Asia/Kolkata', 'dd MMM yyyy') : '';
  ss.getSheetByName(TAB.dash).getRange('A2').setValue('NAV data as of ' + nav + '. Imported ' +
    Utilities.formatDate(new Date(), 'America/New_York', 'EEE dd MMM yyyy, h:mm a') + ' ET. Pipeline refreshes every Saturday morning (US Eastern).');
}
