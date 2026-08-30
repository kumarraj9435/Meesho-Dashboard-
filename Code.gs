/**
 * Raj Meesho — Login & Single-Session Auth Backend
 * -------------------------------------------------
 * SETUP:
 * 1. Open (or create) a Google Sheet.
 * 2. Extensions -> Apps Script. Delete any default code, paste this whole file.
 * 3. Change ADMIN_PASSWORD below to your own secret.
 * 4. Deploy -> New deployment -> Type: Web app.
 *      Execute as: Me
 *      Who has access: Anyone
 * 5. Copy the deployed Web App URL.
 * 6. Paste that URL into raj_meesho.html where it says APPS_SCRIPT_URL.
 *
 * A "Users" tab is created automatically on first run.
 * To add the first user, run addFirstUserFromScript() once manually from the
 * Apps Script editor (Run button) — or use the Admin panel in the dashboard
 * once deployed (Labels-less; it's under "Admin Login" on the login screen).
 */

const ADMIN_PASSWORD = '@9435Love&&&&';   // <-- change this before deploying
const SESSION_TIMEOUT_MINUTES = 30;             // how long a session stays "active" without a heartbeat
const SHEET_NAME = 'Users';

function doGet(e) { return handle_(e); }
function doPost(e) { return handle_(e); }

function handle_(e) {
  let result;
  try {
    const p = (e && e.parameter) || {};
    const action = (p.action || '').toLowerCase();
    switch (action) {
      case 'login':            result = login_(p.userId, p.password); break;
      case 'heartbeat':        result = heartbeat_(p.userId, p.token); break;
      case 'logout':           result = logout_(p.userId, p.token); break;
      case 'adminlist':        result = adminList_(p.adminPassword); break;
      case 'adminadd':         result = adminAdd_(p.adminPassword, p.userId, p.password); break;
      case 'adminremove':      result = adminRemove_(p.adminPassword, p.userId); break;
      case 'adminforcelogout': result = adminForceLogout_(p.adminPassword, p.userId); break;
      case 'readtab':          result = readTab_(p.tab, p.sheetId); break;
      case 'writerows':        result = writeRowsToTab_HTTP(p.tab, p.rowsJson, p.sheetId); break;
      default:                 result = { ok: false, error: 'Unknown action' };
    }
  } catch (err) {
    result = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Reads a full tab's raw values and returns them as a 2D array.
 * Used by raj_meesho.html's "Google Sheet Sync" instead of the public
 * gviz CSV endpoint — going through this same Web App avoids CORS/
 * sharing issues entirely (works even when raj_meesho.html is opened as
 * a local file:// page), and needs no "Anyone with the link" sharing.
 *
 * sheetId (optional): if provided, reads from THAT spreadsheet instead
 * of the one this script is bound to — lets you pull another seller's
 * Sheet by pasting its link/ID in the tool. Works only if the Google
 * account that deployed this script (Execute as: Me) already has at
 * least Viewer access to that other Sheet (owner shared it with you,
 * or it's set to "Anyone with the link").
 */
function readTab_(tabName, sheetId) {
  if (!tabName) return { ok: false, error: 'Tab name chahiye' };
  let ss;
  try {
    ss = sheetId ? SpreadsheetApp.openById(sheetId) : SpreadsheetApp.getActiveSpreadsheet();
  } catch (err) {
    return { ok: false, error: 'Ye Sheet open nahi ho payi — ya to ID/link galat hai, ya is Sheet ka access nahi hai (Sheet owner se apna Google account add karwao, ya "Anyone with link" karwao).' };
  }
  const sh = ss.getSheetByName(tabName);
  if (!sh) return { ok: false, error: 'Tab "' + tabName + '" nahi mila is Sheet me' };
  const data = sh.getDataRange().getValues();
  return { ok: true, rows: data };
}

/**
 * HTTP-exposed wrapper around importRowsToTab() — lets raj_meesho.html write
 * (append/dedupe/update) rows into any configured tab directly from the
 * browser via fetch(), the same way readtab already lets it read a tab.
 * Used by the GST Summary panel's "Google Sheet Sync" to persist raw GST
 * sales/return rows and the invoice/credit-note numbering log across months.
 *
 * rowsJson: a JSON-encoded 2D array (including its header row at index 0),
 * exactly like importRowsToTab() already expects from the sidebar importer.
 * sheetId (optional): write into a different spreadsheet, same rules as readtab.
 */
function writeRowsToTab_HTTP(tabName, rowsJson, sheetId) {
  if (!tabName || !rowsJson) return { ok: false, error: 'tab aur rowsJson dono chahiye' };
  let rows2D;
  try {
    rows2D = JSON.parse(rowsJson);
  } catch (err) {
    return { ok: false, error: 'rowsJson valid JSON nahi hai: ' + err };
  }
  if (!Array.isArray(rows2D) || !rows2D.length) return { ok: false, error: 'rowsJson khali hai' };
  if (!sheetId) return importRowsToTab(tabName, rows2D);

  // sheetId provided — same logic as importRowsToTab but against a different spreadsheet
  let ss;
  try {
    ss = SpreadsheetApp.openById(sheetId);
  } catch (err) {
    return { ok: false, error: 'Ye Sheet open nahi ho payi — ID/link galat hai ya access nahi hai.' };
  }
  return importRowsToTabInSpreadsheet_(ss, tabName, rows2D);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['UserID', 'PasswordHash', 'SessionToken', 'SessionDevice', 'LastActive']);
  }
  return sh;
}

function hash_(s) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8);
  return raw.map(b => ((b < 0 ? b + 256 : b).toString(16).padStart(2, '0'))).join('');
}

function findUserRow_(sh, userId) {
  const data = sh.getDataRange().getValues();
  const target = String(userId || '').trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === target) return i + 1; // 1-indexed sheet row
  }
  return -1;
}

function login_(userId, password) {
  if (!userId || !password) return { ok: false, error: 'User ID and Password are both required' };
  const sh = getSheet_();
  const row = findUserRow_(sh, userId);
  if (row === -1) return { ok: false, error: 'User not found' };
  const [uid, passHash, sessionToken, sessionDevice, lastActive] = sh.getRange(row, 1, 1, 5).getValues()[0];
  if (hash_(password) !== passHash) return { ok: false, error: 'Incorrect password' };

  const now = new Date();
  if (sessionToken && lastActive) {
    const diffMin = (now - new Date(lastActive)) / 60000;
    if (diffMin < SESSION_TIMEOUT_MINUTES) {
      return { ok: false, error: 'This User ID is already logged in on another device/browser. Try again in a few (' + SESSION_TIMEOUT_MINUTES + ' min), or ask an admin to force-logout that session.' };
    }
  }
  const token = Utilities.getUuid();
  sh.getRange(row, 3, 1, 3).setValues([[token, 'device-' + Utilities.getUuid().slice(0, 8), now]]);
  return { ok: true, token: token };
}

function heartbeat_(userId, token) {
  const sh = getSheet_();
  const row = findUserRow_(sh, userId);
  if (row === -1) return { ok: false, error: 'User not found' };
  const stored = sh.getRange(row, 3).getValue();
  if (stored !== token) return { ok: false, error: 'Session is no longer valid — someone else has logged in with this ID' };
  sh.getRange(row, 5).setValue(new Date());
  return { ok: true };
}

function logout_(userId, token) {
  const sh = getSheet_();
  const row = findUserRow_(sh, userId);
  if (row === -1) return { ok: true };
  const stored = sh.getRange(row, 3).getValue();
  if (stored === token) sh.getRange(row, 3, 1, 3).setValues([['', '', '']]);
  return { ok: true };
}

function checkAdmin_(pw) { return pw === ADMIN_PASSWORD; }

function adminList_(adminPassword) {
  if (!checkAdmin_(adminPassword)) return { ok: false, error: 'Incorrect admin password' };
  const sh = getSheet_();
  const data = sh.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < data.length; i++) {
    const lastActive = data[i][4];
    const active = !!data[i][2] && lastActive && ((new Date() - new Date(lastActive)) / 60000) < SESSION_TIMEOUT_MINUTES;
    users.push({ userId: data[i][0], active: active, lastActive: lastActive ? new Date(lastActive).toLocaleString() : '' });
  }
  return { ok: true, users: users };
}

function adminAdd_(adminPassword, userId, password) {
  if (!checkAdmin_(adminPassword)) return { ok: false, error: 'Incorrect admin password' };
  if (!userId || !password) return { ok: false, error: 'User ID and Password are both required' };
  const sh = getSheet_();
  const row = findUserRow_(sh, userId);
  const hashed = hash_(password);
  if (row === -1) sh.appendRow([userId, hashed, '', '', '']);
  else sh.getRange(row, 2).setValue(hashed);
  return { ok: true };
}

function adminRemove_(adminPassword, userId) {
  if (!checkAdmin_(adminPassword)) return { ok: false, error: 'Incorrect admin password' };
  const sh = getSheet_();
  const row = findUserRow_(sh, userId);
  if (row === -1) return { ok: false, error: 'User not found' };
  sh.deleteRow(row);
  return { ok: true };
}

function adminForceLogout_(adminPassword, userId) {
  if (!checkAdmin_(adminPassword)) return { ok: false, error: 'Incorrect admin password' };
  const sh = getSheet_();
  const row = findUserRow_(sh, userId);
  if (row === -1) return { ok: false, error: 'User not found' };
  sh.getRange(row, 3, 1, 3).setValues([['', '', '']]);
  return { ok: true };
}

/**
 * Optional helper — run this once from the Apps Script editor (select this
 * function in the dropdown, click Run) to create your very first login.
 * Change the values below first.
 */
function addFirstUserFromScript() {
  const sh = getSheet_();
  const userId = 'UG';           // <-- change
  const password = 'ChangeMe123'; // <-- change
  const row = findUserRow_(sh, userId);
  const hashed = hash_(password);
  if (row === -1) sh.appendRow([userId, hashed, '', '', '']);
  else sh.getRange(row, 2).setValue(hashed);
  Logger.log('User added/updated: ' + userId);
}

/* ============================================================
 * IN-SHEET FILE IMPORT — custom menu + sidebar
 * Lets you upload Orders/Payment/AdsCost/RateCard/ExtraExpenses/GST
 * files directly into the Sheet. New rows get appended below
 * existing ones; duplicates (by key) are skipped automatically.
 * ============================================================ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Raj Meesho')
    .addItem('🛠️ One-Time Setup (Create All Tabs)', 'setupAllTabs')
    .addItem('📥 Import File to a Tab', 'showImportSidebar')
    .addSeparator()
    .addItem('📈 Build / Refresh Dashboard Tab', 'buildDashboard')
    .addItem('🎨 Colorful Style — All Tabs', 'styleAllTabs')
    .addSeparator()
    .addItem('💾 Backup All Data (Download)', 'backupAllData')
    .addItem('🗑️ Delete All Data', 'deleteAllData')
    .addToUi();
}

/**
 * DASHBOARD — builds a live, formula-driven "Dashboard" tab inside the
 * Sheet itself (separate from the Raj Meesho HTML tool). Numbers update
 * automatically whenever PreviousPayment/OutstandingPayment/AdsCost/
 * Orders/ExtraExpenses tabs change — no script re-run needed for that.
 * Safe to run again anytime to rebuild/refresh the layout.
 */
function buildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Dashboard');
  if (sh) ss.deleteSheet(sh); // pure formulas/labels — safe to rebuild from scratch
  sh = ss.insertSheet('Dashboard', 0);

  const PLUM = '#3D1152', MAGENTA = '#C6167B', PAPER = '#FBF8FA', GREEN = '#1D8A5D', RED = '#D6403A', MUTED = '#6B6273';
  const BIGRANGE = 100000; // safe upper bound for SUM/COUNTIF ranges

  // ---- Title banner ----
  sh.getRange('A1:H2').merge().setValue('📊 Raj Meesho — Live Dashboard')
    .setBackground(PLUM).setFontColor('#FFFFFF').setFontSize(20).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(1, 30); sh.setRowHeight(2, 30);

  // ---- KPI cards (row 4-6) ----
  const kpis = [
    { label: 'Total Orders', formula: '=COUNTA(Orders!B2:B' + BIGRANGE + ')', color: PLUM },
    { label: 'Delivered', formula: '=COUNTIF(Orders!A2:A' + BIGRANGE + ',"*DELIVER*")', color: GREEN },
    { label: 'Cancelled', formula: '=COUNTIF(Orders!A2:A' + BIGRANGE + ',"*CANCEL*")', color: RED },
    { label: 'RTO', formula: '=COUNTIF(Orders!A2:A' + BIGRANGE + ',"*RTO*")', color: '#B67900' },
    { label: 'Total Sale Amount', formula: '=SUM(PreviousPayment!P4:P' + BIGRANGE + ')+SUM(OutstandingPayment!P4:P' + BIGRANGE + ')', color: PLUM, money: true },
    { label: 'Net Settlement', formula: '=SUM(PreviousPayment!N4:N' + BIGRANGE + ')+SUM(OutstandingPayment!N4:N' + BIGRANGE + ')', color: GREEN, money: true },
    { label: 'Ads Spend', formula: '=-SUM(AdsCost!H4:H' + BIGRANGE + ')', color: RED, money: true },
    { label: 'Extra Expenses', formula: '=SUM(ExtraExpenses!D2:D' + BIGRANGE + ')', color: RED, money: true },
  ];
  let col = 1;
  kpis.forEach(k => {
    const labelCell = sh.getRange(4, col);
    const valueCell = sh.getRange(5, col);
    labelCell.setValue(k.label).setFontSize(9).setFontColor(MUTED).setFontWeight('bold');
    valueCell.setFormula(k.formula).setFontSize(16).setFontWeight('bold').setFontColor(k.color);
    if (k.money) valueCell.setNumberFormat('₹#,##0');
    sh.getRange(4, col, 2, 1).setBackground(PAPER).setBorder(true, true, true, true, false, false, '#E7DEE9', SpreadsheetApp.BorderStyle.SOLID);
    col++;
  });

  // ---- Net payout estimate (row 7-8) ----
  sh.getRange('A7:D7').merge().setValue('Net Payout (Settlement − Ads − Extra Expenses)').setFontWeight('bold').setFontColor(MUTED).setFontSize(10);
  sh.getRange('A8:D8').merge().setFormula('=F5-G5-H5').setFontSize(18).setFontWeight('bold').setFontColor(PLUM).setNumberFormat('₹#,##0');
  sh.getRange('A9:D9').merge().setValue('Note: Cost of Goods (Rate Card) is not counted here — full Net Profit with COGS is in the Raj Meesho tool.').setFontSize(9).setFontColor(MUTED).setFontStyle('italic');

  // ---- Status breakdown table (row 11+) ----
  sh.getRange('A11').setValue('Order Status Breakdown').setFontWeight('bold').setFontSize(12).setFontColor(PLUM);
  const statusRows = [
    ['Delivered', '=COUNTIF(Orders!A2:A' + BIGRANGE + ',"*DELIVER*")'],
    ['Cancelled', '=COUNTIF(Orders!A2:A' + BIGRANGE + ',"*CANCEL*")'],
    ['RTO', '=COUNTIF(Orders!A2:A' + BIGRANGE + ',"*RTO*")'],
    ['Other / In Transit', '=COUNTA(Orders!A2:A' + BIGRANGE + ')-COUNTIF(Orders!A2:A' + BIGRANGE + ',"*DELIVER*")-COUNTIF(Orders!A2:A' + BIGRANGE + ',"*CANCEL*")-COUNTIF(Orders!A2:A' + BIGRANGE + ',"*RTO*")']
  ];
  sh.getRange(12, 1, statusRows.length, 2).setValues(statusRows.map(r => [r[0], '']));
  for (let i = 0; i < statusRows.length; i++) sh.getRange(12 + i, 2).setFormula(statusRows[i][1]);
  sh.getRange(12, 1, statusRows.length, 2).setBorder(true, true, true, true, true, true, '#E7DEE9', SpreadsheetApp.BorderStyle.SOLID);

  // ---- Top expense categories (row 11+, next to status table) ----
  sh.getRange('D11').setValue('Top Expense Categories (from Payments)').setFontWeight('bold').setFontSize(12).setFontColor(PLUM);
  const expenseRows = [
    ['Commission', '=SUM(PreviousPayment!W4:W' + BIGRANGE + ')+SUM(OutstandingPayment!W4:W' + BIGRANGE + ')'],
    ['Fixed Fee', '=SUM(PreviousPayment!R4:R' + BIGRANGE + ')+SUM(PreviousPayment!Z4:Z' + BIGRANGE + ')+SUM(OutstandingPayment!R4:R' + BIGRANGE + ')+SUM(OutstandingPayment!Z4:Z' + BIGRANGE + ')'],
    ['Shipping Charge', '=SUM(PreviousPayment!AD4:AD' + BIGRANGE + ')+SUM(OutstandingPayment!AD4:AD' + BIGRANGE + ')'],
    ['TCS', '=SUM(PreviousPayment!AI4:AI' + BIGRANGE + ')+SUM(OutstandingPayment!AI4:AI' + BIGRANGE + ')'],
    ['TDS', '=SUM(PreviousPayment!AK4:AK' + BIGRANGE + ')+SUM(OutstandingPayment!AK4:AK' + BIGRANGE + ')'],
  ];
  sh.getRange(12, 4, expenseRows.length, 2).setValues(expenseRows.map(r => [r[0], '']));
  for (let i = 0; i < expenseRows.length; i++) {
    sh.getRange(12 + i, 5).setFormula(expenseRows[i][1]).setNumberFormat('₹#,##0');
  }
  sh.getRange(12, 4, expenseRows.length, 2).setBorder(true, true, true, true, true, true, '#E7DEE9', SpreadsheetApp.BorderStyle.SOLID);

  // ---- cosmetics ----
  sh.setColumnWidths(1, 8, 130);
  sh.setTabColor(MAGENTA);
  sh.setFrozenRows(3);
  sh.getRange('A1').getSheet().setActiveSelection('A1');

  SpreadsheetApp.getUi().alert('✅ Dashboard tab ban gaya / refresh ho gaya. Numbers khud-b-khud update honge jab bhi data change hoga.');
}

/**
 * COLORFUL STYLING — applies a consistent, good-looking color scheme to
 * every tab: colored tab labels, styled header rows, frozen headers,
 * alternating row banding on data, auto-sized columns. Safe to re-run.
 */
function styleAllTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const PLUM = '#3D1152';
  const tabColors = {
    'Dashboard': '#C6167B', 'Orders': '#3D1152', 'PreviousPayment': '#5B2A86',
    'OutstandingPayment': '#7A3FA8', 'AdsCost': '#B67900', 'RateCard': '#1D8A5D',
    'ExtraExpenses': '#D6403A', 'Users': '#6B6273',
    'GSTSalesRaw': '#0E7C86', 'GSTReturnsRaw': '#17ADBB', 'GSTFilingLog': '#2255A8'
  };
  const headerRowsMap = { 'Orders': 1, 'PreviousPayment': 3, 'OutstandingPayment': 3, 'AdsCost': 3, 'RateCard': 1, 'ExtraExpenses': 1, 'Users': 1, 'GSTSalesRaw': 1, 'GSTReturnsRaw': 1, 'GSTFilingLog': 1 };

  Object.keys(tabColors).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.setTabColor(tabColors[name]);

    const headerRows = headerRowsMap[name] || 1;
    const lastCol = Math.max(sh.getLastColumn(), 1);
    const lastRow = sh.getLastRow();

    sh.getRange(1, 1, headerRows, lastCol)
      .setBackground(PLUM).setFontColor('#FFFFFF').setFontWeight('bold');
    sh.setFrozenRows(headerRows);

    // clear any old banding first (Sheets errors if you double-band a range)
    sh.getBandings().forEach(b => b.remove());
    if (lastRow > headerRows) {
      sh.getRange(headerRows + 1, 1, lastRow - headerRows, lastCol)
        .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
    }
    sh.autoResizeColumns(1, lastCol);
  });

  SpreadsheetApp.getUi().alert('✅ Saari sheets colorful ho gayi — tab colors, header styling, aur row banding sab set ho gaya.');
}

/**
 * BACKUP — makes a full timestamped copy of this entire spreadsheet in
 * Google Drive (same folder). Safe, non-destructive. Run this before
 * "Delete All Data", or anytime you want a snapshot.
 */
function backupAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(ss.getId());
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm');
  const backupName = ss.getName() + ' — Backup ' + ts;
  const parents = file.getParents();
  const folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const copy = file.makeCopy(backupName, folder);
  SpreadsheetApp.getUi().alert(
    '✅ Backup ban gaya',
    'Naam: ' + backupName + '\n\nYe file Google Drive me (isi folder me) save ho gayi hai. Wahan se kholke File → Download se Excel/CSV bhi nikaal sakte ho.\n\nLink:\n' + copy.getUrl(),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * DELETE ALL DATA — clears all data rows (keeps header rows) from every
 * business-data tab. Does NOT touch the Users/login tab. Asks for
 * confirmation first since this can't be undone (always Backup first).
 */
function deleteAllData() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '⚠️ Sab data delete karna hai?',
    'Ye Orders, PreviousPayment, OutstandingPayment, AdsCost, RateCard, ExtraExpenses, aur GST (GSTSalesRaw/GSTReturnsRaw/GSTFilingLog) tabs ka pura data hata dega (sirf headers rahenge). Users/login data safe rahega, delete nahi hoga.\n\nPehle "Backup All Data" zaroor le lo agar nahi liya.\n\nContinue karna hai?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) { ui.alert('Cancel kar diya — kuch delete nahi hua.'); return; }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tabsToClear = {
    'Orders': 1, 'PreviousPayment': 3, 'OutstandingPayment': 3,
    'AdsCost': 3, 'RateCard': 1, 'ExtraExpenses': 1,
    'GSTSalesRaw': 1, 'GSTReturnsRaw': 1, 'GSTFilingLog': 1
  };
  const cleared = [];
  Object.keys(tabsToClear).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const headerRows = tabsToClear[name];
    const lastRow = sh.getLastRow();
    if (lastRow > headerRows) {
      sh.getRange(headerRows + 1, 1, lastRow - headerRows, sh.getLastColumn()).clearContent();
      cleared.push(name);
    }
  });
  ui.alert('✅ Delete ho gaya', cleared.length ? (cleared.join(', ') + ' clear ho gaye.') : 'Sab tabs pehle se hi khali the.', ui.ButtonSet.OK);
}

/**
 * ONE-TIME SETUP — run once (from this menu, or Apps Script editor Run
 * button) to create every tab Raj Meesho needs, with correct headers
 * already in place. Safe to run again later — it skips any tab that
 * already exists, so it never wipes real data.
 */
function setupAllTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const created = [], skipped = [];

  function ensureTab(name, rows) {
    if (ss.getSheetByName(name)) { skipped.push(name); return; }
    const sh = ss.insertSheet(name);
    sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    sh.setFrozenRows(rows.length);
    created.push(name);
  }

  // Orders — same headers as the Meesho "Orders" export CSV
  ensureTab('Orders', [[
    'Reason for Credit Entry', 'Sub Order No', 'Catalog ID', 'Order Date', 'Order source',
    'Customer State', 'Product Name', 'SKU', 'Size', 'Quantity',
    'Supplier Listed Price (Incl. GST + Commission)', 'Supplier Discounted Price (Incl GST and Commision)', 'Packet Id'
  ]]);

  // PreviousPayment / OutstandingPayment — 2 blank/category rows + 1 real
  // label row (3 header rows total), 40 columns matching Meesho's
  // "Order Payments" sheet column order (position-based, not name-based).
  const paymentHeaderRow3 = [
    'Sub Order No','Order Date','Dispatch Date','Product Name','SKU','Catalog ID','Order source','Live Order Status',
    'Product GST %','Listing Price (Incl. taxes)','Quantity','Transaction ID','Payment Date','Final Settlement Amount',
    'Price Type','Total Sale Amount (Incl. Shipping & GST)','Total Sale Return Amount (Incl. Shipping & GST)',
    'Fixed Fee (Incl. GST)','Warehousing fee (Incl. GST)','Return premium (Incl. GST)','Return premium (Incl. GST) of Return',
    'Meesho Commission Percentage','Meesho Commission (Incl. GST)','Meesho Gold Platform Fee (Incl. GST)','Meesho Mall Platform Fee (Incl. GST)',
    'Fixed Fee (Incl. GST)','Warehousing Fee (Incl. GST)','Return Shipping Charge (Incl. GST)','GST Compensation (PRP Shipping)',
    'Shipping Charge (Incl. GST)','Other Support Service Charges (Excl. GST)','Waivers (Excl. GST)','Net Other Support Service Charges (Excl. GST)',
    'GST on Net Other Support Service Charges','TCS','TDS Rate %','TDS','Compensation','Claims','Recovery'
  ];
  const blankRow40 = new Array(40).fill('');
  ['PreviousPayment', 'OutstandingPayment'].forEach(name => {
    ensureTab(name, [blankRow40, blankRow40, paymentHeaderRow3]);
  });

  // AdsCost — 2 blank rows + 1 label row (3 header rows total), 8 columns
  const adsHeaderRow3 = ['Duration', 'Deduction Date', 'Campaign ID', 'Ad Cost', 'Waivers', 'Ad Cost Net', 'GST', 'Total Ads Cost'];
  const blankRow8 = new Array(8).fill('');
  ensureTab('AdsCost', [blankRow8, blankRow8, adsHeaderRow3]);

  // RateCard
  ensureTab('RateCard', [['SKU', 'Product', 'Cost Price']]);

  // ExtraExpenses
  ensureTab('ExtraExpenses', [['Date', 'Category', 'Description', 'Amount']]);

  // GST — raw sales/return rows kept here so numbering & totals persist
  // across months, plus the filing log that remembers invoice/credit-note
  // ranges already used so the next month's numbering can continue on.
  const gstRawHeader = ['SubOrderNo', 'OrderDate', 'HSN', 'Qty', 'GSTRate', 'TaxableValue', 'TaxAmount', 'InvoiceValue', 'State', 'Period', 'Portal', 'GSTIN'];
  ensureTab('GSTSalesRaw', [gstRawHeader]);
  ensureTab('GSTReturnsRaw', [gstRawHeader]);
  ensureTab('GSTFilingLog', [['Period', 'InvoiceFrom', 'InvoiceTo', 'InvoiceCount', 'CreditNoteFrom', 'CreditNoteTo', 'CreditNoteCount', 'SavedAt', 'Portal', 'GSTIN']]);

  // Users (also auto-created by getSheet_ on first login, but nice to have upfront)
  if (!ss.getSheetByName(SHEET_NAME)) { getSheet_(); created.push(SHEET_NAME); }
  else skipped.push(SHEET_NAME);

  const msg = (created.length ? 'Bane: ' + created.join(', ') : 'Kuch nahi bana') +
    '\n' + (skipped.length ? 'Already the (skip kiye): ' + skipped.join(', ') : '');
  SpreadsheetApp.getUi().alert('✅ Setup complete', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function showImportSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Upload')
    .setTitle('Raj Meesho — Import File')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Per-tab dedupe rules. headerRows = how many header rows sit above the data.
const TAB_CONFIGS_ = {
  'Orders':              { headerRows: 1, keyHeader: 'Sub Order No' },
  'PreviousPayment':     { headerRows: 3, keyColIndex: 0 },              // col A = Sub Order No
  'OutstandingPayment':  { headerRows: 3, keyColIndex: 0 },
  'AdsCost':             { headerRows: 3, keyColIndexes: [0, 1, 2] },    // Duration+Date+CampaignId combo
  'RateCard':            { headerRows: 1, keyHeader: 'SKU', updateMode: true }, // existing SKU -> update cost instead of duplicate row
  'ExtraExpenses':       { headerRows: 1, keyHeader: null },             // dedupe by full row content
  'GSTSalesRaw':         { headerRows: 1, keyHeaders: ['Portal','GSTIN','SubOrderNo'] }, // skip rows already stored for this seller+sub-order
  'GSTReturnsRaw':       { headerRows: 1, keyHeaders: ['Portal','GSTIN','SubOrderNo'] },
  'GSTFilingLog':        { headerRows: 1, keyHeaders: ['Portal','GSTIN','Period'], updateMode: true } // re-saving the same seller+period updates its row
};

function normalizeRows_(rows2D) {
  const maxLen = Math.max.apply(null, rows2D.map(r => r.length));
  return rows2D.map(r => {
    const copy = r.slice();
    while (copy.length < maxLen) copy.push('');
    return copy;
  });
}

/**
 * Called from the sidebar (Upload.html) via google.script.run, AND from
 * raj_meesho.html directly over HTTP via the "writerows" action (see
 * writeRowsToTab_HTTP above) — both paths share this same function.
 * rows2D: full 2D array exactly as read from the uploaded file / built
 * in the browser, INCLUDING its header row(s) at the top.
 */
function importRowsToTab(tabName, rows2D) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return importRowsToTabInSpreadsheet_(ss, tabName, rows2D);
}

function importRowsToTabInSpreadsheet_(ss, tabName, rows2D) {
  const cfg = TAB_CONFIGS_[tabName];
  if (!cfg) return { ok: false, error: 'Unknown tab: ' + tabName };
  if (!rows2D || rows2D.length === 0) return { ok: false, error: 'Koi data nahi mila file me' };
  rows2D = normalizeRows_(rows2D);

  let sh = ss.getSheetByName(tabName);
  if (!sh) sh = ss.insertSheet(tabName);

  const existingLastRow = sh.getLastRow();
  const headerRows = cfg.headerRows;

  // Empty tab -> just write everything (headers + data) as-is
  if (existingLastRow === 0) {
    sh.getRange(1, 1, rows2D.length, rows2D[0].length).setValues(rows2D);
    return { ok: true, appended: Math.max(rows2D.length - headerRows, 0), updated: 0, skipped: 0 };
  }

  const incomingDataRows = rows2D.slice(headerRows);
  const existingData = sh.getDataRange().getValues();
  const existingDataRows = existingData.slice(headerRows);

  let keyColIndex = (cfg.keyColIndex !== undefined) ? cfg.keyColIndex : null;
  if (cfg.keyHeader) {
    const headerRowForLookup = existingData[headerRows - 1] || existingData[0];
    const idx = headerRowForLookup.findIndex(h => String(h).trim() === cfg.keyHeader);
    keyColIndex = idx === -1 ? 0 : idx;
  }
  let keyColIndexesByHeader = null;
  if (cfg.keyHeaders) {
    const headerRowForLookup = existingData[headerRows - 1] || existingData[0];
    keyColIndexesByHeader = cfg.keyHeaders.map(h => {
      const idx = headerRowForLookup.findIndex(x => String(x).trim() === h);
      return idx === -1 ? 0 : idx;
    });
  }

  function makeKey(row) {
    if (cfg.keyColIndexes) return cfg.keyColIndexes.map(i => row[i]).join('|');
    if (keyColIndexesByHeader) return keyColIndexesByHeader.map(i => String(row[i] || '').trim().toLowerCase()).join('|');
    if (keyColIndex !== null) return String(row[keyColIndex] || '').trim().toLowerCase();
    return row.join('|').toLowerCase(); // full-row dedupe fallback
  }

  const existingKeys = new Map(); // key -> sheet row number (for update mode)
  existingDataRows.forEach((row, i) => {
    const k = makeKey(row);
    if (k) existingKeys.set(k, headerRows + i + 1);
  });

  let appended = 0, updated = 0, skipped = 0;
  const rowsToAppend = [];
  incomingDataRows.forEach(row => {
    if (row.every(c => c === '' || c === null || c === undefined)) return; // skip blank rows
    const key = makeKey(row);
    if (key && existingKeys.has(key)) {
      if (cfg.updateMode) {
        sh.getRange(existingKeys.get(key), 1, 1, row.length).setValues([row]);
        updated++;
      } else {
        skipped++;
      }
    } else {
      rowsToAppend.push(row);
      appended++;
    }
  });

  if (rowsToAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  }

  return { ok: true, appended, updated, skipped };
}
