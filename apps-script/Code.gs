/**
 * Unit Conversion Calculator
 *
 * This spreadsheet has three pieces of automation, each scoped to its own
 * named tab so edits on one sheet never misfire calculations meant for
 * another:
 *
 * 1. Orders sheet (ORDER_SHEET_NAME) — reads the pack size embedded in the
 *    item Name (Column C) — e.g. "1 LTR", "500 GM", "2.75KG", "1X16KG",
 *    "1Kg X 12 pcs" — multiplies it by the pack count (if an "X n"
 *    multiplier is present) and by the ordered Quantity (Column P), then
 *    writes the result to Column T:
 *      - Weight units (G/GM/GRAM/KG) are expressed in KG.
 *      - Volume units (LTR/ML/CL) are left in their own unit (no density
 *        assumption is made to convert liters to kilograms).
 *      - If a unit keyword is found but with no number in front of it
 *        (e.g. "... SPA KGS"), the size defaults to 1.
 *      - If no unit keyword at all can be found in the Name, Column T is
 *        left untouched.
 *
 * 2. Dashboard sheet (DASHBOARD_SHEET_NAME) — Column C holds a raw gram
 *    total; Column F gets that value converted to KG, rounded UP to the
 *    nearest whole number (e.g. 16450 -> 17, 9199 -> 10).
 *
 * 3. Chilled Orders sheet (CHILLED_SHEET_NAME) — for every row, looks up
 *    its Internal Name (Column A) against the Dashboard's Internal Name
 *    (Column B) and copies the matching converted KG value (Dashboard
 *    Column F) into Column E (DB Consumption). Rows with no match in
 *    Dashboard get 0. If the Dashboard has more than one row with the same
 *    Internal Name, the last one (bottom-most row) wins.
 *
 * Runs automatically: onOpen() backfills every sheet, onEdit() updates
 * incrementally whenever a relevant cell changes.
 */

// ---- Sheet/tab names — verify these match your actual tabs ----
var ORDER_SHEET_NAME = 'Orders'; // TODO: set to your actual Forecast Order tab name
var DASHBOARD_SHEET_NAME = 'Dashboard';
var CHILLED_SHEET_NAME = 'Chilled Orders';

// ---- Orders sheet layout ----
var ORDER_NAME_COL = 3;    // Column C - Name
var ORDER_QTY_COL = 16;    // Column P - Quantity
var ORDER_RESULT_COL = 20; // Column T - Converted Total
var ORDER_HEADER_ROW = 1;
var ORDER_RESULT_HEADER = 'Converted Total';

// ---- Dashboard sheet layout ----
var DASH_NAME_COL = 2; // Column B - Internal Name
var DASH_QTY_COL = 3;  // Column C - raw grams
var DASH_KG_COL = 6;   // Column F - Converted KG
var DASH_HEADER_ROW = 1;

// ---- Chilled Orders sheet layout ----
var CHILLED_NAME_COL = 1;        // Column A - Internal Name
var CHILLED_CONSUMPTION_COL = 5; // Column E - DB Consumption
var CHILLED_HEADER_ROW = 1;

var UNIT_ALTERNATION = '(KGS|KG|GRAMS|GRAM|GMS|GM|G|LTRS|LTR|ML|CL|L)';
var SIZE_UNIT_REGEX = new RegExp('(\\d+(?:\\.\\d+)?)\\s*' + UNIT_ALTERNATION + '\\b', 'i');
var UNIT_ONLY_REGEX = new RegExp('\\b' + UNIT_ALTERNATION + '\\b', 'i');
var PACK_BEFORE_REGEX = /(\d+(?:\.\d+)?)\s*[xX]\s*$/;
var PACK_AFTER_REGEX = /^\s*[xX]\s*(\d+(?:\.\d+)?)/;

var WEIGHT_UNITS = ['G', 'GM', 'GMS', 'GRAM', 'GRAMS'];
var KG_UNITS = ['KG', 'KGS'];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Unit Tools')
    .addItem('Run All Conversions', 'runAll')
    .addItem('Recalculate Orders (KG Conversion)', 'recalcAllOrders')
    .addItem('Recalculate Dashboard (KG Conversion)', 'recalcAllDashboard')
    .addItem('Sync Chilled Orders Consumption', 'syncChilledOrdersConsumption')
    .addToUi();
  runAll();
}

function runAll() {
  recalcAllOrders();
  recalcAllDashboard();
  syncChilledOrdersConsumption();
}

function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  var col = e.range.getColumn();
  var row = e.range.getRow();

  if (sheetName === ORDER_SHEET_NAME) {
    if (row <= ORDER_HEADER_ROW) return;
    if (col === ORDER_NAME_COL || col === ORDER_QTY_COL) {
      recalcOrderRow(sheet, row);
    }
    return;
  }

  if (sheetName === DASHBOARD_SHEET_NAME) {
    if (row <= DASH_HEADER_ROW) return;
    if (col === DASH_QTY_COL) {
      recalcDashboardRow(sheet, row);
      syncChilledOrdersConsumption();
    } else if (col === DASH_NAME_COL) {
      syncChilledOrdersConsumption();
    }
    return;
  }

  if (sheetName === CHILLED_SHEET_NAME) {
    if (row <= CHILLED_HEADER_ROW) return;
    if (col === CHILLED_NAME_COL) {
      syncChilledOrdersConsumption();
    }
    return;
  }
}

// ============================= Orders sheet =============================

function recalcAllOrders() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(ORDER_SHEET_NAME);
  if (!sheet) return;

  sheet.getRange(ORDER_HEADER_ROW, ORDER_RESULT_COL).setValue(ORDER_RESULT_HEADER);

  var lastRow = sheet.getLastRow();
  for (var row = ORDER_HEADER_ROW + 1; row <= lastRow; row++) {
    recalcOrderRow(sheet, row);
  }
}

function recalcOrderRow(sheet, row) {
  var name = sheet.getRange(row, ORDER_NAME_COL).getValue();
  var qty = sheet.getRange(row, ORDER_QTY_COL).getValue();

  if (!name || typeof name !== 'string' || qty === '' || isNaN(qty)) {
    return; // nothing usable to calculate; leave Column T as-is
  }

  var parsed = parseSizeAndUnit(name);
  if (!parsed) {
    return; // no unit keyword found in the name; leave Column T as-is
  }

  var totalPerUnit = parsed.sizeNum * parsed.packCount;
  var grandTotal = totalPerUnit * qty;

  var value, unitLabel;
  if (WEIGHT_UNITS.indexOf(parsed.unitRaw) !== -1) {
    value = grandTotal / 1000;
    unitLabel = 'KG';
  } else if (KG_UNITS.indexOf(parsed.unitRaw) !== -1) {
    value = grandTotal;
    unitLabel = 'KG';
  } else {
    value = grandTotal;
    unitLabel = normalizeVolumeUnit(parsed.unitRaw);
  }

  value = Math.round(value * 1000) / 1000;
  sheet.getRange(row, ORDER_RESULT_COL).setValue(value + ' ' + unitLabel);
}

function normalizeVolumeUnit(unitRaw) {
  if (unitRaw === 'LTR' || unitRaw === 'LTRS' || unitRaw === 'L') return 'LTR';
  return unitRaw; // ML, CL
}

/**
 * Finds the first "<number><unit>" occurrence in the name (e.g. "349GM",
 * "2.75KG", "1 LTR") and looks immediately before/after it for a
 * "<n> X" / "X <n>" pack-count multiplier (e.g. "1X16KG", "1Kg X 12 pcs").
 * If a unit keyword is present but with no number in front of it (e.g.
 * "... SPA KGS"), the size defaults to 1. Returns null only if no unit
 * keyword at all is present.
 */
function parseSizeAndUnit(name) {
  var sizeNum, unitRaw, matchIndex, matchEnd;

  var match = name.match(SIZE_UNIT_REGEX);
  if (match) {
    sizeNum = parseFloat(match[1]);
    unitRaw = match[2].toUpperCase();
    matchIndex = match.index;
    matchEnd = matchIndex + match[0].length;
  } else {
    var unitOnlyMatch = name.match(UNIT_ONLY_REGEX);
    if (!unitOnlyMatch) return null;
    sizeNum = 1;
    unitRaw = unitOnlyMatch[1].toUpperCase();
    matchIndex = unitOnlyMatch.index;
    matchEnd = matchIndex + unitOnlyMatch[0].length;
  }

  var packCount = 1;
  var before = name.substring(Math.max(0, matchIndex - 15), matchIndex);
  var beforeMatch = before.match(PACK_BEFORE_REGEX);
  if (beforeMatch) {
    packCount = parseFloat(beforeMatch[1]);
  } else {
    var after = name.substring(matchEnd, matchEnd + 15);
    var afterMatch = after.match(PACK_AFTER_REGEX);
    if (afterMatch) {
      packCount = parseFloat(afterMatch[1]);
    }
  }

  return { sizeNum: sizeNum, unitRaw: unitRaw, packCount: packCount };
}

// ============================ Dashboard sheet ============================

function recalcAllDashboard() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  for (var row = DASH_HEADER_ROW + 1; row <= lastRow; row++) {
    recalcDashboardRow(sheet, row);
  }
}

function recalcDashboardRow(sheet, row) {
  var grams = sheet.getRange(row, DASH_QTY_COL).getValue();
  if (grams === '' || isNaN(grams)) {
    return; // nothing usable to calculate; leave Column F as-is
  }

  var kg = Math.ceil(grams / 1000);
  sheet.getRange(row, DASH_KG_COL).setValue(kg);
}

// ========================== Chilled Orders sheet ==========================

function syncChilledOrdersConsumption() {
  var ss = SpreadsheetApp.getActive();
  var dashSheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  var chilledSheet = ss.getSheetByName(CHILLED_SHEET_NAME);
  if (!dashSheet || !chilledSheet) return;

  var kgByName = {};
  var dashLastRow = dashSheet.getLastRow();
  for (var dRow = DASH_HEADER_ROW + 1; dRow <= dashLastRow; dRow++) {
    var internalName = dashSheet.getRange(dRow, DASH_NAME_COL).getValue();
    var kgValue = dashSheet.getRange(dRow, DASH_KG_COL).getValue();
    var key = normalizeName(internalName);
    if (!key) continue;
    kgByName[key] = kgValue; // later rows override earlier duplicates
  }

  var chilledLastRow = chilledSheet.getLastRow();
  for (var cRow = CHILLED_HEADER_ROW + 1; cRow <= chilledLastRow; cRow++) {
    var name = chilledSheet.getRange(cRow, CHILLED_NAME_COL).getValue();
    var key2 = normalizeName(name);
    var value = Object.prototype.hasOwnProperty.call(kgByName, key2) ? kgByName[key2] : 0;
    chilledSheet.getRange(cRow, CHILLED_CONSUMPTION_COL).setValue(value);
  }
}

function normalizeName(value) {
  if (!value && value !== 0) return '';
  return value.toString().trim().toUpperCase().replace(/\s+/g, ' ');
}
