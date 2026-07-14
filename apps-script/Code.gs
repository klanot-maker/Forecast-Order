/**
 * Unit Conversion Calculator
 *
 * Reads the pack size embedded in the item Name (Column C) — e.g. "1 LTR",
 * "500 GM", "2.75KG", "1X16KG", "1Kg X 12 pcs" — multiplies it by the
 * pack count (if any "X n" multiplier is present) and by the ordered
 * Quantity (Column P), then writes the result to Column T:
 *   - Weight units (G/GM/GRAM/KG) are expressed in KG.
 *   - Volume units (LTR/ML/CL) are left in their own unit (no density
 *     assumption is made to convert liters to kilograms).
 *   - If a unit keyword is found but with no number in front of it (e.g.
 *     "GRANA PADANO WEDGE BONI SPA KGS"), the size defaults to 1.
 *   - If no unit keyword at all can be found in the Name, Column T is left
 *     untouched.
 *
 * Runs automatically: onOpen() backfills the whole sheet, onEdit() updates
 * a row whenever its Name or Quantity is changed.
 */

// Set to a specific tab name to restrict the script to one sheet, or leave
// null to run against whichever sheet is active/edited.
var SHEET_NAME = null;

var NAME_COL = 3;    // Column C - Name
var QTY_COL = 16;    // Column P - Quantity
var RESULT_COL = 20; // Column T - Converted Total
var HEADER_ROW = 1;
var RESULT_HEADER = 'Converted Total';

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
    .addItem('Recalculate All (KG Conversion)', 'recalcAll')
    .addToUi();
  recalcAll();
}

function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (SHEET_NAME && sheet.getName() !== SHEET_NAME) return;

  var col = e.range.getColumn();
  var row = e.range.getRow();
  if (row <= HEADER_ROW) return;
  if (col !== NAME_COL && col !== QTY_COL) return;

  recalcRow(sheet, row);
}

function recalcAll() {
  var sheet = SHEET_NAME
    ? SpreadsheetApp.getActive().getSheetByName(SHEET_NAME)
    : SpreadsheetApp.getActiveSheet();
  if (!sheet) return;

  sheet.getRange(HEADER_ROW, RESULT_COL).setValue(RESULT_HEADER);

  var lastRow = sheet.getLastRow();
  for (var row = HEADER_ROW + 1; row <= lastRow; row++) {
    recalcRow(sheet, row);
  }
}

function recalcRow(sheet, row) {
  var name = sheet.getRange(row, NAME_COL).getValue();
  var qty = sheet.getRange(row, QTY_COL).getValue();

  if (!name || typeof name !== 'string' || qty === '' || isNaN(qty)) {
    return; // nothing usable to calculate; leave Column T as-is
  }

  var parsed = parseSizeAndUnit(name);
  if (!parsed) {
    return; // no size/unit found in the name; leave Column T as-is
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
  sheet.getRange(row, RESULT_COL).setValue(value + ' ' + unitLabel);
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
