/**
 * Renovation budget sample: Google Sheets setup and helpers.
 * Author: reypage-tech
 */

function computeVariance_(estimate, actual) {
  var e = parseNumber_(estimate);
  var a = parseNumber_(actual);
  return a - e;
}

function parseNumber_(value) {
  if (value === null || value === undefined || value === '') {
    return 0;
  }
  var n = Number(value);
  return isNaN(n) ? 0 : n;
}

function monthlyTotals_(paymentRows) {
  var totals = {};
  if (!paymentRows || paymentRows.length === 0) {
    return totals;
  }
  for (var i = 0; i < paymentRows.length; i++) {
    var row = paymentRows[i];
    if (!row) {
      continue;
    }
    var dateVal = row.date !== undefined ? row.date : row[0];
    var amountVal = row.amount !== undefined ? row.amount : row[1];
    if (isHeaderRow_(dateVal, amountVal)) {
      continue;
    }
    var amount = parseNumber_(amountVal);
    if (amount === 0 && (amountVal === null || amountVal === '')) {
      continue;
    }
    var key = monthKey_(dateVal);
    if (!key) {
      continue;
    }
    totals[key] = (totals[key] || 0) + amount;
  }
  return totals;
}

function isHeaderRow_(dateVal, amountVal) {
  if (typeof dateVal === 'string' && String(dateVal).toLowerCase() === 'date') {
    return true;
  }
  if (typeof amountVal === 'string' && String(amountVal).toLowerCase() === 'amount') {
    return true;
  }
  return false;
}

function monthKey_(dateVal) {
  if (!dateVal && dateVal !== 0) {
    return null;
  }
  if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
    return formatMonth_(dateVal);
  }
  var s = String(dateVal);
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    return s.substring(0, 7);
  }
  var parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    return formatMonth_(parsed);
  }
  return null;
}

function formatMonth_(d) {
  var y = d.getFullYear();
  var m = d.getMonth() + 1;
  return y + '-' + (m < 10 ? '0' + m : '' + m);
}

function overrunLevel_(estimate, actual) {
  var e = parseNumber_(estimate);
  var a = parseNumber_(actual);
  if (e <= 0) {
    return a > 0 ? 2 : 0;
  }
  if (a > e) {
    return 2;
  }
  if (a > e * 0.9) {
    return 1;
  }
  return 0;
}

function lowestBid_(amounts) {
  if (!amounts || amounts.length === 0) {
    return null;
  }
  var nums = [];
  for (var i = 0; i < amounts.length; i++) {
    if (amounts[i] === null || amounts[i] === undefined || amounts[i] === '') {
      continue;
    }
    var n = Number(amounts[i]);
    if (!isNaN(n)) {
      nums.push(n);
    }
  }
  if (nums.length === 0) {
    return null;
  }
  return Math.min.apply(null, nums);
}

if (typeof module !== 'undefined') {
  module.exports = {
    computeVariance_: computeVariance_,
    monthlyTotals_: monthlyTotals_,
    overrunLevel_: overrunLevel_,
    lowestBid_: lowestBid_,
  };
}

var SHEET_DASH = 'Executive Dashboard';
var SHEET_BUDGET = 'Line-Item Budget & Actuals';
var SHEET_BIDS = 'Contractor Bid Leveling';
var SHEET_CO = 'Contingency & Change Order Log';
var SHEET_SCHEDULE = 'Draws & Milestone Payments';
var SHEET_LISTS = 'Lists';
var INPUT_FILL = '#DDEBF7';

function setupWorkbook() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = loadSampleData_();
  var sheetNames = [
    SHEET_DASH,
    SHEET_BUDGET,
    SHEET_BIDS,
    SHEET_CO,
    SHEET_SCHEDULE,
    SHEET_LISTS,
  ];
  sheetNames.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh) {
      ss.deleteSheet(sh);
    }
    ss.insertSheet(name);
  });
  buildLists_(ss.getSheetByName(SHEET_LISTS), data);
  ss.getSheetByName(SHEET_LISTS).hideSheet();
  var budgetMeta = buildBudget_(ss.getSheetByName(SHEET_BUDGET), data);
  buildBids_(ss.getSheetByName(SHEET_BIDS), data);
  buildChangeOrders_(ss.getSheetByName(SHEET_CO), data);
  var scheduleMeta = buildSchedule_(
    ss.getSheetByName(SHEET_SCHEDULE),
    data,
    budgetMeta.totalRow
  );
  buildDashboard_(
    ss.getSheetByName(SHEET_DASH),
    data,
    budgetMeta.lastRow,
    budgetMeta.totalRow,
    scheduleMeta
  );
  sheetNames.forEach(function (name) {
    protectSheet_(ss.getSheetByName(name));
  });
  SpreadsheetApp.getUi().alert('Renovation budget sample workbook is ready.');
}

// SAMPLE_DATA_JSON_START
var SAMPLE_DATA_JSON = "{\"project\": {\"name\": \"Maple Street remodel (sample)\", \"contingencyRate\": 0.1}, \"lists\": {\"rooms\": [\"Kitchen\", \"Primary bath\", \"Living room\", \"Exterior\", \"Whole house\"], \"trades\": [\"General\", \"Plumbing\", \"Electrical\", \"Cabinetry\", \"Flooring\", \"Roofing\"], \"statuses\": [\"Open\", \"Approved\", \"Paid\", \"Void\"], \"vendors\": [\"North Ridge Builders\", \"Summit Trade Co\", \"Harbor Line LLC\", \"Valley Supply Group\", \"Lakeview Mechanical\"]}, \"budget\": [{\"room\": \"Kitchen\", \"trade\": \"Cabinetry\", \"lineItem\": \"Custom base and wall cabinets\", \"estimate\": 14200, \"actual\": 13850}, {\"room\": \"Kitchen\", \"trade\": \"Plumbing\", \"lineItem\": \"Sink, faucet, and rough-in\", \"estimate\": 4200, \"actual\": 4450}, {\"room\": \"Kitchen\", \"trade\": \"Electrical\", \"lineItem\": \"Circuits and under-cabinet lighting\", \"estimate\": 3100, \"actual\": 2950}, {\"room\": \"Kitchen\", \"trade\": \"Flooring\", \"lineItem\": \"Tile floor and transitions\", \"estimate\": 5800, \"actual\": 6100}, {\"room\": \"Primary bath\", \"trade\": \"Plumbing\", \"lineItem\": \"Shower, vanity, and trim\", \"estimate\": 8900, \"actual\": 9200}, {\"room\": \"Primary bath\", \"trade\": \"General\", \"lineItem\": \"Waterproofing and tile labor\", \"estimate\": 7600, \"actual\": 7400}, {\"room\": \"Primary bath\", \"trade\": \"Electrical\", \"lineItem\": \"Fan, heat, and GFCI updates\", \"estimate\": 2400, \"actual\": 2380}, {\"room\": \"Living room\", \"trade\": \"General\", \"lineItem\": \"Drywall patch and paint\", \"estimate\": 5200, \"actual\": 4800}, {\"room\": \"Living room\", \"trade\": \"Flooring\", \"lineItem\": \"Engineered hardwood refinish zone\", \"estimate\": 3900, \"actual\": 4100}, {\"room\": \"Exterior\", \"trade\": \"Roofing\", \"lineItem\": \"Shingle replacement (south slope)\", \"estimate\": 11200, \"actual\": 11200}, {\"room\": \"Exterior\", \"trade\": \"General\", \"lineItem\": \"Gutter and downspout reset\", \"estimate\": 2800, \"actual\": 2650}, {\"room\": \"Whole house\", \"trade\": \"Electrical\", \"lineItem\": \"Panel upgrade and labeling\", \"estimate\": 6800, \"actual\": 7050}, {\"room\": \"Whole house\", \"trade\": \"General\", \"lineItem\": \"Permits, dumpsters, and site protection\", \"estimate\": 4500, \"actual\": 4500}], \"bids\": [{\"trade\": \"Plumbing\", \"bidderA\": \"Summit Trade Co\", \"amountA\": 12800, \"bidderB\": \"Lakeview Mechanical\", \"amountB\": 13150, \"bidderC\": \"Harbor Line LLC\", \"amountC\": 12450}, {\"trade\": \"Electrical\", \"bidderA\": \"Harbor Line LLC\", \"amountA\": 11200, \"bidderB\": \"Valley Supply Group\", \"amountB\": 10950, \"bidderC\": \"Summit Trade Co\", \"amountC\": 11400}, {\"trade\": \"Cabinetry\", \"bidderA\": \"North Ridge Builders\", \"amountA\": 15100, \"bidderB\": \"Valley Supply Group\", \"amountB\": 14200, \"bidderC\": \"Summit Trade Co\", \"amountC\": 14850}, {\"trade\": \"Flooring\", \"bidderA\": \"Summit Trade Co\", \"amountA\": 9800, \"bidderB\": \"North Ridge Builders\", \"amountB\": 9400, \"bidderC\": \"Harbor Line LLC\", \"amountC\": 9650}, {\"trade\": \"Roofing\", \"bidderA\": \"North Ridge Builders\", \"amountA\": 11800, \"bidderB\": \"Harbor Line LLC\", \"amountB\": 11200, \"bidderC\": \"Valley Supply Group\", \"amountC\": 12100}, {\"trade\": \"General\", \"bidderA\": \"North Ridge Builders\", \"amountA\": 22400, \"bidderB\": \"Summit Trade Co\", \"amountB\": 23100, \"bidderC\": \"Harbor Line LLC\", \"amountC\": 21850}], \"changeOrders\": [{\"id\": \"CO-001\", \"date\": \"2026-03-12\", \"room\": \"Kitchen\", \"trade\": \"Cabinetry\", \"description\": \"Upgrade pull-out pantry hardware\", \"amount\": 850, \"status\": \"Approved\"}, {\"id\": \"CO-002\", \"date\": \"2026-04-02\", \"room\": \"Primary bath\", \"trade\": \"Plumbing\", \"description\": \"Add handheld shower set\", \"amount\": 420, \"status\": \"Approved\"}, {\"id\": \"CO-003\", \"date\": \"2026-05-18\", \"room\": \"Whole house\", \"trade\": \"Electrical\", \"description\": \"Two additional recessed cans in hall\", \"amount\": 680, \"status\": \"Open\"}, {\"id\": \"CO-004\", \"date\": \"2026-02-20\", \"room\": \"Living room\", \"trade\": \"General\", \"description\": \"Skim coat extra wall section\", \"amount\": 0, \"status\": \"Void\"}], \"draws\": [{\"draw\": 1, \"milestone\": \"Mobilization and permits\", \"pctOfBudget\": 0.15, \"status\": \"Paid\"}, {\"draw\": 2, \"milestone\": \"Rough plumbing and electrical\", \"pctOfBudget\": 0.25, \"status\": \"Paid\"}, {\"draw\": 3, \"milestone\": \"Kitchen and bath finishes\", \"pctOfBudget\": 0.3, \"status\": \"Approved\"}, {\"draw\": 4, \"milestone\": \"Exterior and final punch\", \"pctOfBudget\": 0.2, \"status\": \"Open\"}, {\"draw\": 5, \"milestone\": \"Retainage release\", \"pctOfBudget\": 0.1, \"status\": \"Open\"}], \"payments\": [{\"date\": \"2026-02-05\", \"vendor\": \"North Ridge Builders\", \"draw\": 1, \"room\": \"Whole house\", \"trade\": \"General\", \"amount\": 12000, \"memo\": \"Draw 1 mobilization\"}, {\"date\": \"2026-02-28\", \"vendor\": \"Lakeview Mechanical\", \"draw\": 2, \"room\": \"Kitchen\", \"trade\": \"Plumbing\", \"amount\": 3200, \"memo\": \"Kitchen rough plumbing\"}, {\"date\": \"2026-03-10\", \"vendor\": \"Harbor Line LLC\", \"draw\": 2, \"room\": \"Whole house\", \"trade\": \"Electrical\", \"amount\": 4500, \"memo\": \"Panel and rough circuits\"}, {\"date\": \"2026-03-22\", \"vendor\": \"Summit Trade Co\", \"draw\": 2, \"room\": \"Primary bath\", \"trade\": \"Plumbing\", \"amount\": 4100, \"memo\": \"Bath rough plumbing\"}, {\"date\": \"2026-04-08\", \"vendor\": \"Valley Supply Group\", \"draw\": 3, \"room\": \"Kitchen\", \"trade\": \"Cabinetry\", \"amount\": 9000, \"memo\": \"Cabinet deposit and install\"}, {\"date\": \"2026-04-25\", \"vendor\": \"Summit Trade Co\", \"draw\": 3, \"room\": \"Kitchen\", \"trade\": \"Flooring\", \"amount\": 3500, \"memo\": \"Kitchen tile install\"}, {\"date\": \"2026-05-06\", \"vendor\": \"North Ridge Builders\", \"draw\": 3, \"room\": \"Primary bath\", \"trade\": \"General\", \"amount\": 5200, \"memo\": \"Bath tile and waterproofing\"}, {\"date\": \"2026-05-20\", \"vendor\": \"Harbor Line LLC\", \"draw\": 3, \"room\": \"Living room\", \"trade\": \"General\", \"amount\": 2800, \"memo\": \"Living room paint phase\"}, {\"date\": \"2026-06-02\", \"vendor\": \"North Ridge Builders\", \"draw\": 4, \"room\": \"Exterior\", \"trade\": \"Roofing\", \"amount\": 8000, \"memo\": \"Roofing progress payment\"}]}";
// SAMPLE_DATA_JSON_END

function loadSampleData_() {
  return JSON.parse(SAMPLE_DATA_JSON);
}

function buildLists_(sheet, data) {
  var headers = ['Rooms', 'Trades', 'Statuses', 'Vendors'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  var lists = data.lists;
  var keys = ['rooms', 'trades', 'statuses', 'vendors'];
  var maxLen = Math.max.apply(null, keys.map(function (k) { return lists[k].length; }));
  var rows = [];
  for (var i = 0; i < maxLen; i++) {
    rows.push(keys.map(function (k) {
      return i < lists[k].length ? lists[k][i] : '';
    }));
  }
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

function buildBudget_(sheet, data) {
  var headers = ['Room', 'Trade', 'Line item', 'Estimate', 'Actual', 'Variance', 'Notes'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  var items = data.budget;
  var values = items.map(function (item) {
    return [item.room, item.trade, item.lineItem, item.estimate, item.actual, '', ''];
  });
  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  var lastRow = 1 + values.length;
  var totalRow = lastRow + 1;
  sheet.getRange(totalRow, 3).setValue('Total').setFontWeight('bold');
  sheet.getRange(totalRow, 4).setFormula('=SUM(D2:D' + lastRow + ')');
  sheet.getRange(totalRow, 5).setFormula('=SUM(E2:E' + lastRow + ')');
  sheet.getRange(totalRow, 6).setFormula('=E' + totalRow + '-D' + totalRow);
  for (var r = 2; r <= lastRow; r++) {
    sheet.getRange(r, 6).setFormula('=E' + r + '-D' + r);
  }
  styleInputRange_(sheet, 2, 1, values.length, 5);
  styleInputRange_(sheet, 2, 7, values.length, 1);
  return { lastRow: lastRow, totalRow: totalRow };
}

function buildBids_(sheet, data) {
  var headers = ['Trade', 'Bidder A', 'Amount A', 'Bidder B', 'Amount B', 'Bidder C', 'Amount C', 'Low bid', 'Low bidder'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  var items = data.bids;
  var startRow = 2;
  items.forEach(function (item, idx) {
    var row = startRow + idx;
    sheet.getRange(row, 1, 1, 7).setValues([[
      item.trade, item.bidderA, item.amountA, item.bidderB, item.amountB, item.bidderC, item.amountC,
    ]]);
    sheet.getRange(row, 8).setFormula('=MIN(C' + row + ',E' + row + ',G' + row + ')');
    sheet.getRange(row, 9).setFormula(
      '=IF(H' + row + '=C' + row + ',B' + row + ',IF(H' + row + '=E' + row + ',D' + row + ',F' + row + '))'
    );
  });
  styleInputRange_(sheet, 2, 1, items.length, 7);
}

function buildChangeOrders_(sheet, data) {
  var headers = ['CO ID', 'Date', 'Room', 'Trade', 'Description', 'Amount', 'Status'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  var rows = data.changeOrders.map(function (co) {
    return [co.id, co.date, co.room, co.trade, co.description, co.amount, co.status];
  });
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  styleInputRange_(sheet, 2, 1, rows.length, headers.length);
}

function scheduleLayout_(drawCount, paymentCount) {
  var drawLast = 1 + drawCount;
  return {
    drawLast: drawLast,
    labelRow: drawLast + 2,
    headerRow: drawLast + 3,
    dataStart: drawLast + 4,
    dataLast: drawLast + 3 + paymentCount,
    totalRow: drawLast + 4 + paymentCount,
  };
}

function buildSchedule_(sheet, data, budgetTotalRow) {
  var budgetRef = "'" + SHEET_BUDGET + "'!D$" + budgetTotalRow;
  var drawHeaders = ['Draw', 'Milestone', 'Pct of budget', 'Scheduled amount', 'Status'];
  sheet.getRange(1, 1, 1, drawHeaders.length).setValues([drawHeaders]).setFontWeight('bold');
  data.draws.forEach(function (d, idx) {
    var row = idx + 2;
    sheet.getRange(row, 1, 1, 3).setValues([[d.draw, d.milestone, d.pctOfBudget]]);
    sheet.getRange(row, 4).setFormula('=ROUND(C' + row + '*' + budgetRef + ',0)');
    sheet.getRange(row, 5).setValue(d.status);
  });
  var layout = scheduleLayout_(data.draws.length, data.payments.length);
  sheet.getRange(layout.labelRow, 1).setValue('Payment log').setFontWeight('bold');
  var payHeaders = ['Date', 'Vendor', 'Draw', 'Room', 'Trade', 'Amount', 'Memo'];
  sheet.getRange(layout.headerRow, 1, 1, payHeaders.length).setValues([payHeaders]).setFontWeight('bold');
  var payRows = data.payments.map(function (p) {
    return [p.date, p.vendor, p.draw, p.room, p.trade, p.amount, p.memo];
  });
  sheet.getRange(layout.dataStart, 1, payRows.length, payHeaders.length).setValues(payRows);
  sheet.getRange(layout.totalRow, 5).setValue('Total paid').setFontWeight('bold');
  sheet.getRange(layout.totalRow, 6).setFormula(
    '=SUM(F' + layout.dataStart + ':F' + layout.dataLast + ')'
  );
  styleInputRange_(sheet, 2, 1, data.draws.length, 3);
  styleInputRange_(sheet, 2, 5, data.draws.length, 1);
  styleInputRange_(sheet, layout.dataStart, 1, payRows.length, payHeaders.length);
  return layout;
}

function monthlySpendFormula_(scheduleMeta, monthCellA1) {
  var sq = "'" + SHEET_SCHEDULE + "'!";
  var start = scheduleMeta.dataStart;
  var last = scheduleMeta.dataLast;
  return (
    '=SUMIFS(' + sq + '$F$' + start + ':$F$' + last + ',' + sq + '$A$' + start + ':$A$' + last +
    ',">="&DATEVALUE(' + monthCellA1 + '&"-01"),' + sq + '$A$' + start + ':$A$' + last +
    ',"<"&EDATE(DATEVALUE(' + monthCellA1 + '&"-01"),1))'
  );
}

function buildDashboard_(sheet, data, budgetLast, budgetTotalRow, scheduleMeta) {
  var bq = "'" + SHEET_BUDGET + "'!";
  var co = "'" + SHEET_CO + "'!";
  var sq = "'" + SHEET_SCHEDULE + "'!";
  sheet.getRange('A1').setValue('Renovation budget dashboard').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A2').setValue('Project');
  sheet.getRange('B2').setValue(data.project.name);
  sheet.getRange('A3').setValue('Author');
  sheet.getRange('B3').setValue('reypage-tech');
  var labels = [
    ['Total estimate', '=' + bq + 'D' + budgetTotalRow],
    ['Total actual', '=' + bq + 'E' + budgetTotalRow],
    ['Variance (actual - estimate)', '=B6-B5'],
    ['Contingency rate (editable)', data.project.contingencyRate],
    ['Contingency amount', '=ROUND(B5*B8,0)'],
    ['Budget plus contingency', '=B5+B9'],
    ['Approved change orders', '=SUMIFS(' + co + '$F$2:$F$20,' + co + '$G$2:$G$20,"Approved")'],
    [
      'Total paid to date',
      '=SUM(' + sq + 'F' + scheduleMeta.dataStart + ':F' + scheduleMeta.dataLast + ')',
    ],
    ['Remaining (budget+contingency - paid)', '=B10-B12'],
  ];
  labels.forEach(function (pair, idx) {
    var row = 5 + idx;
    sheet.getRange(row, 1).setValue(pair[0]).setFontWeight('bold');
    if (String(pair[1]).charAt(0) === '=') {
      sheet.getRange(row, 2).setFormula(pair[1]);
    } else {
      sheet.getRange(row, 2).setValue(pair[1]);
    }
  });
  sheet.getRange('B8').setNumberFormat('0%');
  var months = ['2026-02', '2026-03', '2026-04', '2026-05', '2026-06'];
  sheet.getRange('J5').setValue('Monthly spend').setFontWeight('bold');
  sheet.getRange(6, 10, 1, 2).setValues([['Month', 'Spend']]).setFontWeight('bold');
  months.forEach(function (m, idx) {
    var row = 7 + idx;
    sheet.getRange(row, 10).setValue(m);
    sheet.getRange(row, 11).setFormula(monthlySpendFormula_(scheduleMeta, 'J' + row));
  });
  styleInputRange_(sheet, 2, 2, 1, 1);
  styleInputRange_(sheet, 8, 2, 1, 1);
}

function styleInputRange_(sheet, startRow, startCol, numRows, numCols) {
  if (numRows <= 0 || numCols <= 0) {
    return;
  }
  sheet.getRange(startRow, startCol, numRows, numCols).setBackground(INPUT_FILL);
}

function protectSheet_(sheet) {
  var name = sheet.getName();
  var maxRows = Math.max(sheet.getLastRow() + 5, 40);
  var maxCols = 12;
  sheet.getRange(1, 1, maxRows, maxCols).setLocked(true);
  if (name === SHEET_BUDGET) {
    sheet.getRange('A2:G100').setLocked(false);
    sheet.getRange('F2:F100').setLocked(true);
    var totalRow = budgetDataLastRow_(sheet) + 1;
    sheet.getRange(totalRow, 4, 1, 3).setLocked(true);
  } else if (name === SHEET_DASH) {
    sheet.getRange('B2').setLocked(false);
    sheet.getRange('B8').setLocked(false);
  } else if (name === SHEET_CO || name === SHEET_BIDS) {
    sheet.getRange('A2:Z200').setLocked(false);
  } else if (name === SHEET_SCHEDULE) {
    sheet.getRange('A2:E200').setLocked(false);
    sheet.getRange('D2:D200').setLocked(true);
    sheet.getRange('A10:G200').setLocked(false);
  }
  sheet.protect();
}

function budgetDataLastRow_(sheet) {
  var lastRow = sheet.getLastRow();
  while (lastRow > 1 && sheet.getRange(lastRow, 3).getValue() === 'Total') {
    lastRow--;
  }
  return lastRow;
}

function highlightOverruns_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BUDGET);
  if (!sheet) {
    return;
  }
  var lastRow = budgetDataLastRow_(sheet);
  if (lastRow < 2) {
    return;
  }
  var numRows = lastRow - 1;
  var estimates = sheet.getRange(2, 4, numRows, 1).getValues();
  var actuals = sheet.getRange(2, 5, numRows, 1).getValues();
  var backgrounds = [];
  for (var i = 0; i < estimates.length; i++) {
    var level = overrunLevel_(estimates[i][0], actuals[i][0]);
    var color = '#ffffff';
    if (level === 2) {
      color = '#f4cccc';
    } else if (level === 1) {
      color = '#fce5cd';
    }
    backgrounds.push([color, color, color, color, color, color, color]);
  }
  sheet.getRange(2, 1, numRows, 7).setBackgrounds(backgrounds);
}

function onEdit(e) {
  highlightOverruns_();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Renovation Budget')
    .addItem('Setup sample workbook', 'setupWorkbook')
    .addItem('Highlight budget overruns', 'highlightOverruns_')
    .addToUi();
}
