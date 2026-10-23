/**
 * ============================================================================
 *  uPVC FMS - Google Apps Script Backend (Web.js)
 *  Production-Ready: Sheet-First 2-Way Sync + Multi-File Revisions on All Slots
 * ============================================================================
 */

const APP_NAME = 'uPVC FMS';
const APP_SUBTITLE = 'Quotation & Procurement Portal';
const SPREADSHEET_ID = '';
const DRIVE_ROOT_FOLDER_NAME = 'uPVC FMS Documents';
const SHARE_UPLOADS_WITH_LINK = true;

const ID_PREFIX = { QUOTATION: 'QT-UPVC-', PO: 'PO-UPVC-' };
const ID_PAD = 4;
const CURRENCY = '\u20B9';

const SHEET = {
  QUOTATION_MASTER: 'uPVC_Quotation_Master',
  PO_MASTER: 'uPVC_Purchase_Order_Master',
  CONFIG_DROPDOWNS: 'uPVC_Config_Dropdowns'
};

const CONFIG_HEADERS = [
  'Col A: Vendors',
  'Col B: Variant',
  'Col C: Mesh / Jaali',
  'Col D: Glass Types',
  'Col E: Loss Reasons'
];

const CONFIG_COLUMN_INDEX = {
  vendors: 1,
  variants: 2,
  meshTypes: 3,
  glassTypes: 4,
  lossReasons: 5
};

const CONFIG_DEFAULTS = {
  vendors: ['NCL', 'Grayzo', 'Poonam Alu.', 'WinFrame', 'Magnesta', 'Alumil', 'Prominance'],
  variants: ['Fix-Light', '2-Track Sliding', '2.5-Track Sliding', '3-Track Sliding',
    'Casement / Openable', 'Sliding Folding', 'Ventilator', 'Door Frames'],
  meshTypes: ['None', 'Nylon Mesh', 'SS 304 Mesh', 'Pleated Mesh', 'Fiberglass Mesh'],
  glassTypes: ['5mm, Clear, Toughened', '6mm, Clear, Toughened', '10mm, Clear, Toughened',
    'Toughened / Tempered', 'Clear Float', 'Tinted Glass', 'Frosted / Obscure', 'DGU (Double Glazed)'],
  lossReasons: ['Price Too High', 'Lead Time Too Long', 'Competitor Selected',
    'Client Postponed', 'Site Not Feasible', 'Specification Mismatch']
};

/* ==========================================================================
 * COLUMN MAP
 * ======================================================================== */

const COLUMN_MAP = {};

COLUMN_MAP[SHEET.QUOTATION_MASTER] = [
  ['Quotation ID', 'id', 'text'],
  ['Client / Site Name', 'site', 'text'],
  ['Date', 'date', 'date'],
  ['Location', 'location', 'text'],
  ['Vendor Name', 'vendor', 'text'],
  ['Variant', 'variant', 'text'],
  ['Mesh / Jaali', 'mesh', 'text'],
  ['Glass Specs', 'glass', 'text'],
  ['Total Item Qty', 'qty', 'int'],
  ['Lead Time', 'leadTime', 'int'],
  ['Rough Sizes', 'roughSizes', 'doc'],
  ['Vendor Quote', 'vendorQuote', 'doc'],
  ['Vendor Amount', 'vendorAmount', 'num'],
  ['Our Quote', 'ourQuote', 'doc'],
  ['Our Amount', 'ourAmount', 'num']
];

COLUMN_MAP[SHEET.PO_MASTER] = [
  ['PO ID', 'id', 'text'],
  ['Quotation ID', 'quotationId', 'text'],
  ['Client / Site Name', 'site', 'text'],
  ['PO Date', 'date', 'date'],
  ['Location', 'location', 'text'],
  ['Vendor Name', 'vendor', 'text'],
  ['Variant', 'variant', 'text'],
  ['Mesh / Jaali', 'mesh', 'text'],
  ['Glass Specs', 'glass', 'text'],
  ['Total Item Qty', 'qty', 'int'],
  ['Lead Time (Days)', 'leadTime', 'int'],
  ['Rough Sizes', 'roughSizes', 'doc'],
  ['Vendor Quote', 'vendorQuote', 'doc'],
  ['Vendor Amount', 'vendorAmount', 'num'],
  ['Our Quote', 'ourQuote', 'doc'],
  ['Our Amount', 'ourAmount', 'num'],
  ['PO Order', 'poOrder', 'doc'],
  ['Bills', 'bills', 'bills'],
  ['Bills AMT', 'billsAmt', 'num']
];

const HEADERS = (function () {
  var out = {};
  Object.keys(COLUMN_MAP).forEach(function (sheetName) {
    out[sheetName] = COLUMN_MAP[sheetName].map(function (spec) { return spec[0]; });
  });
  return out;
})();

const HEADER_ALIASES = {
  'quotation id': ['quote id', 'quotationid', 'quote no', 'quotation no', 'quote ref', 'qt id'],
  'client / site name': ['client/site name', 'client name', 'client', 'site', 'site name', 'client / site', 'party name',
    'client name / site name', 'client name / site'],
  'date': ['quote date', 'quotation date', 'timestamp'],
  'location': ['site location', 'address', 'site address'],
  'vendor name': ['vendor', 'supplier', 'supplier name', 'party', 'vendor/supplier', 'vendor name:'],
  'variant': ['variants', 'system variant', 'system variants', 'product type', 'type', 'orientation'],
  'mesh / jaali': ['mesh', 'jaali', 'mesh/jaali', 'mesh jaali', 'which jaali ?', 'mesh type'],
  'glass specs': ['glass', 'glass spec', 'glass specification', 'glass thickness'],
  'total item qty': ['qty', 'quantity', 'total qty', 'total item quantity', 'item qty', 'total items'],
  'lead time': ['leadtime', 'lead time (days)', 'lead time days', 'delivery time'],
  'lead time (days)': ['lead time', 'leadtime', 'lead time days', 'lead time(days)', 'delivery time'],
  'rough sizes': ['rough size', 'roughsizes', 'sizing', 'sizes', 'sizing schedule', 'size pdf'],
  'vendor quote': ['vendor quotation', 'supplier quote', 'vendor quote doc', 'quotation from ncl(pdf)'],
  'vendor amount': ['vendor amt', 'vendoramt', 'vendor price', 'vendor value', 'ncl quotation amount'],
  'our quote': ['our quotation', 'client quote', 'our quote doc', 'co. quote', 'current company quotation'],
  'our amount': ['our amt', 'ouramount', 'our price', 'our value', 'quoted amount', 'quote amount', 'client amount'],
  'po id': ['po no', 'po number', 'purchase order id', 'poid', 'po ref'],
  'po date': ['purchase order date', 'order date', 'po dt', 'timestamp'],
  'po order': ['poorder', 'po copy', 'po document', 'po order copy'],
  'bills': ['bill', 'invoice', 'invoices', 'tax invoice', 'bills/invoices', 'ncl billing uploads'],
  'bills amt': ['bills amount', 'bill amount', 'billsamt', 'bills total', 'bill amt', 'ncl billing amount']
};

/* ==========================================================================
 * WEB APP & UNIFIED MENU ENTRY POINT
 * ======================================================================== */

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP_NAME + ' \u2014 ' + APP_SUBTITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Master onOpen trigger: loads all 3 menus (uPVC FMS, Archiving Tools, BMP Formulas)
 * without any function name collisions across files.
 */
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('uPVC FMS')
      .addItem('\uD83D\uDDA5\uFE0F Open Web App Portal', 'showWebAppDialog')
      .addSeparator()
      .addItem('\uD83D\uDD27 Initialize / Repair Database', 'initDatabase')
      .addItem('\uD83D\uDCE5 Seed Demo Data (Sample Quotations & POs)', 'seedDemoData')
      .addItem('\uD83D\uDD04 Sync Document Links to FMS Sheets (uPVC PO & Quotation)', 'syncAllDocLinksToAllSheets')
      .addItem('\uD83D\uDD17 Convert All Document Cells to Clickable Links', 'migrateAllDocsToRichText')
      .addItem('\uD83C\uDD94 Repair Legacy Data (Backfill IDs)', 'repairLegacyIds')
      .addToUi();

    if (typeof buildArchivingMenu_ === 'function') buildArchivingMenu_(ui);
    if (typeof buildBmpMenu_ === 'function') buildBmpMenu_(ui);
  } catch (err) {
    Logger.log('onOpen menu skipped: ' + err);
  }
}

function showWebAppDialog() {
  var html = HtmlService.createTemplateFromFile('Index').evaluate()
    .setWidth(1500)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, APP_NAME + ' \u2014 ' + APP_SUBTITLE);
}

/* ==========================================================================
 * SCHEMA & VALIDATION BOOTSTRAP
 * ======================================================================== */

function getSpreadsheet_() {
  if (SPREADSHEET_ID && String(SPREADSHEET_ID).trim()) {
    return SpreadsheetApp.openById(String(SPREADSHEET_ID).trim());
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet found.');
  return ss;
}

function getSheet_(sheetName, createIfMissing) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    if (createIfMissing === false) {
      throw new Error('Sheet "' + sheetName + '" not found. Run uPVC FMS > Initialize / Repair Database.');
    }
    sheet = createSheetWithSchema_(ss, sheetName);
  }
  return sheet;
}

function createSheetWithSchema_(ss, sheetName) {
  var sheet = ss.insertSheet(sheetName);
  if (sheetName === SHEET.CONFIG_DROPDOWNS) {
    sheet.getRange(1, 1, 1, CONFIG_HEADERS.length).setValues([CONFIG_HEADERS]);
    Object.keys(CONFIG_COLUMN_INDEX).forEach(function (key) {
      var values = CONFIG_DEFAULTS[key] || [];
      if (values.length) {
        sheet.getRange(2, CONFIG_COLUMN_INDEX[key], values.length, 1)
          .setValues(values.map(function (v) { return [v]; }));
      }
    });
    styleHeader_(sheet, CONFIG_HEADERS.length);
    sheet.setColumnWidths(1, CONFIG_HEADERS.length, 200);
  } else {
    var headers = HEADERS[sheetName];
    if (!headers) throw new Error('No column map defined for sheet: ' + sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader_(sheet, headers.length);
    applyColumnWidths_(sheet, sheetName);
  }
  sheet.setFrozenRows(1);
  return sheet;
}

function styleHeader_(sheet, width) {
  sheet.getRange(1, 1, 1, width)
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontFamily('Roboto Mono')
    .setBackground('#1E40AF')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(1, 36);
}

function applyColumnWidths_(sheet, sheetName) {
  var specs = COLUMN_MAP[sheetName] || [];
  specs.forEach(function (spec, i) {
    var type = spec[2] || 'text';
    var width = 135;
    if (type === 'doc') width = 230;
    else if (type === 'num') width = 120;
    else if (type === 'date') width = 105;
    else if (spec[1] === 'site') width = 220;
    else if (spec[1] === 'location') width = 190;
    sheet.setColumnWidth(i + 1, width);
    if (type === 'doc') {
      sheet.getRange(2, i + 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setWrap(true);
    }
  });
}

function syncSheetColumns_(sheet, headers) {
  var inserted = 0;
  headers.forEach(function (header, index) {
    var last = Math.max(sheet.getLastColumn(), 1);
    var row1 = sheet.getRange(1, 1, 1, last).getValues()[0].map(normHeader_);
    if (headerPresent_(header, row1)) return;

    var position = Math.min(index, last) + 1;
    if (position > sheet.getMaxColumns()) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), position - sheet.getMaxColumns());
    }
    sheet.insertColumnBefore(position);
    sheet.getRange(1, position).setValue(header);
    inserted++;
  });
  return inserted;
}

function initDatabase() {
  var ss = getSpreadsheet_();
  Object.keys(HEADERS).forEach(function (sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      createSheetWithSchema_(ss, sheetName);
      return;
    }
    var headers = HEADERS[sheetName];
    syncSheetColumns_(sheet, headers);
    if (sheet.getLastRow() < 2) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    styleHeader_(sheet, Math.max(sheet.getLastColumn(), headers.length));
    applyColumnWidths_(sheet, sheetName);
    sheet.setFrozenRows(1);
  });

  var cfg = ss.getSheetByName(SHEET.CONFIG_DROPDOWNS);
  if (!cfg) {
    createSheetWithSchema_(ss, SHEET.CONFIG_DROPDOWNS);
  } else if (cfg.getLastRow() < 2) {
    Object.keys(CONFIG_COLUMN_INDEX).forEach(function (key) {
      var values = CONFIG_DEFAULTS[key] || [];
      if (values.length) {
        cfg.getRange(2, CONFIG_COLUMN_INDEX[key], values.length, 1)
          .setValues(values.map(function (v) { return [v]; }));
      }
    });
  }

  applyValidations_();
  SpreadsheetApp.flush();
  return { success: true, sheets: Object.keys(HEADERS), config: getConfigData() };
}

function ensureSchema_() {
  var ss = getSpreadsheet_();
  var missing = false;
  Object.keys(HEADERS).concat([SHEET.CONFIG_DROPDOWNS]).forEach(function (sheetName) {
    if (!ss.getSheetByName(sheetName)) missing = true;
  });
  if (missing) initDatabase();
}

function getValidationSpecs_() {
  var stageSheets = [SHEET.QUOTATION_MASTER, SHEET.PO_MASTER];
  var base = [
    ['Vendor Name', 'vendors'],
    ['Mesh / Jaali', 'meshTypes']
  ];
  var specs = [];
  stageSheets.forEach(function (sheetName) {
    base.forEach(function (pair) {
      if (columnIndexByHeader_(sheetName, pair[0])) {
        specs.push({ sheet: sheetName, header: pair[0], config: pair[1] });
      }
    });
  });
  return specs;
}

function applyValidations_() {
  var ss = getSpreadsheet_();
  var cfg = ss.getSheetByName(SHEET.CONFIG_DROPDOWNS);
  if (!cfg) return;
  if (cfg.getMaxRows() < 2) cfg.insertRowAfter(1);
  var span = Math.max(cfg.getMaxRows() - 1, 1);

  [SHEET.QUOTATION_MASTER, SHEET.PO_MASTER].forEach(function (sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    ['Variant', 'Glass Specs'].forEach(function (header) {
      var col = columnIndexByHeader_(sheetName, header, sheet);
      if (col && sheet.getMaxRows() > 1) {
        sheet.getRange(2, col, Math.max(sheet.getMaxRows() - 1, 1), 1).clearDataValidations();
      }
    });
  });

  getValidationSpecs_().forEach(function (spec) {
    var sheet = ss.getSheetByName(spec.sheet);
    if (!sheet) return;
    if (sheet.getMaxRows() < 2) sheet.insertRowAfter(1);
    var col = columnIndexByHeader_(spec.sheet, spec.header, sheet);
    var sourceCol = CONFIG_COLUMN_INDEX[spec.config];
    if (!col || !sourceCol) return;

    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(cfg.getRange(2, sourceCol, span, 1), true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, col, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
  });
}

/* ==========================================================================
 * HEADER & ROW MAPPING HELPERS
 * ======================================================================== */

function normHeader_(value) {
  return String(value === null || value === undefined ? '' : value)
    .trim().replace(/\s+/g, ' ').toLowerCase();
}

function headerPresent_(header, presentNames) {
  var target = normHeader_(header);
  if (presentNames.indexOf(target) > -1) return true;
  var aliases = HEADER_ALIASES[target] || [];
  for (var i = 0; i < aliases.length; i++) {
    if (presentNames.indexOf(aliases[i]) > -1) return true;
  }
  return false;
}

function actualHeaderMap_(sheet, sheetName) {
  var last = Math.max(sheet.getLastColumn(), 1);
  var row1 = sheet.getRange(1, 1, 1, last).getValues()[0];
  var present = {};
  for (var i = 0; i < row1.length; i++) {
    var name = normHeader_(row1[i]);
    if (name && !Object.prototype.hasOwnProperty.call(present, name)) present[name] = i + 1;
  }
  var map = {};
  headersOf_(sheetName).forEach(function (header) {
    var target = normHeader_(header);
    if (!target) return;
    if (Object.prototype.hasOwnProperty.call(present, target)) { map[header] = present[target]; return; }
    var aliases = HEADER_ALIASES[target] || [];
    for (var a = 0; a < aliases.length; a++) {
      if (Object.prototype.hasOwnProperty.call(present, aliases[a])) { map[header] = present[aliases[a]]; return; }
    }
  });
  return map;
}

function columnIndexMap_(sheetName, sheet) {
  var specs = columnSpecs_(sheetName);
  var headerMap = actualHeaderMap_(sheet, sheetName);
  var last = Math.max(sheet.getLastColumn(), 1);
  var row1 = sheet.getRange(1, 1, 1, last).getValues()[0];
  return specs.map(function (spec, i) {
    var col = headerMap[spec[0]];
    if (col && col > 0) return col - 1;
    var label = normHeader_(row1[i]);
    if (!label || label === normHeader_(spec[0])) return i;
    return -1;
  });
}

function isBlankRow_(values) {
  for (var i = 0; i < values.length; i++) {
    if (String(values[i] === null || values[i] === undefined ? '' : values[i]).trim()) return false;
  }
  return true;
}

function columnSpecs_(sheetName) {
  var specs = COLUMN_MAP[sheetName];
  if (!specs) throw new Error('Unknown sheet in COLUMN_MAP: ' + sheetName);
  return specs;
}

function headersOf_(sheetName) {
  return COLUMN_MAP[sheetName].map(function (spec) { return spec[0]; });
}

function columnSpecByKey_(sheetName, key) {
  var specs = columnSpecs_(sheetName);
  for (var i = 0; i < specs.length; i++) {
    if (specs[i][1] === key) return specs[i];
  }
  return null;
}

function columnSpecByHeader_(sheetName, header) {
  var specs = columnSpecs_(sheetName);
  for (var i = 0; i < specs.length; i++) {
    if (specs[i][0] === header) return specs[i];
  }
  return null;
}

function columnIndexByHeader_(sheetName, header, sheet) {
  if (sheet) {
    var map = actualHeaderMap_(sheet, sheetName);
    if (map[header]) return map[header];
  }
  var specs = COLUMN_MAP[sheetName];
  if (!specs) return 0;
  for (var i = 0; i < specs.length; i++) {
    if (specs[i][0] === header) return i + 1;
  }
  return 0;
}

function scriptZone_() {
  return Session.getScriptTimeZone() || 'Asia/Kolkata';
}

function today_() {
  return Utilities.formatDate(new Date(), scriptZone_(), 'yyyy-MM-dd');
}

function toDateStr_(value) {
  if (value === null || value === undefined || value === '') return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(value, scriptZone_(), 'yyyy-MM-dd');
  }
  var text = String(value).trim();
  var iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[1] + '-' + iso[2] + '-' + iso[3];
  var slash = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (slash) return slash[3] + '-' + pad_(slash[2], 2) + '-' + pad_(slash[1], 2);
  var parsed = new Date(text);
  if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, scriptZone_(), 'yyyy-MM-dd');
  return text;
}

function pad_(value, width) {
  var s = String(value);
  while (s.length < width) s = '0' + s;
  return s;
}

function idSequenceNumber_(raw, prefix) {
  var text = String(raw === null || raw === undefined ? '' : raw).trim().toUpperCase();
  var head = String(prefix).toUpperCase();
  if (text.indexOf(head) !== 0) return -1;
  var digits = text.substring(head.length).replace(/[^0-9]/g, '');
  if (!digits) return -1;
  var n = parseInt(digits, 10);
  return isNaN(n) ? -1 : n;
}

function highestIdNumber_(prefix, sheetOrSheets) {
  var sheets = Array.isArray(sheetOrSheets) ? sheetOrSheets : [sheetOrSheets];
  var ss = getSpreadsheet_();
  var highest = 0;

  sheets.forEach(function (sh) {
    var sheetName = typeof sh === 'string' ? sh : '';
    var targetSheet = sheetName ? ss.getSheetByName(sheetName) : sh;
    if (!targetSheet) return;
    var last = targetSheet.getLastRow();
    if (last < 2) return;

    var idCol = 1;
    if (sheetName && HEADERS[sheetName]) {
      var map = actualHeaderMap_(targetSheet, sheetName);
      var idHeader = headersOf_(sheetName)[0];
      if (map[idHeader]) idCol = map[idHeader];
    }

    var values = targetSheet.getRange(2, idCol, last - 1, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      var n = idSequenceNumber_(values[i][0], prefix);
      if (n > highest) highest = n;
    }
  });
  return highest;
}

function nextId_(prefix, sheetOrSheets) {
  return prefix + pad_(highestIdNumber_(prefix, sheetOrSheets) + 1, ID_PAD);
}

function findRowById_(sheetName, sheet, idValue) {
  var target = String(idValue === null || idValue === undefined ? '' : idValue).trim().toLowerCase();
  if (!target) return -1;
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var idCol = (columnIndexMap_(sheetName, sheet)[0] >= 0) ? (columnIndexMap_(sheetName, sheet)[0] + 1) : 1;
  var values = sheet.getRange(2, idCol, last - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim().toLowerCase() === target) return i + 2;
  }
  return -1;
}

/* ==========================================================================
 * DOCUMENT & SHEET-FIRST 2-WAY SYNC HELPERS
 * ======================================================================== */

function formatShortDate_(value) {
  var d = null;
  if (!value) {
    d = new Date();
  } else if (Object.prototype.toString.call(value) === '[object Date]') {
    d = value;
  } else {
    var text = String(value).trim();
    var iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (iso) {
      d = new Date(parseInt(iso[1], 10), parseInt(iso[2], 10) - 1, parseInt(iso[3], 10));
    } else {
      var slash = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
      if (slash) {
        var yr = parseInt(slash[3], 10);
        if (yr < 100) yr += 2000;
        d = new Date(yr, parseInt(slash[2], 10) - 1, parseInt(slash[1], 10));
      } else {
        d = new Date(text);
      }
    }
  }
  if (!d || isNaN(d.getTime())) d = new Date();
  return d.getDate() + '-' + (d.getMonth() + 1) + '-' + String(d.getFullYear()).slice(-2);
}

function docPrefix_(docKey) {
  var k = String(docKey || '').trim().toLowerCase().replace(/[\s\-_]+/g, '');
  if (k === 'roughsizes' || k === 'roughsize') return 'rough_sizes';
  if (k === 'vendorquote' || k === 'vendorquotes') return 'vendor_quote';
  if (k === 'ourquote' || k === 'ourquotes') return 'our_quote';
  if (k === 'poorder' || k === 'poorders' || k === 'po') return 'po_order';
  if (k === 'bills' || k === 'bill') return 'bill';
  return String(docKey || 'file').toLowerCase();
}

function formatDocLine_(file, index, docKey) {
  var dateVal = file.uploadedAt || file.date;
  var dateStr = formatShortDate_(dateVal);
  var prefix = docPrefix_(docKey);
  return dateStr + ' ' + prefix + '_' + (index + 1);
}

function buildDocRichText_(files, docKey) {
  var list = parseDocs_(files);
  if (!list || !list.length) {
    return SpreadsheetApp.newRichTextValue().setText('').build();
  }
  // Each upload slot holds strictly one file revision for seamless =INDEX() hyperlink compatibility
  var file = list[0];
  var lineText = formatDocLine_(file, 0, docKey);
  var url = file.url || (file.id ? ('https://drive.google.com/file/d/' + file.id + '/view') : '');
  var builder = SpreadsheetApp.newRichTextValue().setText(lineText);
  if (url) {
    builder.setLinkUrl(0, lineText.length, url);
  }
  return builder.build();
}

function writeDocCell_(sheet, rowIndex, col, files, docKey, skipSync) {
  var list = parseDocs_(files);
  var richText = buildDocRichText_(list, docKey);
  var range = sheet.getRange(rowIndex, col);
  var noteStr = list.length ? JSON.stringify(list) : '';
  range.setRichTextValue(richText);
  range.setNote(noteStr);
  range.setWrap(true);

  if (!skipSync) {
    try {
      var sName = sheet.getName();
      if (sName === SHEET.QUOTATION_MASTER || sName === SHEET.PO_MASTER) {
        var idCol = (columnIndexMap_(sName, sheet)[0] >= 0) ? (columnIndexMap_(sName, sheet)[0] + 1) : 1;
        var recordId = String(sheet.getRange(rowIndex, idCol).getValue() || '').trim();
        var siteCol = columnIndexByHeader_(sName, 'Client / Site Name', sheet);
        var siteName = siteCol ? String(sheet.getRange(rowIndex, siteCol).getValue() || '').trim() : '';
        var qIdCol = (sName === SHEET.PO_MASTER) ? columnIndexByHeader_(sName, 'Quotation ID', sheet) : 0;
        var qId = qIdCol ? String(sheet.getRange(rowIndex, qIdCol).getValue() || '').trim() : '';

        syncDocCellToLinkedSheets_(sName, recordId, siteName, qId, docKey, richText, noteStr);
      }
    } catch (syncErr) {
      Logger.log('Linked sheet sync skipped for row ' + rowIndex + ': ' + syncErr);
    }
  }

  return list;
}

// In-memory cache for sheet header/column mappings so we never re-scan workbook tabs multiple times per execution
var _sheetMappingCache = {};

/**
 * Syncs ALL [DOC] columns of a single master row to linked sheets in ONE single pass.
 */
function syncRowDocsToLinkedSheets_(sheet, sheetName, rowIndex) {
  try {
    if (sheetName !== SHEET.QUOTATION_MASTER && sheetName !== SHEET.PO_MASTER) return;

    var idCol = (columnIndexMap_(sheetName, sheet)[0] >= 0) ? (columnIndexMap_(sheetName, sheet)[0] + 1) : 1;
    var recordId = String(sheet.getRange(rowIndex, idCol).getValue() || '').trim();
    var siteCol = columnIndexByHeader_(sheetName, 'Client / Site Name', sheet);
    var siteName = siteCol ? String(sheet.getRange(rowIndex, siteCol).getValue() || '').trim() : '';
    var qIdCol = (sheetName === SHEET.PO_MASTER) ? columnIndexByHeader_(sheetName, 'Quotation ID', sheet) : 0;
    var qId = qIdCol ? String(sheet.getRange(rowIndex, qIdCol).getValue() || '').trim() : '';

    if (!recordId && !siteName) return;

    var docSpecs = columnSpecs_(sheetName).filter(function (s) { return s[2] === 'doc'; });
    var docData = [];
    docSpecs.forEach(function (spec) {
      var col = columnIndexByHeader_(sheetName, spec[0], sheet);
      if (!col) return;
      var range = sheet.getRange(rowIndex, col);
      docData.push({
        docKey: spec[1],
        richText: range.getRichTextValue(),
        note: range.getNote()
      });
    });

    if (!docData.length) return;

    var ss = getSpreadsheet_();
    var allSheets = ss.getSheets();
    var normId = normHeader_(recordId);
    var normSite = normHeader_(siteName);
    var normQId = normHeader_(qId);

    var excludedSheets = [
      sheetName.toLowerCase(),
      SHEET.CONFIG_DROPDOWNS.toLowerCase(),
      'setupconfig'
    ];

    allSheets.forEach(function (dstSheet) {
      var dName = dstSheet.getName();
      if (excludedSheets.indexOf(dName.toLowerCase()) > -1) return;
      if (dName.toLowerCase().indexOf('cancel') > -1 || dName.toLowerCase().indexOf('done') > -1) return;
      if (dName.toLowerCase() === 'summary' || dName.toLowerCase().indexOf('enquiry') > -1) return;

      var lastRow = dstSheet.getLastRow();
      var lastCol = dstSheet.getLastColumn();
      if (lastRow < 2 || lastCol < 1) return;

      var mapping = getSheetDocColumnMapping_(dstSheet, sheetName);
      if (!mapping.siteCol && !mapping.idCol) return;

      var matchingCols = [];
      docData.forEach(function (d) {
        if (mapping.docCols[d.docKey]) {
          matchingCols.push({
            col: mapping.docCols[d.docKey],
            richText: d.richText,
            note: d.note
          });
        }
      });
      if (!matchingCols.length) return;

      var scanStart = mapping.headerRow + 1;
      var numRows = lastRow - scanStart + 1;
      if (numRows < 1) return;

      var siteValues = mapping.siteCol ? dstSheet.getRange(scanStart, mapping.siteCol, numRows, 1).getValues() : null;
      var idValues = mapping.idCol ? dstSheet.getRange(scanStart, mapping.idCol, numRows, 1).getValues() : null;

      for (var i = 0; i < numRows; i++) {
        var isMatch = false;

        if (normSite && siteValues) {
          var cellSite = normHeader_(siteValues[i][0]);
          if (cellSite && cellSite === normSite) isMatch = true;
        }

        if (!isMatch && idValues) {
          var cellId = normHeader_(idValues[i][0]);
          if (cellId && (cellId === normId || (normQId && cellId === normQId))) isMatch = true;
        }

        if (isMatch) {
          var rIdx = scanStart + i;
          matchingCols.forEach(function (mc) {
            var cellRange = dstSheet.getRange(rIdx, mc.col);
            cellRange.setRichTextValue(mc.richText);
            if (mc.note) cellRange.setNote(mc.note);
            cellRange.setWrap(true);
          });
        }
      }
    });
  } catch (e) {
    Logger.log('Batch row doc sync skipped: ' + e);
  }
}

function parseRichTextDocs_(richText) {
  if (!richText) return [];
  var runs = (typeof richText.getRuns === 'function') ? richText.getRuns() : [];
  var out = [];
  var seenUrls = {};
  runs.forEach(function (run) {
    var url = run.getLinkUrl();
    if (!url) return;
    var text = String(run.getText() || '').trim();
    if (!text) return;
    var key = url + '|' + text;
    if (seenUrls[key]) return;
    seenUrls[key] = true;

    var driveId = driveIdFromUrl_(url);
    var dateMatch = text.match(/^(\d{1,4}[-\/]\d{1,2}[-\/]\d{2,4})/);
    var nameMatch = text.match(/([a-zA-Z0-9_\-]+_\d+)/i);
    var cleanName = nameMatch ? nameMatch[1] : (text.replace(/^(\d{1,4}[-\/]\d{1,2}[-\/]\d{2,4})\s+/, '').trim() || 'Attachment');
    out.push({
      id: driveId,
      name: cleanName,
      url: url,
      previewUrl: driveId ? ('https://drive.google.com/file/d/' + driveId + '/preview') : url,
      thumbnailUrl: driveId ? ('https://drive.google.com/thumbnail?id=' + driveId + '&sz=w800') : '',
      type: guessDocType_('', cleanName),
      amount: 0,
      uploadedAt: dateMatch ? dateMatch[1] : ''
    });
  });
  return out;
}

/**
 * Reconciles manual Google Sheet cell edits with the hidden Cell Note JSON.
 * - If the user cleared the cell in Google Sheets, returns [].
 * - If the user pasted new URLs or deleted a link line in Google Sheets,
 *   updates the file list while preserving rich metadata from the Note for matching files.
 */
function reconcileDocCell_(rawVal, richVal, noteStr) {
  var text = String(rawVal === null || rawVal === undefined ? '' : rawVal).trim();
  if (!text) return []; // User cleared the cell in the Google Sheet

  var noteDocs = noteStr ? parseDocs_(noteStr) : [];
  var richDocs = richVal ? parseRichTextDocs_(richVal) : [];
  var textDocs = parseDocs_(text);
  var sheetDocs = richDocs.length ? richDocs : textDocs;

  if (!sheetDocs.length) return noteDocs.slice(0, 1);
  if (!noteDocs.length) return sheetDocs.slice(0, 1);

  var sd = sheetDocs[0];
  var nd = noteDocs[0];
  var isSame = (sd.id && nd.id && String(sd.id) === String(nd.id)) ||
               (sd.url && nd.url && String(sd.url) === String(nd.url));

  return [isSame ? nd : sd];
}

function serializeDocs_(value, docKey) {
  if (value === null || value === undefined || value === '') return '';
  var list = Array.isArray(value) ? value : (typeof value === 'object' ? [value] : []);
  if (list.length === 0) {
    var str = String(value).trim();
    if (str && (str.charAt(0) === '[' || str.charAt(0) === '{')) {
      try {
        return serializeDocs_(JSON.parse(str), docKey);
      } catch (e) { }
    }
    return str;
  }
  return list.map(function (file, index) {
    return formatDocLine_(file, index, docKey);
  }).join('\n');
}

function parseDocs_(value) {
  if (value === null || value === undefined || value === '') return [];
  if (Array.isArray(value)) return value.filter(Boolean);
  if (typeof value === 'object') {
    if (typeof value.getRuns === 'function') {
      var fromRich = parseRichTextDocs_(value);
      if (fromRich.length) return fromRich;
      value = value.getText ? value.getText() : '';
    } else {
      return [value];
    }
  }

  var text = String(value).trim();
  if (!text) return [];

  var first = text.charAt(0);
  if (first === '[' || first === '{') {
    try {
      var parsed = JSON.parse(text);
      var arr = Array.isArray(parsed) ? parsed : [parsed];
      return arr.filter(function (item) {
        return item && (item.id || item.url || item.name);
      });
    } catch (err) { }
  }

  if (text.indexOf('\n') !== -1) {
    var lines = text.split('\n');
    var parsedLines = [];
    lines.forEach(function (ln) {
      var trimmed = ln.trim();
      if (!trimmed) return;
      var item = parseDocs_(trimmed);
      if (item && item.length) parsedLines.push(item[0]);
    });
    if (parsedLines.length) return parsedLines;
  }

  var cleanUrl = text;
  var displayName = '';
  var hypMatch = text.match(/HYPERLINK\s*\(\s*"([^"]+)"(?:\s*,\s*"([^"]+)")?\s*\)/i);
  if (hypMatch) {
    cleanUrl = hypMatch[1];
    displayName = hypMatch[2] || '';
  }

  var driveId = driveIdFromUrl_(cleanUrl);
  if (!driveId && cleanUrl.indexOf('http') !== 0 && cleanUrl.indexOf('data:') !== 0) {
    return [];
  }

  var fileName = displayName || cleanUrl.split('/').pop().split('?')[0] || 'Attachment';
  if (!fileName || fileName.toLowerCase() === 'view' || fileName.indexOf('open') === 0) fileName = 'Attachment';

  return [{
    id: driveId,
    name: fileName,
    url: cleanUrl,
    previewUrl: driveId ? ('https://drive.google.com/file/d/' + driveId + '/preview') : cleanUrl,
    thumbnailUrl: driveId ? ('https://drive.google.com/thumbnail?id=' + driveId + '&sz=w800') : '',
    type: /\.pdf($|\?)/i.test(cleanUrl) ? 'pdf' : (/\.(png|jpe?g|webp|gif|svg)($|\?)/i.test(cleanUrl) ? 'image' : 'file'),
    size: 0,
    uploadedAt: ''
  }];
}

function driveIdFromUrl_(url) {
  var text = String(url || '');
  var patterns = [/\/d\/([-\w]{20,})/, /[?&]id=([-\w]{20,})/, /\/file\/d\/([-\w]{20,})/];
  for (var i = 0; i < patterns.length; i++) {
    var m = text.match(patterns[i]);
    if (m) return m[1];
  }
  return '';
}

/* ==========================================================================
 * VALUE COERCION & RECORD READ/WRITE
 * ======================================================================== */

function coerceValue_(spec, value) {
  var type = spec[2] || 'text';
  if (type === 'doc') return serializeDocs_(value, spec[1]);
  if (type === 'bills') return formatBillCell_(parseBillEntries_(value));
  if (type === 'int') {
    if (value === '' || value === null || value === undefined) return '';
    var asInt = parseInt(String(value).replace(/[^0-9\-]/g, ''), 10);
    return isNaN(asInt) ? '' : asInt;
  }
  if (type === 'num') {
    if (value === '' || value === null || value === undefined) return '';
    var asNum = parseFloat(String(value).replace(/[^0-9.\-]/g, ''));
    return isNaN(asNum) ? '' : asNum;
  }
  if (type === 'date') return toDateStr_(value);
  return value === null || value === undefined ? '' : String(value).trim();
}

function buildRow_(sheetName, sheet, payload, existingRow, forcedId) {
  var specs = columnSpecs_(sheetName);
  var data = payload || {};
  var indexMap = columnIndexMap_(sheetName, sheet);
  var numHeaders = headersOf_(sheetName).length;
  var row = new Array(numHeaders);
  for (var c = 0; c < numHeaders; c++) {
    row[c] = (existingRow && existingRow[c] !== undefined) ? existingRow[c] : '';
  }

  specs.forEach(function (spec, i) {
    var key = spec[1];
    var col = indexMap[i];
    if (col < 0 || col >= numHeaders) return;
    if (forcedId && key === 'id') { row[col] = forcedId; return; }
    if (Object.prototype.hasOwnProperty.call(data, key) && data[key] !== undefined) {
      row[col] = coerceValue_(spec, data[key]);
    }
  });
  return row;
}

function recordFromValues_(sheetName, values, indexMap, richValues, notesRow) {
  var specs = columnSpecs_(sheetName);
  var record = {};

  specs.forEach(function (spec, i) {
    var type = spec[2] || 'text';
    var idx = indexMap ? indexMap[i] : i;
    var raw = (idx >= 0 && idx < values.length) ? values[idx] : undefined;

    if (type === 'doc') {
      var rich = (richValues && idx >= 0 && idx < richValues.length) ? richValues[idx] : null;
      var note = (notesRow && idx >= 0 && idx < notesRow.length) ? notesRow[idx] : '';
      record[spec[1]] = reconcileDocCell_(raw, rich, note);
    } else if (type === 'bills') {
      var bNote = (notesRow && idx >= 0 && idx < notesRow.length) ? notesRow[idx] : '';
      record[spec[1]] = parseBillEntries_(raw, bNote);
    } else if (type === 'int') {
      var asInt = typeof raw === 'number' ? Math.round(raw) : parseInt(String(raw || '').replace(/[^0-9\-]/g, ''), 10);
      record[spec[1]] = isNaN(asInt) ? 0 : asInt;
    } else if (type === 'num') {
      var asNum = typeof raw === 'number' ? raw : parseFloat(String(raw || '').replace(/[^0-9.\-]/g, ''));
      record[spec[1]] = isNaN(asNum) ? 0 : asNum;
    } else if (type === 'date') {
      record[spec[1]] = toDateStr_(raw);
    } else {
      record[spec[1]] = raw === null || raw === undefined ? '' : String(raw).trim();
    }
  });

  return record;
}

function records_(sheetName) {
  var sheet = getSheet_(sheetName, false);
  var last = sheet.getLastRow();
  if (last < 2) return [];

  var numHeaders = headersOf_(sheetName).length;
  var dataRange = sheet.getRange(2, 1, last - 1, numHeaders);
  var values = dataRange.getValues();
  var richValues = dataRange.getRichTextValues();
  var notes = dataRange.getNotes();
  var indexMap = columnIndexMap_(sheetName, sheet);
  var out = [];

  var idColIdx = indexMap[0];
  var siteColIdx = columnIndexByHeader_(sheetName, 'Client / Site Name', sheet) - 1;

  for (var i = 0; i < values.length; i++) {
    var idVal = (idColIdx >= 0 && idColIdx < values[i].length) ? String(values[i][idColIdx] || '').trim() : '';
    var siteVal = (siteColIdx >= 0 && siteColIdx < values[i].length) ? String(values[i][siteColIdx] || '').trim() : '';
    // Skip empty template rows that may hold dragged-down formulas
    if (!idVal && !siteVal) continue;

    var record = recordFromValues_(sheetName, values[i], indexMap, richValues[i], notes[i]);
    record._row = i + 2;
    out.push(record);
  }
  return out;
}

function readRowRecord_(sheetName, sheet, rowIndex) {
  var numHeaders = headersOf_(sheetName).length;
  var rowRange = sheet.getRange(rowIndex, 1, 1, numHeaders);
  var values = rowRange.getValues()[0];
  var richValues = rowRange.getRichTextValues()[0];
  var notes = rowRange.getNotes()[0];
  var indexMap = columnIndexMap_(sheetName, sheet);
  var record = recordFromValues_(sheetName, values, indexMap, richValues, notes);
  record._row = rowIndex;
  return record;
}

/**
 * Finds the first available row strictly within the submitting range (Headers width),
 * ensuring formulas in Bills AMT, Our Amount, or downstream columns
 * do NOT cause rows to be bypassed/skipped.
 */
function findNextAvailableRow_(sheetName, sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;

  var idCol = (columnIndexMap_(sheetName, sheet)[0] >= 0) ? (columnIndexMap_(sheetName, sheet)[0] + 1) : 1;
  var siteCol = columnIndexByHeader_(sheetName, 'Client / Site Name', sheet) || 2;

  var numRowsToCheck = Math.max(lastRow - 1, 1);
  var idValues = sheet.getRange(2, idCol, numRowsToCheck, 1).getValues();
  var siteValues = (siteCol && siteCol !== idCol)
    ? sheet.getRange(2, siteCol, numRowsToCheck, 1).getValues()
    : null;

  for (var i = 0; i < idValues.length; i++) {
    var idVal = String(idValues[i][0] || '').trim();
    var siteVal = siteValues ? String(siteValues[i][0] || '').trim() : '';

    // If both ID and Client / Site Name are blank, this row is empty in the submitting range!
    if (!idVal && !siteVal) {
      return i + 2;
    }
  }

  return lastRow + 1;
}

function upsertRow_(sheetName, payload, forcedId) {
  var sheet = getSheet_(sheetName, true);
  var idValue = forcedId || payload.id || '';
  var rowIndex = findRowById_(sheetName, sheet, idValue);
  var numHeaders = headersOf_(sheetName).length;
  var targetRow = rowIndex > 0 ? rowIndex : findNextAvailableRow_(sheetName, sheet);

  var existingRange = sheet.getRange(targetRow, 1, 1, numHeaders);
  var existing = existingRange.getValues()[0];
  var existingFormulas = existingRange.getFormulas()[0];

  var row = buildRow_(sheetName, sheet, payload, existing, idValue);

  // Preserve any formula function (e.g. in Bills AMT, Our Amount, or downstream columns)
  for (var c = 0; c < numHeaders; c++) {
    if (existingFormulas[c] && String(existingFormulas[c]).indexOf('=') === 0) {
      row[c] = existingFormulas[c];
    }
  }

  // Write strictly inside the submitting range defined by headers length
  sheet.getRange(targetRow, 1, 1, numHeaders).setValues([row.slice(0, numHeaders)]);

  var specs = columnSpecs_(sheetName);
  var hasDocUpdate = false;
  specs.forEach(function (spec) {
    if (spec[2] === 'doc') {
      var col = columnIndexByHeader_(sheetName, spec[0], sheet);
      if (col > 0 && payload && payload[spec[1]] !== undefined) {
        writeDocCell_(sheet, targetRow, col, payload[spec[1]], spec[1], true);
        hasDocUpdate = true;
      }
    } else if (spec[2] === 'bills') {
      var bCol = columnIndexByHeader_(sheetName, spec[0], sheet);
      if (bCol > 0 && payload && payload[spec[1]] !== undefined) {
        writeBillCell_(sheet, targetRow, bCol, payload[spec[1]]);
      }
    }
  });
  if (hasDocUpdate) {
    syncRowDocsToLinkedSheets_(sheet, sheetName, targetRow);
  }

  return readRowRecord_(sheetName, sheet, targetRow);
}

function deleteRowById_(sheetName, idValue) {
  var sheet = getSheet_(sheetName, false);
  var rowIndex = findRowById_(sheetName, sheet, idValue);
  if (rowIndex > 0) {
    sheet.deleteRow(rowIndex);
    return true;
  }
  return false;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    throw new Error('Another save is in progress. Please retry in a moment.');
  }
  try {
    return fn();
  } finally {
    try { SpreadsheetApp.flush(); } catch (e) { }
    lock.releaseLock();
  }
}

/* ==========================================================================
 * READ APIs
 * ======================================================================== */

function getInitialData() {
  ensureSchema_();
  return {
    success: true,
    masterQuotations: records_(SHEET.QUOTATION_MASTER),
    quoteWin: [],
    quoteLoss: [],
    purchaseOrders: records_(SHEET.PO_MASTER),
    doneArchive: [],
    config: getConfigData(),
    meta: {
      app: APP_NAME,
      subtitle: APP_SUBTITLE,
      currency: CURRENCY,
      generatedAt: new Date().toISOString(),
      sheetNames: {
        master: SHEET.QUOTATION_MASTER,
        po: SHEET.PO_MASTER,
        config: SHEET.CONFIG_DROPDOWNS
      }
    }
  };
}

function getConfigData() {
  var sheet = getSheet_(SHEET.CONFIG_DROPDOWNS, true);
  var last = sheet.getLastRow();
  var out = {};

  Object.keys(CONFIG_COLUMN_INDEX).forEach(function (key) {
    out[key] = (CONFIG_DEFAULTS[key] || []).slice();
  });

  if (last >= 2) {
    var values = sheet.getRange(2, 1, last - 1, CONFIG_HEADERS.length).getValues();
    Object.keys(CONFIG_COLUMN_INDEX).forEach(function (key) {
      var col = CONFIG_COLUMN_INDEX[key] - 1;
      var seen = {};
      var list = [];
      for (var r = 0; r < values.length; r++) {
        var text = String(values[r][col] || '').trim();
        if (!text) continue;
        var dedupe = text.toLowerCase();
        if (seen[dedupe]) continue;
        seen[dedupe] = true;
        list.push(text);
      }
      if (list.length) out[key] = list;
    });
  }
  return out;
}

/* ==========================================================================
 * WRITE APIs
 * ======================================================================== */

function createQuotation(payload) {
  return withLock_(function () {
    var id = nextId_(ID_PREFIX.QUOTATION, [SHEET.QUOTATION_MASTER]);
    var record = upsertRow_(SHEET.QUOTATION_MASTER, payload || {}, id);
    return { success: true, quotationId: id, record: record };
  });
}

function createDirectPO(payload) {
  return withLock_(function () {
    var id = nextId_(ID_PREFIX.PO, [SHEET.PO_MASTER]);
    var record = upsertRow_(SHEET.PO_MASTER, payload || {}, id);
    return { success: true, poId: id, record: record };
  });
}

function updateRecordFull(sheetName, idValue, payload) {
  if (!HEADERS[sheetName]) throw new Error('Unknown sheet: ' + sheetName);
  if (!idValue) throw new Error('Missing primary key for ' + sheetName);

  return withLock_(function () {
    var sheet = getSheet_(sheetName, false);
    var rowIndex = findRowById_(sheetName, sheet, idValue);
    if (rowIndex < 0) throw new Error('Record not found in ' + sheetName + ' -> ' + idValue);

    var numHeaders = headersOf_(sheetName).length;
    var existingRange = sheet.getRange(rowIndex, 1, 1, numHeaders);
    var existing = existingRange.getValues()[0];
    var existingFormulas = existingRange.getFormulas()[0];
    var existingRich = existingRange.getRichTextValues()[0];
    var existingNotes = existingRange.getNotes()[0];

    var row = buildRow_(sheetName, sheet, payload || {}, existing, idValue);
    var idColIdx = columnIndexMap_(sheetName, sheet)[0];
    if (idColIdx >= 0) row[idColIdx] = String(idValue).trim();

    var billsCol = columnIndexByHeader_(sheetName, 'Bills', sheet);
    var billsAmtCol = columnIndexByHeader_(sheetName, 'Bills AMT', sheet);
    if (billsCol && billsAmtCol) {
      var hasBillsAmtFormula = existingFormulas[billsAmtCol - 1] && String(existingFormulas[billsAmtCol - 1]).indexOf('=') === 0;
      if (!hasBillsAmtFormula) {
        var sourceBills = (payload && payload.bills !== undefined)
          ? parseBillEntries_(payload.bills)
          : parseBillEntries_(existing[billsCol - 1], existingNotes[billsCol - 1]);
        var calculatedSum = sumBillEntries_(sourceBills);
        if (calculatedSum > 0) {
          row[billsAmtCol - 1] = calculatedSum;
        } else if (payload && payload.billsAmt !== undefined && payload.billsAmt !== '') {
          row[billsAmtCol - 1] = toAmount_(payload.billsAmt);
        }
      }
    }

    // Preserve any existing formula function (e.g. Our Amount, Bills AMT)
    for (var c = 0; c < numHeaders; c++) {
      if (existingFormulas[c] && String(existingFormulas[c]).indexOf('=') === 0) {
        row[c] = existingFormulas[c];
      }
    }

    sheet.getRange(rowIndex, 1, 1, numHeaders).setValues([row.slice(0, numHeaders)]);

    var specs = columnSpecs_(sheetName);
    var hasDocUpdate = false;
    specs.forEach(function (spec) {
      if (spec[2] === 'doc') {
        var col = columnIndexByHeader_(sheetName, spec[0], sheet);
        if (col > 0) {
          var docToSave = (payload && payload[spec[1]] !== undefined)
            ? payload[spec[1]]
            : reconcileDocCell_(existing[col - 1], existingRich[col - 1], existingNotes[col - 1]);
          writeDocCell_(sheet, rowIndex, col, docToSave, spec[1], true);
          hasDocUpdate = true;
        }
      } else if (spec[2] === 'bills') {
        var bCol = columnIndexByHeader_(sheetName, spec[0], sheet);
        if (bCol > 0) {
          var billsToSave = (payload && payload[spec[1]] !== undefined)
            ? payload[spec[1]]
            : parseBillEntries_(existing[bCol - 1], existingNotes[bCol - 1]);
          writeBillCell_(sheet, rowIndex, bCol, billsToSave);
        }
      }
    });
    if (hasDocUpdate) {
      syncRowDocsToLinkedSheets_(sheet, sheetName, rowIndex);
    }

    return { success: true, sheet: sheetName, id: idValue, record: readRowRecord_(sheetName, sheet, rowIndex) };
  });
}

function updateTableCell(sheetName, idValue, columnName, value) {
  if (!HEADERS[sheetName]) throw new Error('Unknown sheet: ' + sheetName);
  return withLock_(function () {
    var sheet = getSheet_(sheetName, false);
    var rowIndex = findRowById_(sheetName, sheet, idValue);
    if (rowIndex < 0) throw new Error('Record not found in ' + sheetName + ' -> ' + idValue);

    var col = columnIndexByHeader_(sheetName, columnName, sheet);
    var spec = columnSpecByHeader_(sheetName, columnName);
    if (!col || !spec) throw new Error('Unknown column "' + columnName + '" on ' + sheetName);

    var written = coerceValue_(spec, value);
    if (spec[2] === 'doc') {
      writeDocCell_(sheet, rowIndex, col, value, spec[1]);
    } else if (spec[2] === 'bills') {
      writeBillCell_(sheet, rowIndex, col, value);
    } else {
      sheet.getRange(rowIndex, col).setValue(written);
    }

    return {
      success: true, sheet: sheetName, id: idValue,
      column: spec[0], value: written,
      record: readRowRecord_(sheetName, sheet, rowIndex)
    };
  });
}

function deleteRecord(sheetName, idValue) {
  if (!HEADERS[sheetName]) throw new Error('Unknown sheet: ' + sheetName);
  return withLock_(function () {
    var removed = deleteRowById_(sheetName, idValue);
    return { success: removed, sheet: sheetName, id: idValue };
  });
}

/* ==========================================================================
 * PIPELINE TRANSITION (Quotation Master -> Purchase Order Master)
 * ======================================================================== */

function submitDecision(quotationId, outcome, reason, payload) {
  return withLock_(function () {
    var isWon = !outcome || String(outcome || '').toLowerCase() === 'won';
    if (!isWon) {
      deleteRowById_(SHEET.QUOTATION_MASTER, quotationId);
      return { success: true, outcome: 'Lost', quotationId: quotationId };
    }

    var masterSheet = getSheet_(SHEET.QUOTATION_MASTER, false);
    var rowIndex = findRowById_(SHEET.QUOTATION_MASTER, masterSheet, quotationId);
    if (rowIndex < 0) throw new Error('Quotation not found in Master: ' + quotationId);

    var master = readRowRecord_(SHEET.QUOTATION_MASTER, masterSheet, rowIndex);
    var overrides = payload || {};
    var merged = {};
    Object.keys(master).forEach(function (key) {
      if (key.indexOf('_') === 0) return;
      merged[key] = master[key];
    });
    Object.keys(overrides).forEach(function (key) {
      if (overrides[key] !== undefined && overrides[key] !== null && overrides[key] !== '') {
        merged[key] = overrides[key];
      }
    });

    var decisionDate = toDateStr_(overrides.decisionDate) || today_();
    var poId = String(overrides.poId || '').trim() || nextId_(ID_PREFIX.PO, [SHEET.PO_MASTER]);
    var poPayload = {};
    Object.keys(merged).forEach(function (key) { poPayload[key] = merged[key]; });
    poPayload.id = poId;
    poPayload.quotationId = quotationId;
    poPayload.date = toDateStr_(overrides.poDate) || decisionDate;
    poPayload.poOrder = overrides.poOrder !== undefined ? overrides.poOrder : (merged.poOrder || []);
    poPayload.bills = overrides.bills !== undefined ? overrides.bills : (merged.bills || []);
    var sumAmt = sumBillAmounts_(poPayload.bills);
    poPayload.billsAmt = sumAmt > 0 ? sumAmt : toAmount_(overrides.billsAmt);

    var poRecord = upsertRow_(SHEET.PO_MASTER, poPayload, poId);
    deleteRowById_(SHEET.QUOTATION_MASTER, quotationId);

    return {
      success: true,
      outcome: 'Won',
      quotationId: quotationId,
      poId: poId,
      poRecord: poRecord
    };
  });
}

/* ==========================================================================
 * CONFIG LOOKUP APIs
 * ======================================================================== */

function saveConfigOption(columnKey, value) {
  var col = CONFIG_COLUMN_INDEX[columnKey];
  if (!col) throw new Error('Unknown config column: ' + columnKey);
  var text = String(value === null || value === undefined ? '' : value).trim();
  if (!text) throw new Error('Cannot save an empty lookup value.');

  return withLock_(function () {
    var sheet = getSheet_(SHEET.CONFIG_DROPDOWNS, true);
    var config = getConfigData();
    var list = config[columnKey] || [];

    var exists = list.some(function (item) { return item.toLowerCase() === text.toLowerCase(); });
    if (exists) return { success: true, duplicated: true, columnKey: columnKey, config: config };

    var last = sheet.getLastRow();
    var targetRow = 2;
    if (last >= 2) {
      var used = sheet.getRange(2, col, last - 1, 1).getValues();
      for (var i = 0; i < used.length; i++) {
        if (String(used[i][0] || '').trim()) targetRow = i + 3;
      }
    }
    sheet.getRange(targetRow, col).setValue(text);
    return { success: true, duplicated: false, columnKey: columnKey, value: text, config: getConfigData() };
  });
}

function deleteConfigOption(columnKey, value) {
  var col = CONFIG_COLUMN_INDEX[columnKey];
  if (!col) throw new Error('Unknown config column: ' + columnKey);
  var text = String(value === null || value === undefined ? '' : value).trim().toLowerCase();

  return withLock_(function () {
    var sheet = getSheet_(SHEET.CONFIG_DROPDOWNS, true);
    var last = sheet.getLastRow();
    if (last >= 2) {
      var range = sheet.getRange(2, col, last - 1, 1);
      var values = range.getValues();
      var changed = false;
      for (var i = 0; i < values.length; i++) {
        if (String(values[i][0] || '').trim().toLowerCase() === text) {
          values[i][0] = '';
          changed = true;
        }
      }
      if (changed) range.setValues(values);
    }
    return { success: true, columnKey: columnKey, value: value, config: getConfigData() };
  });
}

function refreshValidations() {
  applyValidations_();
  return { success: true, config: getConfigData() };
}

/* ==========================================================================
 * MULTI-FILE UPLOAD & DELETE ENGINE (Google Drive Backed)
 * ======================================================================== */

function getRootFolder_() {
  var iterator = DriveApp.getFoldersByName(DRIVE_ROOT_FOLDER_NAME);
  if (iterator.hasNext()) return iterator.next();
  return DriveApp.createFolder(DRIVE_ROOT_FOLDER_NAME);
}

function getRecordFolder_(recordId) {
  var root = getRootFolder_();
  var folderName = String(recordId || 'unassigned').trim() || 'unassigned';
  var iterator = root.getFoldersByName(folderName);
  if (iterator.hasNext()) return iterator.next();
  return root.createFolder(folderName);
}

function guessDocType_(mimeType, fileName) {
  var mime = String(mimeType || '').toLowerCase();
  var name = String(fileName || '').toLowerCase();
  if (mime.indexOf('pdf') > -1 || /\.pdf$/.test(name)) return 'pdf';
  if (mime.indexOf('image') > -1 || /\.(png|jpe?g|gif|bmp|webp|svg|heic)$/.test(name)) return 'image';
  if (/\.(xlsx?|csv)$/.test(name)) return 'sheet';
  if (/\.(docx?|txt)$/.test(name)) return 'doc';
  if (/\.(dwg|dxf)$/.test(name)) return 'cad';
  return 'file';
}

/**
 * Batch-uploads multiple document slots (e.g. roughSizes, vendorQuote, ourQuote, poOrder)
 * in a single call. Uploads to Drive outside the sheet lock, then writes all slots
 * and synchronizes linked operational sheets in ONE single pass.
 *
 * @param {string} sheetName - e.g. SHEET.QUOTATION_MASTER or SHEET.PO_MASTER
 * @param {string} recordId  - e.g. 'QT-UPVC-0001'
 * @param {Array<Object>} filePayloads - Array of { docKey, base64Data, mimeType, fileName }
 */
function uploadRecordFilesBatch(sheetName, recordId, filePayloads) {
  if (!HEADERS[sheetName] && SHEET[sheetName]) sheetName = SHEET[sheetName];
  if (!HEADERS[sheetName]) throw new Error('Unknown sheet: ' + sheetName);
  if (!Array.isArray(filePayloads) || !filePayloads.length) {
    return { success: true, id: recordId, uploaded: {} };
  }

  var sheet = getSheet_(sheetName, false);
  if (findRowById_(sheetName, sheet, recordId) < 0) {
    throw new Error('Record not found in ' + sheetName + ' -> ' + recordId);
  }

  // 1. Upload all files to Drive FIRST (outside the Sheet Lock so network I/O doesn't block the sheet)
  var folder = getRecordFolder_(recordId);
  var preparedBySlot = {}; // { docKey: { spec: spec, attachment: attachment } }

  filePayloads.forEach(function (item) {
    if (!item || !item.docKey || !item.base64Data) return;
    var spec = columnSpecByKey_(sheetName, item.docKey);
    if (!spec || spec[2] !== 'doc') return;

    var base64 = String(item.base64Data || '');
    var comma = base64.indexOf(',');
    if (comma > -1 && base64.substring(0, comma).indexOf('base64') > -1) {
      base64 = base64.substring(comma + 1);
    }

    var safeName = String(item.fileName || 'upload').replace(/[\\\/:*?"<>|]/g, '_');
    var mimeType = item.mimeType || 'application/octet-stream';
    var bytes = Utilities.base64Decode(base64);
    var blob = Utilities.newBlob(bytes, mimeType, safeName);
    var file = folder.createFile(blob);

    if (SHARE_UPLOADS_WITH_LINK) {
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (err) {
        Logger.log('Sharing skipped for ' + safeName + ': ' + err);
      }
    }

    var fileId = file.getId();
    preparedBySlot[item.docKey] = {
      spec: spec,
      attachment: {
        id: fileId,
        name: safeName,
        url: file.getUrl(),
        previewUrl: 'https://drive.google.com/file/d/' + fileId + '/preview',
        thumbnailUrl: 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w800',
        downloadUrl: 'https://drive.google.com/uc?export=download&id=' + fileId,
        type: guessDocType_(mimeType, safeName),
        mime: mimeType,
        size: bytes.length,
        uploadedAt: new Date().toISOString()
      }
    };
  });

  // 2. Acquire lock ONCE and write all uploaded slots + sync once
  try {
    return withLock_(function () {
      var currentRow = findRowById_(sheetName, sheet, recordId);
      if (currentRow < 0) throw new Error('Record not found in ' + sheetName + ' -> ' + recordId);

      var uploadedResult = {};
      var docKeys = Object.keys(preparedBySlot);

      docKeys.forEach(function (docKey) {
        var entry = preparedBySlot[docKey];
        var col = columnIndexByHeader_(sheetName, entry.spec[0], sheet);
        if (col > 0) {
          var cellRange = sheet.getRange(currentRow, col);
          var oldFiles = reconcileDocCell_(cellRange.getValue(), cellRange.getRichTextValue(), cellRange.getNote());

          // Skip individual sync; we will sync all slots together below
          writeDocCell_(sheet, currentRow, col, [entry.attachment], docKey, true);
          uploadedResult[docKey] = [entry.attachment];

          // Trash superseded file(s) from Drive
          oldFiles.forEach(function (oldFile) {
            var oldDriveId = oldFile.id || driveIdFromUrl_(oldFile.url);
            if (oldDriveId && oldDriveId !== entry.attachment.id) {
              try {
                var oldDriveFile = DriveApp.getFileById(oldDriveId);
                if (isManagedFile_(oldDriveFile)) oldDriveFile.setTrashed(true);
              } catch (trashErr) {
                Logger.log('Old revision trash skipped for ' + oldDriveId + ': ' + trashErr);
              }
            }
          });
        }
      });

      // Sync all updated doc cells to linked FMS sheets in a SINGLE pass
      syncRowDocsToLinkedSheets_(sheet, sheetName, currentRow);

      return {
        success: true,
        sheet: sheetName,
        id: recordId,
        uploaded: uploadedResult,
        record: readRowRecord_(sheetName, sheet, currentRow)
      };
    });
  } catch (err) {
    // If saving to the sheet failed/timed out, clean up the newly created Drive file(s)
    Object.keys(preparedBySlot).forEach(function (k) {
      try {
        var fId = preparedBySlot[k].attachment.id;
        DriveApp.getFileById(fId).setTrashed(true);
      } catch (cleanupErr) { }
    });
    throw err;
  }
}

/**
 * Replaces the file in any [DOC] column (Rough Sizes, Vendor Quote,
 * Our Quote, PO Order) and trashes the superseded Drive file.
 * Wrapper around uploadRecordFilesBatch for backward compatibility.
 */
function uploadRecordFile(sheetName, recordId, docKey, base64Data, mimeType, fileName) {
  var batchRes = uploadRecordFilesBatch(sheetName, recordId, [{
    docKey: docKey,
    base64Data: base64Data,
    mimeType: mimeType,
    fileName: fileName
  }]);
  var uploaded = (batchRes.uploaded && batchRes.uploaded[docKey]) || [];
  return {
    success: true,
    sheet: sheetName,
    id: recordId,
    docKey: docKey,
    column: columnSpecByKey_(sheetName, docKey) ? columnSpecByKey_(sheetName, docKey)[0] : '',
    file: uploaded[0] || null,
    files: uploaded
  };
}

/* ==========================================================================
 * INVOICE / BILLS ENTRY HELPERS
 * Formats Bills as:
 * 1.Date Invoice Number Amt
 * 2.Date Invoice Number Amt
 * in one cell.
 * ======================================================================== */

function parseBillEntries_(value, note) {
  if (Array.isArray(value)) return value;
  if (note) {
    try {
      var parsed = JSON.parse(note);
      if (Array.isArray(parsed) && parsed.length) return parsed;
    } catch (e) { }
  }
  var text = String(value || '').trim();
  if (!text) return [];

  if (text.indexOf('[') === 0) {
    try {
      var arr = JSON.parse(text);
      if (Array.isArray(arr)) return arr;
    } catch (e) { }
  }

  var lines = text.split('\n');
  var out = [];
  lines.forEach(function (ln) {
    var line = ln.trim();
    if (!line) return;
    var m = line.match(/^(?:\d+[\.\)]\s*)?(\S+)\s+(.+?)\s+([0-9,]+(?:\.[0-9]+)?)$/);
    if (m) {
      out.push({
        date: m[1],
        invoiceNo: m[2],
        amount: toAmount_(m[3].replace(/,/g, ''))
      });
    } else {
      var parts = line.replace(/^\d+[\.\)]\s*/, '').split(/\s+/);
      if (parts.length >= 2) {
        var lastPart = parts[parts.length - 1];
        var amt = toAmount_(lastPart.replace(/,/g, ''));
        var dt = parts[0];
        var inv = parts.slice(1, parts.length - 1).join(' ') || '-';
        out.push({ date: dt, invoiceNo: inv, amount: amt });
      } else {
        out.push({ date: '', invoiceNo: line, amount: 0 });
      }
    }
  });
  return out;
}

function formatBillCell_(entries) {
  var list = Array.isArray(entries) ? entries : [];
  if (!list.length) return '';
  return list.map(function (b, i) {
    var d = toDateStr_(b.date) || formatShortDate_(b.date);
    var inv = String(b.invoiceNo || b.invoice || b.billNo || '').trim() || '-';
    var amt = toAmount_(b.amount !== undefined ? b.amount : b.amt);
    return (i + 1) + '.' + d + ' ' + inv + ' ' + amt;
  }).join('\n');
}

function sumBillEntries_(entries) {
  var list = Array.isArray(entries) ? entries : [];
  return list.reduce(function (sum, b) {
    return sum + toAmount_(b && (b.amount !== undefined ? b.amount : b.amt));
  }, 0);
}

function writeBillCell_(sheet, rowIndex, col, value) {
  var list = parseBillEntries_(value);
  var text = formatBillCell_(list);
  var range = sheet.getRange(rowIndex, col);
  range.setValue(text);
  range.setNote(list.length ? JSON.stringify(list) : '');
  range.setWrap(true);

  var sheetName = sheet.getName();
  var billsAmtCol = columnIndexByHeader_(sheetName, 'Bills AMT', sheet);
  if (billsAmtCol) {
    var formula = sheet.getRange(rowIndex, billsAmtCol).getFormula();
    if (!formula) {
      var total = sumBillEntries_(list);
      sheet.getRange(rowIndex, billsAmtCol).setValue(total);
    }
  }
  return list;
}

function saveRecordBills(sheetName, recordId, billsList) {
  if (!HEADERS[sheetName]) throw new Error('Unknown sheet: ' + sheetName);
  return withLock_(function () {
    var sheet = getSheet_(sheetName, false);
    var rowIndex = findRowById_(sheetName, sheet, recordId);
    if (rowIndex < 0) throw new Error('Record not found -> ' + recordId);

    var billsCol = columnIndexByHeader_(sheetName, 'Bills', sheet);
    if (!billsCol) throw new Error('Bills column not found on ' + sheetName);

    var list = writeBillCell_(sheet, rowIndex, billsCol, billsList);
    var total = sumBillEntries_(list);

    return {
      success: true,
      sheet: sheetName,
      id: recordId,
      bills: list,
      billsAmt: total,
      formattedText: formatBillCell_(list)
    };
  });
}

function toAmount_(value) {
  if (value === '' || value === null || value === undefined) return 0;
  var n = parseFloat(String(value).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function sumBillAmounts_(files) {
  return sumBillEntries_(parseBillEntries_(files));
}

function writeBillsTotal_(sheet, rowIndex, sheetName, files) {
  var total = sumBillAmounts_(files);
  var totalCol = columnIndexByHeader_(sheetName, 'Bills AMT', sheet);
  if (totalCol) {
    var formula = sheet.getRange(rowIndex, totalCol).getFormula();
    // Do NOT overwrite user formula function in Bills AMT
    if (!formula && total > 0) {
      sheet.getRange(rowIndex, totalCol).setValue(total);
    }
  }
  return total;
}

function deleteRecordFile(sheetName, recordId, docKey) {
  if (!HEADERS[sheetName]) throw new Error('Unknown sheet: ' + sheetName);
  var spec = columnSpecByKey_(sheetName, docKey);
  if (!spec || spec[2] !== 'doc') throw new Error('Not a [DOC] column: ' + docKey);

  return withLock_(function () {
    var sheet = getSheet_(sheetName, false);
    var rowIndex = findRowById_(sheetName, sheet, recordId);
    if (rowIndex < 0) throw new Error('Record not found -> ' + recordId);

    var col = columnIndexByHeader_(sheetName, spec[0], sheet);
    var cellRange = sheet.getRange(rowIndex, col);
    var existing = reconcileDocCell_(cellRange.getValue(), cellRange.getRichTextValue(), cellRange.getNote());
    var removed = existing.length ? existing[0] : null;

    // Clear cell & sync empty state to linked sheets
    writeDocCell_(sheet, rowIndex, col, [], docKey);

    if (removed) {
      var driveId = removed.id || driveIdFromUrl_(removed.url);
      if (driveId) {
        try {
          var driveFile = DriveApp.getFileById(driveId);
          if (isManagedFile_(driveFile)) driveFile.setTrashed(true);
        } catch (err) {
          Logger.log('Drive trash skipped for ' + driveId + ': ' + err);
        }
      }
    }

    return {
      success: true,
      sheet: sheetName,
      id: recordId,
      docKey: docKey,
      removed: removed,
      files: []
    };
  });
}

function isManagedFile_(file) {
  try {
    var rootName = DRIVE_ROOT_FOLDER_NAME;
    var parents = file.getParents();
    while (parents.hasNext()) {
      var parent = parents.next();
      if (parent.getName() === rootName) return true;
      var grandParents = parent.getParents();
      while (grandParents.hasNext()) {
        if (grandParents.next().getName() === rootName) return true;
      }
    }
  } catch (err) { }
  return false;
}

/* ==========================================================================
 * MAINTENANCE & MIGRATION UTILITIES
 * ======================================================================== */

function repairLegacyIds() {
  return withLock_(function () {
    var report = {
      master: repairSheetIds_(SHEET.QUOTATION_MASTER, ID_PREFIX.QUOTATION, [SHEET.QUOTATION_MASTER]),
      po: repairSheetIds_(SHEET.PO_MASTER, ID_PREFIX.PO, [SHEET.PO_MASTER])
    };
    SpreadsheetApp.flush();
    var assigned = report.master + report.po;
    return { success: true, assigned: assigned, perSheet: report };
  });
}

function repairSheetIds_(sheetName, prefix, familySheets) {
  var sheet = getSheet_(sheetName, false);
  syncSheetColumns_(sheet, headersOf_(sheetName));
  var last = sheet.getLastRow();
  if (last < 2) return 0;

  var width = Math.max(sheet.getLastColumn(), headersOf_(sheetName).length, 1);
  var idCol = columnIndexMap_(sheetName, sheet)[0] + 1;
  if (idCol < 1) return 0;
  var values = sheet.getRange(2, 1, last - 1, width).getValues();
  var counter = highestIdNumber_(prefix, familySheets);
  var assigned = 0;

  for (var i = 0; i < values.length; i++) {
    var raw = values[i][idCol - 1];
    var original = String(raw === null || raw === undefined ? '' : raw);
    var current = original.trim();
    if (!current) {
      if (isBlankRow_(values[i])) continue;
      current = prefix + pad_(++counter, ID_PAD);
      assigned++;
    }
    if (current !== original) {
      sheet.getRange(i + 2, idCol).setValue(current);
    }
  }
  return assigned;
}

function migrateAllDocsToRichText() {
  var targetSheets = [SHEET.QUOTATION_MASTER, SHEET.PO_MASTER];
  var convertedCount = 0;
  var details = [];

  targetSheets.forEach(function (sheetName) {
    var sheet = getSheet_(sheetName, false);
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    var specs = columnSpecs_(sheetName);
    var docSpecs = specs.filter(function (s) { return s[2] === 'doc'; });
    var sheetConverted = 0;

    docSpecs.forEach(function (spec) {
      var col = columnIndexByHeader_(sheetName, spec[0], sheet);
      if (!col) return;

      var range = sheet.getRange(2, col, lastRow - 1, 1);
      var values = range.getValues();
      var richVals = range.getRichTextValues();
      var notes = range.getNotes();

      for (var i = 0; i < values.length; i++) {
        var rowIndex = i + 2;
        var files = reconcileDocCell_(values[i][0], richVals[i][0], notes[i][0]);
        if (files && files.length) {
          writeDocCell_(sheet, rowIndex, col, files, spec[1], true);
          if (spec[1] === 'bills') {
            writeBillsTotal_(sheet, rowIndex, sheetName, files);
          }
          sheetConverted++;
        }
      }
    });

    convertedCount += sheetConverted;
    details.push(sheetName + ': ' + sheetConverted + ' document cell(s) formatted');
  });

  // Automatically propagate converted links to any linked operational sheets (uPVC PO, uPVC Quotation)
  try {
    syncAllDocLinksToAllSheets(true);
  } catch (errSync) {
    Logger.log('Automatic syncAllDocLinksToAllSheets warning: ' + errSync);
  }

  var msg = 'Migration complete. ' + convertedCount + ' document cell(s) converted to clickable hyperlinks.\n' + details.join('\n');
  try {
    SpreadsheetApp.getUi().alert('Document Links Conversion', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) { }
  return { success: true, convertedCount: convertedCount, details: details };
}

/* ==========================================================================
 * LINKED FMS SHEET RICH-TEXT SYNC ENGINE (OPTION B)
 * Propagates multi-file clickable hyperlinks to operational sheets (e.g. uPVC PO,
 * uPVC Quotation) that reference master records via =INDEX() formulas or
 * Client / Site Name lookups.
 * ======================================================================== */

function letterToColumn_(letter) {
  var str = String(letter || '').toUpperCase().trim();
  var col = 0;
  for (var i = 0; i < str.length; i++) {
    col = col * 26 + (str.charCodeAt(i) - 64);
  }
  return col;
}

/**
 * Intelligently inspects any operational sheet to locate the header row,
 * the record matching column (Client / Site Name or ID), and the document columns.
 */
function getSheetDocColumnMapping_(dstSheet, masterSheetName) {
  var dName = dstSheet.getName();
  var cacheKey = dName + '|' + masterSheetName;
  if (_sheetMappingCache[cacheKey]) return _sheetMappingCache[cacheKey];

  var lastRow = dstSheet.getLastRow();
  var lastCol = dstSheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    var emptyRes = { headerRow: 6, idCol: 0, siteCol: 0, docCols: {} };
    _sheetMappingCache[cacheKey] = emptyRes;
    return emptyRes;
  }

  var docSpecs = columnSpecs_(masterSheetName).filter(function (s) { return s[2] === 'doc'; });
  var idSpec = columnSpecs_(masterSheetName)[0];

  // FMS tabs typically use Row 6, while flat tables use Row 1. Scan candidate rows.
  var candidateRows = [6, 1, 2, 3, 4, 5, 7, 8, 9, 10];
  var headerRow = 6;
  var bestScore = -1;

  for (var i = 0; i < candidateRows.length; i++) {
    var r = candidateRows[i];
    if (r > lastRow) continue;
    var rowVals = dstSheet.getRange(r, 1, 1, lastCol).getValues()[0];
    var score = 0;
    for (var c = 0; c < rowVals.length; c++) {
      var val = normHeader_(rowVals[c]);
      if (!val) continue;
      if (headerPresent_('Client / Site Name', [val])) score += 6;
      if (headerPresent_(idSpec[0], [val])) score += 5;
      if (headerPresent_('Quotation ID', [val])) score += 4;
      for (var d = 0; d < docSpecs.length; d++) {
        if (headerPresent_(docSpecs[d][0], [val])) score += 3;
      }
    }
    if (score > bestScore) {
      bestScore = score;
      headerRow = r;
    }
  }

  var headers = dstSheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(normHeader_);
  var idCol = 0;
  var siteCol = 0;
  var docCols = {};

  // 1. Identify Client / Site Name column
  for (var c = 0; c < headers.length; c++) {
    if (headerPresent_('Client / Site Name', [headers[c]])) {
      siteCol = c + 1;
      break;
    }
  }

  // 2. Identify ID column (e.g. PO ID or Quotation ID)
  for (var c = 0; c < headers.length; c++) {
    if (headerPresent_(idSpec[0], [headers[c]])) {
      idCol = c + 1;
      break;
    }
  }
  if (!idCol && masterSheetName === SHEET.PO_MASTER) {
    for (var c = 0; c < headers.length; c++) {
      if (headerPresent_('Quotation ID', [headers[c]])) {
        idCol = c + 1;
        break;
      }
    }
  }

  // 3. Fallback: inspect data rows for ID prefixes (e.g. PO-UPVC- or QT-UPVC-) if idCol not found
  if (!idCol && lastRow > headerRow) {
    var scanSample = Math.min(lastRow - headerRow, 15);
    var sampleCols = Math.min(lastCol, 10);
    var sampleValues = dstSheet.getRange(headerRow + 1, 1, scanSample, sampleCols).getValues();
    var prefix = (masterSheetName === SHEET.QUOTATION_MASTER) ? 'QT-UPVC-' : 'PO-UPVC-';

    for (var sc = 0; sc < sampleCols; sc++) {
      var matchCount = 0;
      for (var sr = 0; sr < sampleValues.length; sr++) {
        var str = String(sampleValues[sr][sc] || '').trim().toUpperCase();
        if (str.indexOf(prefix) === 0) matchCount++;
      }
      if (matchCount > 0) {
        idCol = sc + 1;
        break;
      }
    }
  }

  // 4. Identify document columns by header
  docSpecs.forEach(function (spec) {
    for (var c = 0; c < headers.length; c++) {
      if (headerPresent_(spec[0], [headers[c]])) {
        docCols[spec[1]] = c + 1;
        break;
      }
    }
  });

  // 5. Fallback formula inspection for =INDEX(masterSheet!...)
  if (lastRow > headerRow) {
    var checkFormulaRows = Math.min(lastRow - headerRow, 5);
    var formulas = dstSheet.getRange(headerRow + 1, 1, checkFormulaRows, lastCol).getFormulas();
    var masterEscaped = masterSheetName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    var formulaRegex = new RegExp("['\"]?" + masterEscaped + "['\"]?!\\$?([A-Z]+)", "i");

    for (var fc = 0; fc < lastCol; fc++) {
      for (var fr = 0; fr < formulas.length; fr++) {
        var form = formulas[fr][fc];
        if (!form) continue;
        var match = form.match(formulaRegex);
        if (match) {
          var targetMasterCol = letterToColumn_(match[1]);
          docSpecs.forEach(function (ds) {
            var mCol = columnIndexByHeader_(masterSheetName, ds[0]);
            if (mCol === targetMasterCol) {
              docCols[ds[1]] = fc + 1;
            }
          });

          if (!siteCol && !idCol) {
            var matchColMatch = form.match(/MATCH\s*\(\s*\$?([A-Z]+)\d+/i);
            if (matchColMatch) {
              var colNum = letterToColumn_(matchColMatch[1]);
              if (headers[colNum - 1] && headerPresent_('Client / Site Name', [headers[colNum - 1]])) {
                siteCol = colNum;
              } else {
                idCol = colNum;
              }
            }
          }
          break;
        }
      }
    }
  }

  var res = {
    headerRow: headerRow,
    idCol: idCol,
    siteCol: siteCol,
    docCols: docCols
  };
  _sheetMappingCache[cacheKey] = res;
  return res;
}

/**
 * Synchronizes a single document cell update across all linked operational sheets.
 */
function syncDocCellToLinkedSheets_(masterSheetName, recordId, siteName, quotationId, docKey, richText, noteStr) {
  if (!recordId && !siteName) return;
  var ss = getSpreadsheet_();
  var allSheets = ss.getSheets();

  var normId = normHeader_(recordId);
  var normSite = normHeader_(siteName);
  var normQId = normHeader_(quotationId);

  var excludedSheets = [
    masterSheetName.toLowerCase(),
    SHEET.CONFIG_DROPDOWNS.toLowerCase(),
    'setupconfig'
  ];

  allSheets.forEach(function (dstSheet) {
    var dName = dstSheet.getName();
    if (excludedSheets.indexOf(dName.toLowerCase()) > -1) return;
    if (dName.toLowerCase().indexOf('cancel') > -1 || dName.toLowerCase().indexOf('done') > -1) return;
    if (dName.toLowerCase() === 'summary' || dName.toLowerCase().indexOf('enquiry') > -1) return;

    var lastRow = dstSheet.getLastRow();
    var lastCol = dstSheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return;

    var mapping = getSheetDocColumnMapping_(dstSheet, masterSheetName);
    if (!mapping.docCols[docKey]) return;
    if (!mapping.siteCol && !mapping.idCol) return;

    var targetCol = mapping.docCols[docKey];
    var scanStart = mapping.headerRow + 1;
    var numRows = lastRow - scanStart + 1;
    if (numRows < 1) return;

    var siteValues = mapping.siteCol ? dstSheet.getRange(scanStart, mapping.siteCol, numRows, 1).getValues() : null;
    var idValues = mapping.idCol ? dstSheet.getRange(scanStart, mapping.idCol, numRows, 1).getValues() : null;

    for (var i = 0; i < numRows; i++) {
      var isMatch = false;

      // 1. Match by Client / Site Name (primary for FMS tabs like uPVC PO)
      if (normSite && siteValues) {
        var cellSite = normHeader_(siteValues[i][0]);
        if (cellSite && cellSite === normSite) isMatch = true;
      }

      // 2. Match by PO ID or Quotation ID
      if (!isMatch && idValues) {
        var cellId = normHeader_(idValues[i][0]);
        if (cellId && (cellId === normId || (normQId && cellId === normQId))) isMatch = true;
      }

      if (isMatch) {
        var rIdx = scanStart + i;
        var cellRange = dstSheet.getRange(rIdx, targetCol);
        cellRange.setRichTextValue(richText);
        if (noteStr) cellRange.setNote(noteStr);
        cellRange.setWrap(true);
      }
    }
  });
}

/**
 * Bulk-synchronizes all document cells from Master sheets to all linked operational sheets.
 */
function syncAllDocLinksToAllSheets(silent) {
  var ss = getSpreadsheet_();
  var masterSheets = [SHEET.QUOTATION_MASTER, SHEET.PO_MASTER];
  var totalUpdated = 0;
  var sheetsUpdated = {};

  masterSheets.forEach(function (masterName) {
    var masterSheet = ss.getSheetByName(masterName);
    if (!masterSheet) return;
    var lastRow = masterSheet.getLastRow();
    if (lastRow < 2) return;

    var specs = columnSpecs_(masterName).filter(function (s) { return s[2] === 'doc'; });
    if (!specs.length) return;

    var idCol = (columnIndexMap_(masterName, masterSheet)[0] >= 0) ? (columnIndexMap_(masterName, masterSheet)[0] + 1) : 1;
    var siteCol = columnIndexByHeader_(masterName, 'Client / Site Name', masterSheet);
    var qIdCol = (masterName === SHEET.PO_MASTER) ? columnIndexByHeader_(masterName, 'Quotation ID', masterSheet) : 0;

    var idValues = masterSheet.getRange(2, idCol, lastRow - 1, 1).getValues();
    var siteValues = siteCol ? masterSheet.getRange(2, siteCol, lastRow - 1, 1).getValues() : null;
    var qIdValues = qIdCol ? masterSheet.getRange(2, qIdCol, lastRow - 1, 1).getValues() : null;

    var masterMapById = {};
    var masterMapBySite = {};

    specs.forEach(function (spec) {
      var col = columnIndexByHeader_(masterName, spec[0], masterSheet);
      if (!col) return;
      var range = masterSheet.getRange(2, col, lastRow - 1, 1);
      var richVals = range.getRichTextValues();
      var notes = range.getNotes();

      for (var r = 0; r < idValues.length; r++) {
        var recId = normHeader_(idValues[r][0]);
        var site = siteValues ? normHeader_(siteValues[r][0]) : '';
        var qId = qIdValues ? normHeader_(qIdValues[r][0]) : '';

        var item = {
          richText: richVals[r][0],
          note: notes[r][0]
        };

        if (recId) {
          if (!masterMapById[recId]) masterMapById[recId] = {};
          masterMapById[recId][spec[1]] = item;
        }
        if (qId) {
          if (!masterMapById[qId]) masterMapById[qId] = {};
          masterMapById[qId][spec[1]] = item;
        }
        if (site) {
          if (!masterMapBySite[site]) masterMapBySite[site] = {};
          masterMapBySite[site][spec[1]] = item;
        }
      }
    });

    var allSheets = ss.getSheets();
    allSheets.forEach(function (dstSheet) {
      var dName = dstSheet.getName();
      if (dName.toLowerCase() === masterName.toLowerCase()) return;
      if (dName.toLowerCase() === SHEET.CONFIG_DROPDOWNS.toLowerCase()) return;
      if (dName.toLowerCase() === 'setupconfig') return;
      if (dName.toLowerCase().indexOf('cancel') > -1 || dName.toLowerCase().indexOf('done') > -1) return;
      if (dName.toLowerCase() === 'summary' || dName.toLowerCase().indexOf('enquiry') > -1) return;

      var dLastRow = dstSheet.getLastRow();
      var dLastCol = dstSheet.getLastColumn();
      if (dLastRow < 2 || dLastCol < 1) return;

      var mapping = getSheetDocColumnMapping_(dstSheet, masterName);
      if (!mapping.siteCol && !mapping.idCol) return;
      if (Object.keys(mapping.docCols).length === 0) return;

      var startRow = mapping.headerRow + 1;
      var numRows = dLastRow - startRow + 1;
      if (numRows < 1) return;

      var dstSites = mapping.siteCol ? dstSheet.getRange(startRow, mapping.siteCol, numRows, 1).getValues() : null;
      var dstIds = mapping.idCol ? dstSheet.getRange(startRow, mapping.idCol, numRows, 1).getValues() : null;

      Object.keys(mapping.docCols).forEach(function (docKey) {
        var targetCol = mapping.docCols[docKey];
        for (var i = 0; i < numRows; i++) {
          var matchedData = null;

          // 1. Match by Client / Site Name
          if (dstSites) {
            var sKey = normHeader_(dstSites[i][0]);
            if (sKey && masterMapBySite[sKey] && masterMapBySite[sKey][docKey]) {
              matchedData = masterMapBySite[sKey][docKey];
            }
          }

          // 2. Match by ID
          if (!matchedData && dstIds) {
            var idKey = normHeader_(dstIds[i][0]);
            if (idKey && masterMapById[idKey] && masterMapById[idKey][docKey]) {
              matchedData = masterMapById[idKey][docKey];
            }
          }

          if (matchedData && matchedData.richText && matchedData.richText.getText().trim()) {
            var cell = dstSheet.getRange(startRow + i, targetCol);
            cell.setRichTextValue(matchedData.richText);
            if (matchedData.note) cell.setNote(matchedData.note);
            cell.setWrap(true);
            totalUpdated++;
            sheetsUpdated[dName] = (sheetsUpdated[dName] || 0) + 1;
          }
        }
      });
    });
  });

  SpreadsheetApp.flush();

  var sheetSummary = Object.keys(sheetsUpdated).map(function (k) {
    return '\u2022 ' + k + ': ' + sheetsUpdated[k] + ' cell(s)';
  }).join('\n');

  var msg = totalUpdated > 0
    ? 'Synced ' + totalUpdated + ' document cell(s) across sheet(s):\n' + sheetSummary + '\n\nAll multi-file links are now openable hyperlinks in operational sheets!'
    : 'No linked document cells required synchronization.';

  if (!silent) {
    try {
      SpreadsheetApp.getUi().alert('Document Link Sync Complete', msg, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {
      Logger.log('Sync result: ' + msg);
    }
  }

  return { success: true, totalUpdated: totalUpdated, sheetsUpdated: sheetsUpdated };
}

/* ==========================================================================
 * DEMO DATA SEEDER
 * ======================================================================== */

/**
 * Seeds realistic sample records into uPVC_Quotation_Master and uPVC_Purchase_Order_Master.
 */
function seedDemoData() {
  return withLock_(function () {
    ensureSchema_();
    var demoQuotes = getSampleQuotations_();
    demoQuotes.forEach(function (q) {
      upsertRow_(SHEET.QUOTATION_MASTER, q, q.id);
    });

    var demoPOs = getSamplePOs_();
    demoPOs.forEach(function (p) {
      upsertRow_(SHEET.PO_MASTER, p, p.id);
    });

    SpreadsheetApp.flush();
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Demo data inserted: ' + demoQuotes.length + ' Quotes & ' + demoPOs.length + ' POs',
        'uPVC FMS',
        5
      );
    } catch (e) { }

    return {
      success: true,
      quotationsCount: demoQuotes.length,
      poCount: demoPOs.length
    };
  });
}

function getSampleQuotations_() {
  return [
    {
      id: 'QT-UPVC-0001',
      site: 'Prestige Falcon City - Tower 4',
      date: '2026-10-14',
      location: 'Kanakapura Road, Bengaluru',
      vendor: 'NCL',
      variant: '3-Track Sliding',
      mesh: 'SS 304 Mesh',
      glass: '6mm, Clear, Toughened',
      qty: 14,
      leadTime: 15,
      roughSizes: [
        {
          id: 'demo_dwg_01',
          name: 'Tower4_Window_Schedule.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 348160,
          uploadedAt: '2026-10-14 09:30'
        }
      ],
      vendorQuote: [
        {
          id: 'demo_vq_01',
          name: 'NCL_Quote_Ref_9821.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 184320,
          amount: 210000,
          uploadedAt: '2026-10-15 11:20'
        }
      ],
      vendorAmount: 210000,
      ourQuote: [
        {
          id: 'demo_oq_01',
          name: 'Quotation_PFC_Tower4_Final.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 225280,
          amount: 285000,
          uploadedAt: '2026-10-15 17:00'
        }
      ],
      ourAmount: 285000
    },
    {
      id: 'QT-UPVC-0002',
      site: 'Oberoi Sky City - Penthouse B',
      date: '2026-10-15',
      location: 'Borivali East, Western Express Hwy, Mumbai',
      vendor: 'Prominance',
      variant: 'Casement / Openable, Sliding Folding',
      mesh: 'Pleated Mesh',
      glass: 'DGU (Double Glazed), 6mm, Clear, Toughened',
      qty: 22,
      leadTime: 21,
      roughSizes: [
        {
          id: 'demo_dwg_02',
          name: 'Penthouse_Glazing_Layout.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 460800,
          uploadedAt: '2026-10-15 10:15'
        }
      ],
      vendorQuote: [
        {
          id: 'demo_vq_02',
          name: 'Prominance_UPVC_Special_Fab.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 317440,
          amount: 480000,
          uploadedAt: '2026-10-16 14:00'
        }
      ],
      vendorAmount: 480000,
      ourQuote: [
        {
          id: 'demo_oq_02',
          name: 'Oberoi_SkyCity_Client_Proposal.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 266240,
          amount: 640000,
          uploadedAt: '2026-10-16 18:30'
        }
      ],
      ourAmount: 640000
    },
    {
      id: 'QT-UPVC-0003',
      site: 'Sobha Dream Acres - Villa 58',
      date: '2026-10-16',
      location: 'Panathur Road, Balagere, Bengaluru',
      vendor: 'Grayzo',
      variant: '2.5-Track Sliding',
      mesh: 'Fiberglass Mesh',
      glass: '5mm, Clear, Toughened',
      qty: 9,
      leadTime: 12,
      roughSizes: [
        {
          id: 'demo_dwg_03',
          name: 'Villa58_Site_Measurements.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 194560,
          uploadedAt: '2026-10-16 12:45'
        }
      ],
      vendorQuote: [],
      vendorAmount: 0,
      ourQuote: [],
      ourAmount: 0
    },
    {
      id: 'QT-UPVC-0004',
      site: 'Greenfield Commercial Hub - Wing B',
      date: '2026-10-17',
      location: 'Hitec City, Phase 2, Hyderabad',
      vendor: 'Alumil',
      variant: 'Fix-Light, Ventilator',
      mesh: 'None',
      glass: '10mm, Clear, Toughened',
      qty: 36,
      leadTime: 25,
      roughSizes: [
        {
          id: 'demo_dwg_04',
          name: 'Facade_Window_Specs_WingB.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 532480,
          uploadedAt: '2026-10-17 09:00'
        }
      ],
      vendorQuote: [
        {
          id: 'demo_vq_04',
          name: 'Alumil_Commercial_Costing.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 286720,
          amount: 560000,
          uploadedAt: '2026-10-17 15:30'
        }
      ],
      vendorAmount: 560000,
      ourQuote: [
        {
          id: 'demo_oq_04',
          name: 'Greenfield_Approved_Commercial_Offer.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 296960,
          amount: 720000,
          uploadedAt: '2026-10-18 11:00'
        }
      ],
      ourAmount: 720000
    },
    {
      id: 'QT-UPVC-0005',
      site: 'Palm Meadows Estate - Villa 14',
      date: '2026-10-18',
      location: 'Airport Road, Yelahanka, Bengaluru',
      vendor: 'WinFrame',
      variant: 'Door Frames, 2-Track Sliding',
      mesh: 'SS 304 Mesh',
      glass: 'Frosted / Obscure, 5mm, Clear, Toughened',
      qty: 16,
      leadTime: 14,
      roughSizes: [
        {
          id: 'demo_dwg_05',
          name: 'Villa14_Opening_Schedule.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 215040,
          uploadedAt: '2026-10-18 14:20'
        }
      ],
      vendorQuote: [
        {
          id: 'demo_vq_05',
          name: 'WinFrame_UPVC_Quotation.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 179200,
          amount: 195000,
          uploadedAt: '2026-10-18 16:50'
        }
      ],
      vendorAmount: 195000,
      ourQuote: [],
      ourAmount: 0
    }
  ];
}

function getSamplePOs_() {
  return [
    {
      id: 'PO-UPVC-0001',
      quotationId: 'QT-UPVC-0006',
      site: 'Grand Hyatt Residences - Tower A',
      date: '2026-10-02',
      location: 'Bandra Kurla Complex, Mumbai',
      vendor: 'Prominance',
      variant: '3-Track Sliding, Ventilator',
      mesh: 'SS 304 Mesh',
      glass: '6mm, Clear, Toughened, DGU (Double Glazed)',
      qty: 28,
      leadTime: 18,
      roughSizes: [
        {
          id: 'demo_po_dwg_01',
          name: 'Hyatt_TowerA_Approved_DWG.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 491520,
          uploadedAt: '2026-10-02 09:15'
        }
      ],
      vendorQuote: [
        {
          id: 'demo_po_vq_01',
          name: 'Prominance_Final_Billing_Quote.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 215040,
          amount: 520000,
          uploadedAt: '2026-10-02 12:30'
        }
      ],
      vendorAmount: 520000,
      ourQuote: [
        {
          id: 'demo_po_oq_01',
          name: 'Client_Contract_Hyatt_Signed.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 358400,
          amount: 695000,
          uploadedAt: '2026-10-03 14:00'
        }
      ],
      ourAmount: 695000,
      poOrder: [
        {
          id: 'demo_po_doc_01',
          name: 'Official_PO_HYATT_001.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 296960,
          uploadedAt: '2026-10-04 10:00'
        }
      ],
      bills: [
        { date: '2026-10-08', invoiceNo: 'INV-PRM-1044', amount: 260000 },
        { date: '2026-10-15', invoiceNo: 'INV-PRM-1089', amount: 260000 }
      ],
      billsAmt: 520000
    },
    {
      id: 'PO-UPVC-0002',
      quotationId: 'QT-UPVC-0007',
      site: 'Embassy Boulevard - Luxury Villa 19',
      date: '2026-10-06',
      location: 'Bellary Road, Yelahanka, Bengaluru',
      vendor: 'NCL',
      variant: 'Casement / Openable',
      mesh: 'Pleated Mesh',
      glass: '5mm, Clear, Toughened',
      qty: 15,
      leadTime: 14,
      roughSizes: [
        {
          id: 'demo_po_dwg_02',
          name: 'Villa19_Elevation_CAD.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 317440,
          uploadedAt: '2026-10-06 11:00'
        }
      ],
      vendorQuote: [
        {
          id: 'demo_po_vq_02',
          name: 'NCL_Revised_Quote_772.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 194560,
          amount: 240000,
          uploadedAt: '2026-10-06 15:40'
        }
      ],
      vendorAmount: 240000,
      ourQuote: [
        {
          id: 'demo_po_oq_02',
          name: 'Client_Invoice_Quote_EB19.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 209920,
          amount: 320000,
          uploadedAt: '2026-10-07 10:15'
        }
      ],
      ourAmount: 320000,
      poOrder: [
        {
          id: 'demo_po_doc_02',
          name: 'PO_NCL_EB19_Signed.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 184320,
          uploadedAt: '2026-10-07 14:00'
        }
      ],
      bills: [
        { date: '2026-10-12', invoiceNo: 'INV-NCL-992', amount: 120000 }
      ],
      billsAmt: 120000
    },
    {
      id: 'PO-UPVC-0003',
      quotationId: '',
      site: 'Brigade Gateway Commercial Block',
      date: '2026-10-10',
      location: 'Rajajinagar, Malleshwaram, Bengaluru',
      vendor: 'Alumil',
      variant: 'Sliding Folding',
      mesh: 'SS 304 Mesh',
      glass: '10mm, Clear, Toughened',
      qty: 8,
      leadTime: 20,
      roughSizes: [
        {
          id: 'demo_po_dwg_03',
          name: 'Direct_PO_Site_Specs.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 286720,
          uploadedAt: '2026-10-10 10:00'
        }
      ],
      vendorQuote: [
        {
          id: 'demo_po_vq_03',
          name: 'Alumil_Direct_Price_Agreed.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 163840,
          amount: 380000,
          uploadedAt: '2026-10-10 12:00'
        }
      ],
      vendorAmount: 380000,
      ourQuote: [
        {
          id: 'demo_po_oq_03',
          name: 'Brigade_Gate_Contract.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 225280,
          amount: 490000,
          uploadedAt: '2026-10-10 16:00'
        }
      ],
      ourAmount: 490000,
      poOrder: [
        {
          id: 'demo_po_doc_03',
          name: 'PO_BrigadeGateway_Signed.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 235520,
          uploadedAt: '2026-10-11 11:30'
        }
      ],
      bills: [],
      billsAmt: 0
    },
    {
      id: 'PO-UPVC-0004',
      quotationId: 'QT-UPVC-0008',
      site: 'Godrej Woods Phase 1',
      date: '2026-10-12',
      location: 'Sector 43, Noida, NCR',
      vendor: 'Poonam Alu.',
      variant: '2-Track Sliding',
      mesh: 'Fiberglass Mesh',
      glass: '5mm, Clear, Toughened',
      qty: 40,
      leadTime: 16,
      roughSizes: [
        {
          id: 'demo_po_dwg_04',
          name: 'Godrej_UnitPlans_AllWindows.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 634880,
          uploadedAt: '2026-10-12 09:40'
        }
      ],
      vendorQuote: [
        {
          id: 'demo_po_vq_04',
          name: 'Poonam_Bulk_Pricing.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 220160,
          amount: 410000,
          uploadedAt: '2026-10-12 14:10'
        }
      ],
      vendorAmount: 410000,
      ourQuote: [
        {
          id: 'demo_po_oq_04',
          name: 'Our_Approved_Offer_GW1.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 245760,
          amount: 535000,
          uploadedAt: '2026-10-13 11:20'
        }
      ],
      ourAmount: 535000,
      poOrder: [
        {
          id: 'demo_po_doc_04',
          name: 'Official_PO_Godrej_Woods.pdf',
          url: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          previewUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
          type: 'pdf',
          size: 199680,
          uploadedAt: '2026-10-13 15:30'
        }
      ],
      bills: [
        { date: '2026-10-16', invoiceNo: 'INV-PA-3021', amount: 205000 }
      ],
      billsAmt: 205000
    }
  ];
}