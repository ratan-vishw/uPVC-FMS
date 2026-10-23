/**
 * ============================================================================
 *  AUTO-CONFIGURING ARCHIVING SYSTEM (Archive.js)
 *  - Automatically creates and populates the "SetupConfig" sheet if missing.
 *  - Dynamically detects whether "Status" headers sit in Row 6 (FMS tabs)
 *    or Row 1 (Master tabs).
 *  - Live-scans all "Status" columns on every run (never relies on stale letters).
 *  - Moves rows with ANY Status = "cancel" -> Cancel Sheet.
 *  - Moves rows with LAST Status = "yes" / "done" -> Done Sheet.
 *  - Preserves RichText clickable hyperlinks and Cell Notes during archiving.
 * ============================================================================
 */

const ARCHIVE_SETUP_SHEET = 'SetupConfig';
const DEFAULT_FMS_HEADER_ROW = 6;

const ARCHIVE_SETUP_HEADERS = [
    'Sheet Name',
    'Cancel Columns',
    'Completion Columns',
    'Response Sheet',
    'Response Row Offset',
    'Cancel Sheet',
    'Done Sheet'
];

/**
 * Default sheet mappings auto-written to SetupConfig when initialized.
 */
const DEFAULT_ARCHIVE_RULES = [
    ['Quotation', '', '', 'Quotation Form Responses', 5, 'Quotation Cancel', 'Quotation Done'],
    ['Purchase Orders', '', '', 'PO Form Responses', 5, 'PO Cancel', 'PO Done'],
    ['uPVC_Quotation_Master', '', '', '', 0, 'uPVC_Quote_Cancel', 'uPVC_Quote_Done'],
    ['uPVC_Purchase_Order_Master', '', '', '', 0, 'uPVC_PO_Cancel', 'uPVC_PO_Done']
];

/**
 * Called automatically by onOpen() in Web.js, or standalone if run directly.
 */
function buildArchivingMenu_(ui) {
    var menuUi = ui || SpreadsheetApp.getUi();
    menuUi.createMenu('Archiving Tools')
        .addItem('\uD83D\uDCE6 Archive Now', 'archiveAllRows')
        .addSeparator()
        .addItem('\u2699\uFE0F Auto-Configure / Repair SetupConfig', 'autoConfigureArchiveSetup')
        .addItem('\uD83D\uDD04 Refresh Detected Status Columns', 'refreshSetupConfigColumns')
        .addToUi();
}

/**
 * Automatically creates SetupConfig (if missing) and populates default rules
 * for existing sheets in the workbook, then refreshes live Status column letters.
 */
function autoConfigureArchiveSetup(silent) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var setupSheet = ss.getSheetByName(ARCHIVE_SETUP_SHEET);

    if (!setupSheet) {
        setupSheet = ss.insertSheet(ARCHIVE_SETUP_SHEET);
    }

    // Ensure Header Row exists in SetupConfig
    var lastRow = setupSheet.getLastRow();
    if (lastRow < 1) {
        setupSheet.getRange(1, 1, 1, ARCHIVE_SETUP_HEADERS.length).setValues([ARCHIVE_SETUP_HEADERS]);
    } else {
        var currentHeaders = setupSheet.getRange(1, 1, 1, ARCHIVE_SETUP_HEADERS.length).getValues()[0];
        var isBlankHeader = currentHeaders.every(function (h) { return !String(h || '').trim(); });
        if (isBlankHeader) {
            setupSheet.getRange(1, 1, 1, ARCHIVE_SETUP_HEADERS.length).setValues([ARCHIVE_SETUP_HEADERS]);
        }
    }

    // Style SetupConfig header
    setupSheet.getRange(1, 1, 1, ARCHIVE_SETUP_HEADERS.length)
        .setFontWeight('bold')
        .setBackground('#0F172A')
        .setFontColor('#FFFFFF')
        .setHorizontalAlignment('center');
    setupSheet.setFrozenRows(1);

    // Read existing configured sheet names so we never duplicate rows
    lastRow = setupSheet.getLastRow();
    var existingNames = {};
    if (lastRow >= 2) {
        var existingColA = setupSheet.getRange(2, 1, lastRow - 1, 1).getValues();
        existingColA.forEach(function (r) {
            var name = String(r[0] || '').trim();
            if (name) existingNames[name.toLowerCase()] = true;
        });
    }

    // Add default rules for sheets that exist in the spreadsheet (or default FMS sheets if completely empty)
    var rowsToAdd = [];
    DEFAULT_ARCHIVE_RULES.forEach(function (rule) {
        var sheetName = rule[0];
        if (existingNames[sheetName.toLowerCase()]) return;
        if (ss.getSheetByName(sheetName) || (lastRow < 2 && (sheetName === 'Quotation' || sheetName === 'Purchase Orders'))) {
            rowsToAdd.push(rule);
            existingNames[sheetName.toLowerCase()] = true;
        }
    });

    if (rowsToAdd.length > 0) {
        var startRow = Math.max(setupSheet.getLastRow() + 1, 2);
        setupSheet.getRange(startRow, 1, rowsToAdd.length, ARCHIVE_SETUP_HEADERS.length).setValues(rowsToAdd);
    }

    setupSheet.setColumnWidths(1, ARCHIVE_SETUP_HEADERS.length, 175);
    refreshSetupConfigColumns(true);

    if (!silent) {
        try {
            SpreadsheetApp.getUi().alert('SetupConfig Ready',
                '"' + ARCHIVE_SETUP_SHEET + '" has been auto-configured and synced with live Status columns.',
                SpreadsheetApp.getUi().ButtonSet.OK);
        } catch (e) { }
    }
    return setupSheet;
}

/**
 * Main Archiving Function — run from "Archiving Tools > Archive Now"
 */
function archiveAllRows() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var config = loadSheetConfig_(ss);
    if (!config || Object.keys(config).length === 0) {
        autoConfigureArchiveSetup(true);
        config = loadSheetConfig_(ss);
    }
    if (!config) return;

    var totalDone = 0;
    var totalCancelled = 0;

    Object.keys(config).forEach(function (sheetName) {
        var cfg = config[sheetName];
        var sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            Logger.log('Sheet "' + sheetName + '" not found, skipping.');
            return;
        }

        var scan = detectStatusHeaderAndColumns_(sheet);
        var headerRow = scan.headerRow;
        var statusCols = scan.statusCols;

        if (statusCols.length === 0) {
            Logger.log('No "Status" columns found in "' + sheetName + '", skipping.');
            return;
        }

        var completionCol = statusCols[statusCols.length - 1];
        var lastRow = sheet.getLastRow();
        var lastCol = sheet.getLastColumn();
        if (lastRow <= headerRow) return;

        var numDataRows = lastRow - headerRow;
        var dataRange = sheet.getRange(headerRow + 1, 1, numDataRows, lastCol);
        var values = dataRange.getValues();
        var richValues = dataRange.getRichTextValues();
        var notes = dataRange.getNotes();

        var doneBundle = [];
        var cancelBundle = [];
        var rowsToDelete = [];

        values.forEach(function (rowData, idx) {
            var rowNumber = headerRow + 1 + idx;
            var valueAt = function (col1Based) {
                var v = rowData[col1Based - 1];
                return typeof v === 'string' ? v.trim().toLowerCase() : String(v || '').trim().toLowerCase();
            };

            var isCancelled = statusCols.some(function (c) {
                var val = valueAt(c);
                return val === 'cancel' || val === 'cancelled';
            });
            var finalVal = valueAt(completionCol);
            var isDone = (finalVal === 'yes' || finalVal === 'done' || finalVal === 'completed');

            if (isCancelled) {
                cancelBundle.push({ values: rowData, rich: richValues[idx], notes: notes[idx] });
                rowsToDelete.push(rowNumber);
            } else if (isDone) {
                doneBundle.push({ values: rowData, rich: richValues[idx], notes: notes[idx] });
                rowsToDelete.push(rowNumber);
            }
        });

        if (doneBundle.length > 0 && cfg.doneSheet) {
            appendArchivedRows_(ss, cfg.doneSheet, sheet, headerRow, doneBundle);
        }
        if (cancelBundle.length > 0 && cfg.cancelSheet) {
            appendArchivedRows_(ss, cfg.cancelSheet, sheet, headerRow, cancelBundle);
        }

        // Delete matched rows bottom-to-top from the tracked sheet
        rowsToDelete.sort(function (a, b) { return b - a; }).forEach(function (r) {
            sheet.deleteRow(r);
        });

        // Clean up corresponding rows on the linked Form Responses sheet (if configured)
        if (cfg.responseSheet && rowsToDelete.length > 0) {
            var responsesSheet = ss.getSheetByName(cfg.responseSheet);
            if (responsesSheet) {
                rowsToDelete
                    .map(function (r) { return r - cfg.responseRowOffset; })
                    .filter(function (r) { return r > 1; })
                    .sort(function (a, b) { return b - a; })
                    .forEach(function (r) {
                        if (r <= responsesSheet.getLastRow()) {
                            responsesSheet.deleteRow(r);
                        }
                    });
            }
        }

        totalDone += doneBundle.length;
        totalCancelled += cancelBundle.length;
    });

    refreshSetupConfigColumns(true);

    if (typeof generatePendingSummary === 'function') {
        try { generatePendingSummary(); } catch (e) { Logger.log('generatePendingSummary error: ' + e); }
    }

    var msg = 'Archiving complete:\n\u2022 ' + totalDone + ' moved to Done\n\u2022 ' + totalCancelled + ' moved to Cancel';
    Logger.log(msg);
    try {
        SpreadsheetApp.getUi().alert('Archive Complete', msg, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (err) { }
}

/**
 * Automatically detects whether "Status" headers are on Row 6 (FMS style)
 * or Row 1 (Master table style), and returns 1-based column indices.
 */
function detectStatusHeaderAndColumns_(sheet) {
    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    if (lastCol < 1 || lastRow < 1) return { headerRow: 1, statusCols: [] };

    var candidateRows = [DEFAULT_FMS_HEADER_ROW, 1];
    for (var c = 0; c < candidateRows.length; c++) {
        var rowNum = candidateRows[c];
        if (lastRow < rowNum) continue;
        var rowVals = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
        var cols = [];
        rowVals.forEach(function (val, i) {
            var clean = String(val === null || val === undefined ? '' : val).trim().toLowerCase();
            if (clean === 'status' || clean === 'archive status' || clean === 'order status') {
                cols.push(i + 1);
            }
        });
        if (cols.length > 0) {
            return { headerRow: rowNum, statusCols: cols };
        }
    }

    return { headerRow: DEFAULT_FMS_HEADER_ROW, statusCols: [] };
}

/**
 * Appends archived rows to the destination sheet while preserving RichText links,
 * Cell Notes, and appending an "Archived Date" timestamp column.
 */
function appendArchivedRows_(ss, destSheetName, sourceSheet, headerRow, bundleList) {
    var destSheet = ss.getSheetByName(destSheetName);
    if (!destSheet) {
        destSheet = createArchiveDestSheet_(ss, destSheetName, sourceSheet, headerRow);
    }

    var timestamp = new Date();
    var numRows = bundleList.length;
    var sourceCols = bundleList[0].values.length;
    var startRow = Math.max(destSheet.getLastRow() + 1, headerRow + 1);

    var valuesMatrix = bundleList.map(function (b) {
        return b.values.concat([timestamp]);
    });
    var notesMatrix = bundleList.map(function (b) {
        return b.notes.concat(['']);
    });

    destSheet.getRange(startRow, 1, numRows, sourceCols + 1).setValues(valuesMatrix);
    destSheet.getRange(startRow, 1, numRows, sourceCols + 1).setNotes(notesMatrix);

    // Restore clickable RichText links on cells that contain hyperlinks
    for (var r = 0; r < numRows; r++) {
        var richRow = bundleList[r].rich;
        for (var c = 0; c < sourceCols; c++) {
            var richVal = richRow[c];
            if (richVal && hasRichLinks_(richVal)) {
                destSheet.getRange(startRow + r, c + 1).setRichTextValue(richVal);
            }
        }
    }
}

function hasRichLinks_(richText) {
    if (!richText || typeof richText.getRuns !== 'function') return false;
    var runs = richText.getRuns();
    for (var i = 0; i < runs.length; i++) {
        if (runs[i].getLinkUrl()) return true;
    }
    return false;
}

function createArchiveDestSheet_(ss, sheetName, sourceSheet, headerRow) {
    var newSheet = ss.insertSheet(sheetName);
    var lastCol = Math.max(sourceSheet.getLastColumn(), 1);
    var copyRows = Math.max(headerRow, 1);

    sourceSheet.getRange(1, 1, copyRows, lastCol).copyTo(newSheet.getRange(1, 1, copyRows, lastCol));
    newSheet.getRange(copyRows, lastCol + 1).setValue('Archived Date').setFontWeight('bold');
    newSheet.setFrozenRows(copyRows);
    return newSheet;
}

/**
 * Reads SetupConfig dynamically by header name (creates SetupConfig if missing).
 */
function loadSheetConfig_(ss) {
    var setupSheet = ss.getSheetByName(ARCHIVE_SETUP_SHEET);
    if (!setupSheet) {
        setupSheet = autoConfigureArchiveSetup(true);
    }

    var lastRow = setupSheet.getLastRow();
    var lastCol = Math.max(setupSheet.getLastColumn(), 7);
    if (lastRow < 2) return {};

    var headers = setupSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
        return String(h || '').trim().toLowerCase();
    });

    var idxSheet = headers.indexOf('sheet name') > -1 ? headers.indexOf('sheet name') : 0;
    var idxResp = headers.indexOf('response sheet') > -1 ? headers.indexOf('response sheet') : 3;
    var idxOffset = headers.indexOf('response row offset') > -1 ? headers.indexOf('response row offset') : 4;
    var idxCancel = headers.indexOf('cancel sheet') > -1 ? headers.indexOf('cancel sheet') : 5;
    var idxDone = headers.indexOf('done sheet') > -1 ? headers.indexOf('done sheet') : 6;

    var rows = setupSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var config = {};

    rows.forEach(function (row) {
        var sheetName = String(row[idxSheet] || '').trim();
        if (!sheetName) return;
        config[sheetName] = {
            responseSheet: String(row[idxResp] || '').trim(),
            responseRowOffset: Number(row[idxOffset]) || 0,
            cancelSheet: String(row[idxCancel] || '').trim() || (sheetName + ' Cancel'),
            doneSheet: String(row[idxDone] || '').trim() || (sheetName + ' Done')
        };
    });

    return config;
}

/**
 * Updates Columns B & C in SetupConfig with the live-detected Status column letters.
 */
function refreshSetupConfigColumns(silent) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var setupSheet = ss.getSheetByName(ARCHIVE_SETUP_SHEET);
    if (!setupSheet) {
        setupSheet = autoConfigureArchiveSetup(true);
    }

    var lastRow = setupSheet.getLastRow();
    if (lastRow < 2) return;

    for (var r = 2; r <= lastRow; r++) {
        var sheetName = String(setupSheet.getRange(r, 1).getValue() || '').trim();
        if (!sheetName) continue;
        var sheet = ss.getSheetByName(sheetName);
        if (!sheet) continue;

        var scan = detectStatusHeaderAndColumns_(sheet);
        var statusLetters = scan.statusCols.map(columnToLetter_);
        if (statusLetters.length === 0) continue;

        setupSheet.getRange(r, 2).setValue(statusLetters.join(','));
        setupSheet.getRange(r, 3).setValue(statusLetters[statusLetters.length - 1]);
    }

    if (!silent) {
        try {
            SpreadsheetApp.getUi().alert('Status Columns Refreshed',
                'SetupConfig columns B & C have been updated with live Status column letters.',
                SpreadsheetApp.getUi().ButtonSet.OK);
        } catch (e) { }
    }
}

function columnToLetter_(column) {
    var letter = '';
    while (column > 0) {
        var mod = (column - 1) % 26;
        letter = String.fromCharCode(65 + mod) + letter;
        column = Math.floor((column - mod) / 26);
    }
    return letter;
}