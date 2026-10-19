/**
 * ============================================================================
 *  BMP / FMS FORMULA HELPER & ACTUAL TIME TRIGGER (Code.js)
 * ============================================================================
 */

function buildBmpMenu_(ui) {
  var menuUi = ui || SpreadsheetApp.getUi();
  menuUi.createMenu('BMP Formulas')
    .addItem('Run This First Time After Install', 'initial')
    .addSeparator()
    .addItem('Assisted ImportRange', 'importRangeFormula')
    .addSeparator()
    .addSubMenu(menuUi.createMenu('FMS Formulas')
      .addItem('TAT in Hours (Inside Working Hours)', 'plannedwwh')
      .addSeparator()
      .addItem('TAT in Hours (Outside Working Hours)', 'plannedwwh2')
      .addSeparator()
      .addItem('TAT in Days', 'plannedindays')
      .addSeparator()
      .addItem('T-x Formula', 'plannedlead')
      .addSeparator()
      .addItem('Specific Time', 'specificTime')
      .addSeparator()
      .addItem('Show Planned Only When Status is NO', 'tatifno')
      .addSeparator()
      .addItem('Show Planned Only When Status is YES', 'tatifyes')
      .addSeparator()
      .addItem('Set Actual Time (Formula)', 'actualTime')
      .addSeparator()
      .addItem('Set Actual Time (Script Trigger)', 'createTrigger')
      .addSeparator()
      .addItem('Time Delay Formula', 'timeDelay'))
    .addToUi();

  if (typeof generatePendingSummary === 'function') {
    try { generatePendingSummary(); } catch (e) { }
  }
}

function createTrigger() {
  removeTrigger();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onChange_new').forSpreadsheet(ss).onChange().create();
}

function removeTrigger() {
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getHandlerFunction() === 'onChange_new') {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }
}

function importRangeFormula() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var donorResp = ui.prompt('URL of the spreadsheet from where data has to be imported', ui.ButtonSet.OK_CANCEL);
  if (donorResp.getSelectedButton() !== ui.Button.OK) return;
  var donorSheet = donorResp.getResponseText();

  var tabResp = ui.prompt('Name of the sheet/tab from where data has to be imported', ui.ButtonSet.OK_CANCEL);
  if (tabResp.getSelectedButton() !== ui.Button.OK) return;
  var tabName = tabResp.getResponseText();

  var rangeResp = ui.prompt('Enter the range from where data has to be imported', ui.ButtonSet.OK_CANCEL);
  if (rangeResp.getSelectedButton() !== ui.Button.OK) return;
  var rangeName = rangeResp.getResponseText();

  ss.getCurrentCell().setFormula('=IMPORTRANGE("' + donorSheet + '","' + tabName + '!' + rangeName + '")');
}

function initial() {
  var ui = SpreadsheetApp.getUi();
  var openResp = ui.prompt('Opening Time (Hour in 24h format, e.g. 9)', ui.ButtonSet.OK_CANCEL);
  if (openResp.getSelectedButton() !== ui.Button.OK) return;
  var openingTime = openResp.getResponseText();

  var closeResp = ui.prompt('Closing Time (Hour in 24h format, e.g. 18)', ui.ButtonSet.OK_CANCEL);
  if (closeResp.getSelectedButton() !== ui.Button.OK) return;
  var closingTime = closeResp.getResponseText();

  iterative();
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange('A1').setFormula('=NOW()');
  sheet.getRange('C1').setValue(openingTime);
  sheet.getRange('D1').setValue(closingTime);
  sheet.getRange('B1').setFormula('=' + closingTime + '/24-' + openingTime + '/24');
  sheet.hideRows(1, 1);
  sheet.getRange('A2').activate();
}

function copyFormulaDown_(ss) {
  var currentCell = ss.getCurrentCell();
  ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  ss.getCurrentCell().copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
}

function plannedwwh() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var fromResp = ui.prompt('Date Cell in which you want to add TAT', ui.ButtonSet.OK_CANCEL);
  if (fromResp.getSelectedButton() !== ui.Button.OK) return;
  var fromDate = fromResp.getResponseText();

  var tatResp = ui.prompt('TAT cell (example G$5)', ui.ButtonSet.OK_CANCEL);
  if (tatResp.getSelectedButton() !== ui.Button.OK) return;
  var tatInHours = tatResp.getResponseText();

  ss.getCurrentCell().setFormula(
    '=IF(' + fromDate + '<>"",IFS(HOUR(' + fromDate + '+' + tatInHours + ')>$D$1,WORKDAY.INTL(' + fromDate + ',1,"0000001")+$C$1/24+' + tatInHours +
    ',HOUR(' + fromDate + '+' + tatInHours + ')<$C$1,DATEVALUE(' + fromDate + ')+$C$1/24+' + tatInHours +
    ',AND(HOUR(' + fromDate + '+' + tatInHours + ')>=$C$1,HOUR(' + fromDate + '+' + tatInHours + ')<=$D$1),' + fromDate + '+' + tatInHours + '),"")'
  );
  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
  copyFormulaDown_(ss);
}

function plannedwwh2() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var ofs = ss.getRange('C1').getValue();
  var ofe = ss.getRange('D1').getValue();

  var fromResp = ui.prompt('Date Cell in which you want to add TAT', ui.ButtonSet.OK_CANCEL);
  if (fromResp.getSelectedButton() !== ui.Button.OK) return;
  var fromDate = fromResp.getResponseText();

  var tatResp = ui.prompt('TAT cell (example G$5)', ui.ButtonSet.OK_CANCEL);
  if (tatResp.getSelectedButton() !== ui.Button.OK) return;
  var tatInHours = tatResp.getResponseText();

  ss.getCurrentCell().setFormula(
    '=IF(' + fromDate + ',IF(AND(HOUR(' + fromDate + '+' + tatInHours + ')>' + ofs + ',(HOUR(' + fromDate + '+' + tatInHours + ')<' + ofe + ')),' +
    fromDate + '+' + tatInHours + ',WORKDAY.INTL(INT(' + fromDate + '),1,"0000001",Holidays!A:A)+HOUR(' + fromDate + '+' + tatInHours + '-$B$1)/24+MINUTE(' + fromDate + ')/1440),"")'
  );
  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
  copyFormulaDown_(ss);
}

function plannedindays() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var fromResp = ui.prompt('Date Cell in which you want to add TAT', ui.ButtonSet.OK_CANCEL);
  if (fromResp.getSelectedButton() !== ui.Button.OK) return;
  var fromDate = fromResp.getResponseText();

  var tatResp = ui.prompt('TAT cell (example G$5)', ui.ButtonSet.OK_CANCEL);
  if (tatResp.getSelectedButton() !== ui.Button.OK) return;
  var tatInDays = tatResp.getResponseText();

  ss.getCurrentCell().setFormula(
    '=IF(' + fromDate + ',WORKDAY.INTL(' + fromDate + ',' + tatInDays + ',"0000001",Holidays!A:A)+HOUR(' + fromDate + ')/24+MINUTE(' + fromDate + ')/1440,"")'
  );
  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
  copyFormulaDown_(ss);
}

function plannedlead() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var fromResp = ui.prompt('Date Cell in which you want to add lead time', ui.ButtonSet.OK_CANCEL);
  if (fromResp.getSelectedButton() !== ui.Button.OK) return;
  var fromDate = fromResp.getResponseText();

  var leadResp = ui.prompt('Lead Time Cell', ui.ButtonSet.OK_CANCEL);
  if (leadResp.getSelectedButton() !== ui.Button.OK) return;
  var leadtime = leadResp.getResponseText();

  var daysResp = ui.prompt('Number of Days Before Lead Time', ui.ButtonSet.OK_CANCEL);
  if (daysResp.getSelectedButton() !== ui.Button.OK) return;
  var daysBefore = daysResp.getResponseText();

  ss.getCurrentCell().setFormula('=IF(' + leadtime + ',' + fromDate + '+' + leadtime + '-' + daysBefore + ',"")');
  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
  copyFormulaDown_(ss);
}

function specificTime() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var fromResp = ui.prompt('Date Cell', ui.ButtonSet.OK_CANCEL);
  if (fromResp.getSelectedButton() !== ui.Button.OK) return;
  var fromDate = fromResp.getResponseText();

  var leadResp = ui.prompt('Number of Days after previous planned (Write 0 if same day)', ui.ButtonSet.OK_CANCEL);
  if (leadResp.getSelectedButton() !== ui.Button.OK) return;
  var leadtime = leadResp.getResponseText();

  var timeResp = ui.prompt('Time of day in hour/24 format', ui.ButtonSet.OK_CANCEL);
  if (timeResp.getSelectedButton() !== ui.Button.OK) return;
  var timeOfDay = timeResp.getResponseText();

  ss.getCurrentCell().setFormula('=IF(' + fromDate + ',WORKDAY.INTL(INT(' + fromDate + '),' + leadtime + ',"0000001",Holidays!A:A)+' + timeOfDay + ',"")');
  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
  copyFormulaDown_(ss);
}

function actualTime() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var statusResp = ui.prompt('Status Cell', ui.ButtonSet.OK_CANCEL);
  if (statusResp.getSelectedButton() !== ui.Button.OK) return;
  var statusCell = statusResp.getResponseText();

  var currcella1 = ss.getCurrentCell().getA1Notation();
  ss.getCurrentCell().setFormula('=IF(' + currcella1 + ',' + currcella1 + ',IF(' + statusCell + '<>"",$A$1,""))');
  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
  copyFormulaDown_(ss);
}

function timeDelay() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var planResp = ui.prompt('Planned Cell', ui.ButtonSet.OK_CANCEL);
  if (planResp.getSelectedButton() !== ui.Button.OK) return;
  var plannedCell = planResp.getResponseText();

  var actResp = ui.prompt('Actual Cell', ui.ButtonSet.OK_CANCEL);
  if (actResp.getSelectedButton() !== ui.Button.OK) return;
  var actualCell = actResp.getResponseText();

  ss.getCurrentCell().setFormula(
    '=IF(' + plannedCell + ',IF(' + actualCell + '<>"",IF(' + actualCell + '>' + plannedCell + ',' + actualCell + '-' + plannedCell + ',""),$A$1-' + plannedCell + '),"")'
  );
  ss.getActiveRangeList().setNumberFormat('[h]:mm:ss');

  var sheet = ss.getActiveSheet();
  var rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges([ss.getActiveRange()])
    .whenFormulaSatisfied('=IF(' + actualCell + ',IF(' + actualCell + '>' + plannedCell + ',1,0),0)')
    .setBackground('#F4C7C3')
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges([ss.getActiveRange()])
    .whenFormulaSatisfied('=IF(' + actualCell + ',0,IF(' + plannedCell + '<$A$1,1,0))')
    .setBackground('#FCE8B2')
    .build());
  sheet.setConditionalFormatRules(rules);

  copyFormulaDown_(ss);
}

function tatifno() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var statusResp = ui.prompt('Status Cell', ui.ButtonSet.OK_CANCEL);
  if (statusResp.getSelectedButton() !== ui.Button.OK) return;
  var statusCell = statusResp.getResponseText();

  var formulaincell = ss.getCurrentCell().getFormula().substr(1);
  ss.getCurrentCell().setFormula('=IF(' + statusCell + '="No",' + formulaincell + ',"")');
  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
  copyFormulaDown_(ss);
}

function tatifyes() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var statusResp = ui.prompt('Status Cell', ui.ButtonSet.OK_CANCEL);
  if (statusResp.getSelectedButton() !== ui.Button.OK) return;
  var statusCell = statusResp.getResponseText();

  var formulaincell = ss.getCurrentCell().getFormula().substr(1);
  ss.getCurrentCell().setFormula('=IF(' + statusCell + '="Yes",' + formulaincell + ',"")');
  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
  copyFormulaDown_(ss);
}

function iterative() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setRecalculationInterval(SpreadsheetApp.RecalculationInterval.ON_CHANGE);
  spreadsheet.setIterativeCalculationEnabled(true);
  spreadsheet.setMaxIterativeCalculationCycles(1);
  spreadsheet.setIterativeCalculationConvergenceThreshold(0.05);
}

function onChange_new() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var active = ss.getActiveCell();
  var row = active.getRow();
  var col = active.getColumn();
  if (col <= 1) return;
  var val = active.getValue();

  if (val === 'Done' || val === 'Yes' || val === 'No' || val === true) {
    var tsCell = sheet.getRange(row, col - 1);
    if (tsCell.getValue() === '') {
      tsCell.setValue(new Date());
    }
  }
}