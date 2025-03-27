// @author Cherry Ronao - https://github.com/cvronao
/**
 * Backs up the whole COGS table. Applicable to both PRODUCTS and BUNDLES tabs.
 * Run this script through the "BACKUP" button.
 */

function backupCOGS() {
  SpreadsheetApp.flush();

  // Get active sheet
  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let activeSheet = activeSpreadsheet.getActiveSheet();

  if(activeSheet.getSheetName() == 'PRODUCTS - HISTORICAL'){
    Logger.log('PRODUCTS - HISTORICAL');
    const prodSourceRange = 'A3:K';
    const prodSourceSheetName = 'PRODUCTS';
    const prodSourceSheet = activeSpreadsheet.getSheetByName(prodSourceSheetName);
    const prodDateTargetColumn = 'L';
    const prodLastRunRange = 'O2';
    backupCOGSperSheet(activeSheet, prodSourceSheet, prodSourceRange, prodDateTargetColumn, prodLastRunRange);
  } else if(activeSheet.getSheetName() == 'BUNDLES - HISTORICAL') {
    Logger.log('BUNDLES - HISTORICAL');
    const bundlesSourceRange = 'A3:O';
    const bundlesSourceSheetName = 'BUNDLES';
    const bundlesSourceSheet = activeSpreadsheet.getSheetByName(bundlesSourceSheetName);
    const bundlesDateTargetColumn = 'P';
    const bundlesLastRunRange = 'S2';
    backupCOGSperSheet(activeSheet, bundlesSourceSheet, bundlesSourceRange, bundlesDateTargetColumn, bundlesLastRunRange);
  }
}

function backupCOGSperSheet(activeSheet, sourceSheet, sourceRange, dateTargetColumn, lastRunRange) {
  Logger.log('backupCOGSperSheet Function');

  // Check if script was run during the current week
  var lastGoodUpdate = activeSheet.getRange(lastRunRange).getDisplayValues().toString();
  let [y,m,d,hr,mn,ss] = lastGoodUpdate.split(/[- :]/);
  Logger.log('y: %s m: %s d: %s hr: %s mn: %s ss: %s',y,m,d,hr,mn,ss);
  lastGoodUpdate = new Date(y, m-1, d, hr, mn, ss, 0);

  let currentDateTime = new Date();
  let currentWeekMonday = getMonday(currentDateTime);
  Logger.log('Last good update: ' + lastGoodUpdate);
  Logger.log('Current week Monday: ' + currentWeekMonday);
  
  if(lastGoodUpdate.valueOf() < currentWeekMonday.valueOf() | activeSheet.getRange(lastRunRange).isBlank()) {
    // Get products table
    let source1stColumn = sourceRange.split(":")[0] + ":" + sourceRange.split(":")[0].match(/[A-Z]+/g);
    Logger.log(source1stColumn);
    let sourceLastRow = sourceSheet.getRange(source1stColumn).getNextDataCell(SpreadsheetApp.Direction.DOWN);
    let sourceLastRowNum = sourceLastRow.getA1Notation().match(/[\d]+/g);
    Logger.log(sourceLastRowNum);
    let sourceTable = sourceSheet.getRange(sourceRange + sourceLastRowNum).getValues();
    // Logger.log(sourceTable);

    // Get target range details for product conversion data
    let targetLastRow = activeSheet.getLastRow();
    let targetRange = source1stColumn.split(":")[1] + (targetLastRow + 1) + ':' + sourceRange.split(":")[1] + (targetLastRow + sourceTable.length);
    Logger.log(targetRange);

    // Paste values
    activeSheet.getRange(targetRange).setValues(sourceTable);

    // Paste current date
    let currentDateTime = new Date();
    let dateTargetRange = dateTargetColumn + targetRange.split(":")[0].match(/[\d]+/g);
    activeSheet.getRange(dateTargetRange).setValue(currentDateTime);
    activeSheet.getRange(dateTargetRange).setNumberFormat('yyy-MM-dd');
    let dateTargetRangeColumnIndex = activeSheet.getRange(dateTargetRange).getColumn();
    activeSheet.getRange(dateTargetRange).copyValuesToRange(activeSheet, dateTargetRangeColumnIndex, dateTargetRangeColumnIndex, targetLastRow + 1, targetLastRow + sourceTable.length);

    // Get current time and write it on active sheet
    currentDateTime = Utilities.formatDate(new Date(), 'GMT+1', 'yyyy-MM-dd HH:mm:ss ',);
    activeSheet.getRange(lastRunRange).setValue(currentDateTime);
  } else {
    Logger.log('Script was run already.');
  }
}

function getMonday(d) {
  d = new Date(d);
  var day = d.getDay(),
    diff = d.getDate() - day + (day == 0 ? -6 : 1); // adjust when day is sunday
  return new Date(d.setDate(diff));
}
