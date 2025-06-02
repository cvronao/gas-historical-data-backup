// @author Cherry Ronao - https://github.com/cvronao
/**
 * Backs up the whole COGS table. Applicable to both PRODUCTS and BUNDLES tabs.
 * Runs automatically through a time-based trigger. Runs weekly on Tuesdays at 02:00.
 */

function autoBackupCOGS() {
  SpreadsheetApp.flush();

  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log('PRODUCTS - HISTORICAL');
  const prodSourceRange = 'A3:K';
  const prodSourceSheetName = 'PRODUCTS';
  const prodSourceSheet = activeSpreadsheet.getSheetByName(prodSourceSheetName);
  const prodTargetSheetName = 'PRODUCTS - HISTORICAL';
  const prodTargetSheet = activeSpreadsheet.getSheetByName(prodTargetSheetName);
  const prodDateTargetColumn = 'L';
  const prodLastRunRange = 'O2';
  autoBackupCOGSperSheet(prodTargetSheet, prodSourceSheet, prodSourceRange, prodDateTargetColumn, prodLastRunRange);
  
  Logger.log('BUNDLES - HISTORICAL');
  const bundlesSourceRange = 'A3:O';
  const bundlesSourceSheetName = 'BUNDLES';
  const bundlesSourceSheet = activeSpreadsheet.getSheetByName(bundlesSourceSheetName);
  const bundlesTargetSheetName = 'BUNDLES - HISTORICAL';
  const bundlesTargetSheet = activeSpreadsheet.getSheetByName(bundlesTargetSheetName);
  const bundlesDateTargetColumn = 'P';
  const bundlesLastRunRange = 'S2';
  autoBackupCOGSperSheet(bundlesTargetSheet, bundlesSourceSheet, bundlesSourceRange, bundlesDateTargetColumn, bundlesLastRunRange);
}

function autoBackupCOGSperSheet(targetSheet, sourceSheet, sourceRange, dateTargetColumn, lastRunRange) {
  Logger.log('backupCOGSperSheet Function');

  // Get source table
  let source1stColumn = sourceRange.split(":")[0] + ":" + sourceRange.split(":")[0].match(/[A-Z]+/g);
  Logger.log(source1stColumn);
  let sourceLastRow = sourceSheet.getRange(source1stColumn).getNextDataCell(SpreadsheetApp.Direction.DOWN);
  let sourceLastRowNum = sourceLastRow.getA1Notation().match(/[\d]+/g);
  Logger.log(sourceLastRowNum);
  let sourceTable = sourceSheet.getRange(sourceRange + sourceLastRowNum).getValues();
  // Logger.log(sourceTable);

  // Get target range details
  let targetLastRow = targetSheet.getLastRow();
  let targetRange = source1stColumn.split(":")[1] + (targetLastRow + 1) + ':' + sourceRange.split(":")[1] + (targetLastRow + sourceTable.length);
  Logger.log(targetRange);

  // Paste values
  targetSheet.getRange(targetRange).setValues(sourceTable);

  // Paste current date
  let currentDateTime = new Date();
  let dateTargetRange = dateTargetColumn + targetRange.split(":")[0].match(/[\d]+/g);
  targetSheet.getRange(dateTargetRange).setValue(currentDateTime);
  targetSheet.getRange(dateTargetRange).setNumberFormat('yyy-MM-dd');
  let dateTargetRangeColumnIndex = targetSheet.getRange(dateTargetRange).getColumn();
  targetSheet.getRange(dateTargetRange).copyValuesToRange(targetSheet, dateTargetRangeColumnIndex, dateTargetRangeColumnIndex, targetLastRow + 1, targetLastRow + sourceTable.length);

  // Get current time and write it on active sheet
  currentDateTime = Utilities.formatDate(new Date(), 'GMT+1', 'yyyy-MM-dd HH:mm:ss ',);
  targetSheet.getRange(lastRunRange).setValue(currentDateTime);
}

function getMonday(d) {
  d = new Date(d);
  var day = d.getDay(),
    diff = d.getDate() - day + (day == 0 ? -6 : 1); // adjust when day is sunday
  return new Date(d.setDate(diff));
}
