// @author Cherry Ronao - https://github.com/cvronao
/**
 * Installs the trigger for function autoBackupCOGS() in Code2.gs
 * Should NOT be run more than once. See all installed triggers (per account) on https://script.google.com/home/my
 */

function createTimeDrivenTriggers() {
  // Triggers autoBackupCOGS() function to run every Tuesday at 02:00 (plus or minus 15 minutes).   
  ScriptApp.newTrigger('autoBackupCOGS')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.TUESDAY)
    .atHour(2)
    .nearMinute(0)
    .inTimezone('Europe/Berlin')
    .create();
}