// SHEET CONSTANTS
const MASTER_ATTENDANCE_SHEET_ID = '30380045';
const MASTER_ATTENDANCE_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetById(MASTER_ATTENDANCE_SHEET_ID);

const COLUMN_MAP = {
  TIMESTAMP: 1,
  HEADRUNNERS: 2,
  DAY_WEEK: 3,
  SCHEDULED_TIME: 4,
  HEADRUN: 5,
  RUN_LEVEL: 6,
  ATTENDEES: 7,
  CONFIRMATION: 8,
  DISTANCE: 9,
  COMMENTS: 10,
  IS_EXPORTED: 11,
}

const TIMEZONE = getUserTimeZone_();


/**
 * Returns timezone for currently running script.
 *
 * Prevents incorrect time formatting during time changes like Daylight Savings Time.
 *
 * @return {string}  Timezone as geographical location (e.g.`'America/Montreal'`).
 */

function getUserTimeZone_() {
  return Session.getScriptTimeZone();
}
