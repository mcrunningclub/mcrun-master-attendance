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
 * Returns the timezone for the currently running script as a geographical location string.
 *
 * This function ensures that all date and time formatting operations use the correct timezone,
 * preventing issues such as incorrect time display during Daylight Savings Time transitions.
 *
 * @returns {string} The timezone in IANA format (e.g., 'America/Montreal').
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date Feb 8, 2025
 */

function getUserTimeZone_() {
  return Session.getScriptTimeZone();
}
