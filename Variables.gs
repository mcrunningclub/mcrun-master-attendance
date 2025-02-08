// SHEET CONSTANTS
const SHEET_ID = '30380045';
const MASTER_ATTENDANCE_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetById(SHEET_ID);

// LIST OF COLUMNS IN SHEET_NAME
const TIMESTAMP_COL = 1;
const HEADRUNNERS_COL = 2;
const DAY_WEEK_COL = 3;
const SCHEDULED_TIME_COL = 4;
const HEADRUN_COL = 5;
const RUN_LEVEL_COL = 6;
const ATTENDEES_COL = 7;
const CONFIRMATION_COL = 8;
const DISTANCE_COL = 9;
const COMMENTS_COL = 10;
const IS_EXPORTED_COL = 12;

const TIMEZONE = getUserTimeZone_();

const LEVEL_COUNT = 3;  // Beginner/Easy, Intermediate, Hard

// SEMESTER ATTENDANCE SHEET
const SEMESTER_ATTENDANCE = '1874229448';
const SEMESTER_ATTENDANCE_URL = 'https://docs.google.com/spreadsheets/d/1SnaD9UO4idXXb07X8EakiItOOORw5UuEOg0dX_an3T4/';


// EXTERNAL SHEETS USED IN SCRIPTS
const MASTER_NAME = 'MASTER';
const SEMESTER_NAME = 'Winter 2025';
const MEMBERSHIP_URL = "https://docs.google.com/spreadsheets/d/1qvoL3mJXCvj3m7Y70sI-FAktCiSWqEmkDxfZWz0lFu4/edit?usp=sharing";



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


/**
 * Returns email of current user executing Google Apps Script functions.
 *
 * Prevents incorrect account executing Google automations (e.g. McRUN bot.)
 *
 * @return {string}  Email of current user.
 */

function getCurrentUserEmail_() {
  return Session.getActiveUser().toString();
}
