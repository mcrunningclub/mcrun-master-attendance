// SEMESTER ATTENDANCE SHEET
const ATTENDANCE_IMPORT_ID = '82376152';
const SEMESTER_ATTENDANCE_URL = 'https://docs.google.com/spreadsheets/d/1SnaD9UO4idXXb07X8EakiItOOORw5UuEOg0dX_an3T4/';


function onChange(e) {
  const thisSource = e.source;
  const thisChange = e.changeType;

  // Verify if thisSource valid
  if (!thisSource) {
    console.log(`thisSource is not defined. Value: ${thisSource}`);
    return;
  }

  const thisSheetID = thisSource.getSheetId();

  // Exit early if the event is not related to the import sheet
  if (thisChange !== 'INSERT_ROW' || thisSheetID != MASTER_ATTENDANCE_SHEET_ID) {
    console.log(`
      Early exit. Either e.changeType or source.sheetId() not as expected.
      Type of change: ${thisChange} \tExpected: INSERT_ROW
      thisSheetID: ${thisSheetID} \tExpected: ${MASTER_ATTENDANCE_SHEET_ID}`
    );

    return;
  }

  // Trigger formatting and transfer functions if new submission
  transferToSemesterSheet();
  formatAllNamesInRow();

  /* try {
    thisSource = e.source;
    const thisChange = e.changeType;
    const thisSheetID = thisSource.getSheetId();

    // Exit early if the event is not related to the import sheet
    if (thisChange !== 'INSERT_ROW' || thisSheetID != MASTER_ATTENDANCE_SHEET_ID) {
      console.log(`
        Early exit. Either e.changeType or source.sheetId() not as expected.
        Type of change: ${thisChange}. Expected 'INSERT_ROW'
        thisSheetID: ${thisSheetID}. Expected ${MASTER_ATTENDANCE_SHEET_ID}`
      );

      return;
    }

    // Trigger formatting and transfer functions if new submission
    transferToSemesterSheet();
    formatAllNamesInRow();
  }
  catch (error) {
    if (thisSource) {
      console.log(`thisSource is not defined. Value: ${thisSource}`);
      return;
    }

    const errMsg = `(onChange): ${error}\n${e}`;
    throw new Error(errMsg);
  } */
}


/**
 * Transfers attendance submission to semester attendance sheet.
 * 
 * Exports using `openByUrl` and creating JSON object.
 * 
 * @trigger New app submission.
 * 
 * @param {integer} [row=getLastSubmission_()] row  Row index in GSheet.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Feb 8, 2025
 * @update  Feb 8, 2025
 */

function transferToSemesterSheet(row=getLastSubmission_()-1) {
  const sheet = MASTER_ATTENDANCE_SHEET;
  const sourceRow = row;
  const sourceColSize = sheet.getLastColumn();

  const rangeSource = sheet.getRange(sourceRow, 1, 1, sourceColSize);
  const values = rangeSource.getValues()[0];  // Get submission row

  // Prepare registration data to export
  const exportJson = prepareAttendanceSubmission(values);

  // `Memberships Collected (Main)` GSheet
  const sheetURL = SEMESTER_ATTENDANCE_URL;
  const ss = SpreadsheetApp.openByUrl(sheetURL);
  const importSheet = ss.getSheetById(ATTENDANCE_IMPORT_ID);
   
  // Export registration to `Import` sheet
  const newRow = importSheet.getLastRow();
  const rangeNewImport = importSheet.getRange(newRow, 1);
  rangeNewImport.setValue(exportJson);

  // This triggers the `onChange(e)` function
  importSheet.insertRowAfter(newRow);
}

/**
 * Prepare the attendance values into JSON object.
 * 
 * @param {String[]} values  Run attendance information.
 * 
 * @return {String<JSON>}  Converts `values` to a JSON string.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Feb 8, 2025
 * @update  Feb 8, 2025
 */

function prepareAttendanceSubmission(values) {

  /** -> CURRENT SEMESTER ATTENDANCE INDICES (1-indexed) <-
   * 
   * 1: Timestamp
   * 2: Headrunner Email Address
   * 3: Headrunner Name(s)
   * 4: Headrun
   * 5: Run Level
   * 6: Beginner Attendees
   * 7: Intermediate Attendees
   * 8: Advanced Attendees
   * 9: Validation
   * 10: Distance
   * 11: Comments
   * 12: Copy Sent
   * 13: Submission Platform
   * 14: Not Found (Names)
   */

  // Return value from rawData using `index` and substitute newline with semi-colon.
  // JSON does not accept multi-line values.
  const get = (index => String(values[index - 1]).replace(/\n/g, ';'));

  const timestamp = new Date(`${get(COLUMN_MAP.TIMESTAMP)}`);

  const formattedTimestamp = Utilities.formatDate(
    timestamp,
    TIMEZONE,
    "yyyy-MM-dd HH:mm:ss"
  );
  
  // Initial Mapping
  const exportObj = {
    'timestamp' : formattedTimestamp,
    'headrunners' : get(COLUMN_MAP.HEADRUNNERS),
    'headRun' : get(COLUMN_MAP.HEADRUN),
    'runLevel' : get(COLUMN_MAP.RUN_LEVEL),
    'attendees' : get(COLUMN_MAP.ATTENDEES),
    'confirmation' : get(COLUMN_MAP.CONFIRMATION),
    'distance' : get(COLUMN_MAP.DISTANCE),
    'comments' : get(COLUMN_MAP.COMMENTS),
    'platform' : 'McRUN App',
  }


  return JSON.stringify(exportObj);
}

