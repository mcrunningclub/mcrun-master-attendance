/**
 * Sorts sheet according to submission time.
 *
 * @trigger  Edit time.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Feb 8, 2025
 * @update  Feb 8, 2025
 *
 */

function sortAttendanceForm() {
  const sheet = MASTER_ATTENDANCE_SHEET;

  const numRows = sheet.getLastRow() - 1;     // Remove header row from count
  const numCols = sheet.getLastColumn();

  // Sort all the way to the last row, without the header row
  const range = sheet.getRange(2, 1, numRows, numCols);

  // Sorts values by `Timestamp`
  range.sort([{ column: COLUMN_MAP.TIMESTAMP, ascending: true }]);
}


function prettifySheet() {
  formatSpecificColumns_();
}


/**
 * Formats certain columns of Master sheet for a consistent view.
 * 
 * @trigger New Google form or app submission.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Feb 8, 2025
 * @update  Feb 8, 2025
 */

function formatSpecificColumns_() {
  const sheet = MASTER_ATTENDANCE_SHEET;

  // Helper fuction to improve readability
  const getThisRange = (ranges) => 
    Array.isArray(ranges) ? sheet.getRangeList(ranges) : sheet.getRange(ranges);
  
  // 1. Freeze panes
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  // 2. Bold formatting
  getThisRange([
    'A1:N1',  // Header Row
    'A2:A',   // Registration
    'F2:F',   // Run Level
  ]).setFontWeight('bold');

  // 3. Set Italic
  getThisRange('F2:F').setFontStyle('italic')  // Run Level

  // 4. Font size adjustments
  getThisRange('A1:N1').setFontSize(11);  // Header row to size 11
  getThisRange(['A2:F','H2:K']).setFontSize(10);  // Set all rows to 10, except for attendees
  getThisRange(['H1', 'G2:G']).setFontSize(9); // Confirmation Header + Attendees
  
  // 5. Font family adjustment
  getThisRange(['H1','K1']).setFontFamily('Roboto');

  // 6. Format timestamp + headrun time
  getThisRange('A2:A').setNumberFormat('yyyy-MM-dd hh:mm:ss');
  getThisRange('D2:D').setNumberFormat('h:mm AM/PM');
  
  // 7. Horizontal and vertical alignment
  getThisRange('A2:A').setHorizontalAlignment('right');  // Timestamp

  getThisRange([
    'H2:I',   // Confirmation + Distance
    'J2:J',   // Comments
    'K2:K'    // Exported
  ]).setVerticalAlignment('middle');

  getThisRange([
    'C2:E',   // Day of the Week + Time + Headrun
    'H2:I',   // Confirmation + Distance
    'K2:K',   // Exported
  ]).setHorizontalAlignment('center');

  // 8. Set wrapping for comments
  getThisRange('J2:J').setWrap(true);

  // 9. Add checkboxes to confirmation + exported
  const lastRow = getLastSubmission_();
  getThisRange([
    'H2:H' + lastRow,   // Confirmation
    'K2:K' + lastRow,   // Exported
  ]).insertCheckboxes();


  // Link pixel size to column index
  const sizeMap = {
    [COLUMN_MAP.TIMESTAMP]: 150,
    [COLUMN_MAP.HEADRUNNERS]: 200,
    [COLUMN_MAP.DAY_WEEK]: 135,
    [COLUMN_MAP.SCHEDULED_TIME]: 115,
    [COLUMN_MAP.HEADRUN]: 155,
    [COLUMN_MAP.RUN_LEVEL]: 120,
    [COLUMN_MAP.ATTENDEES]: 340,
    [COLUMN_MAP.CONFIRMATION]: 90,
    [COLUMN_MAP.DISTANCE]: 110,
    [COLUMN_MAP.COMMENTS]: 260,
    [COLUMN_MAP.IS_EXPORTED]: 80,
  }
  
  // 10. Resize columns by corresponding pixel size
  for (const [col, width] of Object.entries(sizeMap)) {
    sheet.setColumnWidth(col, width);
  }

}
