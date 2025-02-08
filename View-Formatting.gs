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
  range.sort([{ column: TIMESTAMP_COL, ascending: true }]);
}


function prettifySheet() {
  formatSpecificColumns();
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

function formatSpecificColumns() {
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
    'K2:K'    // Exported
  ]).setVerticalAlignment('middle');

  getThisRange([
    'C2:E',   // Day of the Week + Time + Headrun
    'H2:I',   // Confirmation + Distance
    'K2:K',   // Exported
  ]).setHorizontalAlignment('center');


  // Link pixel size to column index
  const sizeMap = {
    [TIMESTAMP_COL]: 150,
    [HEADRUNNERS_COL]: 200,
    [DAY_WEEK_COL]: 135,
    [SCHEDULED_TIME_COL]: 115,
    [HEADRUN_COL]: 155,
    [RUN_LEVEL_COL]: 120,
    [ATTENDEES_COL]: 340,
    [CONFIRMATION_COL]: 90,
    [DISTANCE_COL]: 110,
    [COMMENTS_COL]: 260,
    [IS_EXPORTED_COL]: 80,
  }
  
  // 8. Resize columns by corresponding pixel size
  for (const [col, width] of Object.entries(sizeMap)) {
    sheet.setColumnWidth(col, width);
  }

}
