/**
 * Find row index of last submission in reverse using while-loop.
 * 
 * Used to prevent native `sheet.getLastRow()` from returning empty row.
 * 
 * @return {integer}  Returns 1-index of last row in GSheet.
 *  
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Feb 8, 2025
 * @update  Feb 8, 2025
 */

function getLastSubmission_() {
  const sheet = MASTER_ATTENDANCE_SHEET;
  const startRow = 1;
  const numRow = sheet.getLastRow();
  
  // Fetch all values in the TIMESTAMP_COL
  const values = sheet.getRange(startRow, COLUMN_MAP.TIMESTAMP, numRow).getValues();
  let lastRow = values.length;

  // Loop through the values in reverse order
  while (values[lastRow - 1][0] === "") {
    lastRow--;
  }

  return lastRow;
}


/**
 * Formats headrunner names from `row` into uniform view, separated by newline.
 * 
 * @param {Array<Integer>} targetCols  The column(s) with names to format.
 *
 * @param {integer} [row=ATTENDANCE_SHEET.getLastRow()]  The row in the `ATTENDANCE_SHEET` sheet (1-indexed).
 *                                                       Defaults to the last row in the sheet.
 *
 * @param {integer} [numRow=1] numRow  Number of rows to format from `startRow`.
 *
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Feb 8, 2025
 * @update  Feb 8, 2025
 *
 * ```javascript
 * // Sample Script ➜ Format names in last row for ATTENDEES.
 * formatHeadRunnerInRow([ATTENDEES_COL]);
 * 
 * // Sample Script ➜ Format names in row `7` in TIMESTAMP and ATTENDEES.
 * const targetCols = [TIMESTAMP_COL, ATTENDEES_COL]
 * const rowToFormat = 7;
 * formatHeadRunnerInRow(targetCols, rowToFormat);
 *
 * // Sample Script ➜ Format names from row `3` to `9` in TIMESTAMP.
 * const targetCols = [TIMESTAMP_COL]
 * const startRow = 3;
 * const numRow = 9 - startRow;
 * formatHeadRunnerInRow(targetCols, startRow, numRow);
 * ```
 */

function formatNamesInRow_(targetCols, startRow=getLastSubmission(), numRow=1) {
  const sheet = MASTER_ATTENDANCE_SHEET;

  targetCols.forEach(targetCol => {
    // Get all the values in target col(s)
    const rangeNames = sheet.getRange(startRow, targetCol, numRow);
    const rawValues = rangeNames.getValues();

    // Callback function to process the raw value into the formatted format
    function processRow(row) {
      const names = row[0]       // Get first column from 2D array
        .split(/,\s*|\s*,\s*/)   // Split by comma and/or spaces
        .join('\n');             // Join the names with a newline

      return [names];   // Return as a 2D array for .setValues()
    };

    // Map over each row to process and format by applying `processRow()`
    const formattedNames = rawValues.map(processRow);

    // Update the sheet with formatted names
    rangeNames.setValues(formattedNames);

  });
}


