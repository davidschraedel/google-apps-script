/** David Schraedel - March 2024 */
// clock in and out, via assigning in and out functions to buttons on a google sheet

const SHEET_NAMES = {
  TIMESHEET: "TimeSheet",
}

const SHEET_COLUMNS = {
  IN: "In",
  OUT: "Out",
  HOURS: "Hours",
  TOTAL: "Total"
}

function clockIn() {
  // get date and time
  let timeIn = new Date()

  // get sheet doc
  const sheetApp = SpreadsheetApp;
  const spreadSheet = sheetApp.getActiveSpreadsheet();

  // get sheet
  let timeSheet = spreadSheet.getSheetByName(SHEET_NAMES.TIMESHEET);

  // get columns
  let {
    clockInColumn,
    clockOutColumn,
    hoursColumn,
    totalColumn,
  } = getSheetColumns(timeSheet);

  // get next row
  let nextRow = 2; // inserting a row makes 2 the next row indefinitely

  // check if clocked out. (empty "in" cell, next to empty "out")
  let isClockInColFilled = timeSheet.getRange(nextRow,clockInColumn).getValue() !== "";
  let isClockOutColFilled = timeSheet.getRange(nextRow,clockOutColumn).getValue() !== "";
  let isClockedOut = false;
  if (isClockInColFilled) {
    if (isClockOutColFilled) isClockedOut = true;
  } else {
    if (!isClockOutColFilled) isClockedOut = true;
  }
  if (!isClockedOut) {
    throw("⚠️ You are already clocked in");
  } else {
    // populate "In" column cell with date and time
    timeSheet.insertRows(nextRow);
    timeSheet.getRange(nextRow,clockInColumn).setValue(timeIn);
  }

}


function clockOut() {
  // get date and time
  let timeOut = new Date()

  // get sheet doc
  const sheetApp = SpreadsheetApp;
  const spreadSheet = sheetApp.getActiveSpreadsheet();

  // get sheet
  let timeSheet = spreadSheet.getSheetByName(SHEET_NAMES.TIMESHEET);

  // get columns
  let {
    clockInColumn,
    clockOutColumn,
    hoursColumn,
    totalColumn,
  } = getSheetColumns(timeSheet);

  // get next row
  let nextRow = 2; // inserting a row makes 2 the next row indefinitely

  // check if clocked in
  const isClockInColFilled = timeSheet.getRange(nextRow,clockInColumn).getValue() !== "";
  const isClockOutColFilled = timeSheet.getRange(nextRow,clockOutColumn).getValue() !== "";
  let isClockedIn = isClockInColFilled && !isClockOutColFilled;
  if (!isClockedIn) {
    throw("⚠️ You are already clocked out");
  }

  // populate "Out" column cell with date and time
  timeSheet.getRange(nextRow,clockOutColumn).setValue(timeOut);

  // update "Hours" col
  let clockInCell = timeSheet.getRange(nextRow,clockInColumn).getA1Notation();
  let clockOutCell = timeSheet.getRange(nextRow,clockOutColumn).getA1Notation();

  timeSheet.getRange(nextRow,hoursColumn).setValue(`=SUM(${clockOutCell}-${clockInCell})`);
  timeSheet.getRange(nextRow,hoursColumn).setNumberFormat('[h]:mm:ss');

  // update "Total" col
  let lastRow = timeSheet.getLastRow() - 1;
  let startRow = nextRow;
  let numColumns = 1;

  let dataRange = timeSheet.getDataRange();
  let allRows = dataRange.getNumRows();

  let hoursRange = timeSheet.getRange(startRow,hoursColumn,lastRow,numColumns).getA1Notation();
  timeSheet.getRange(startRow,totalColumn,allRows,numColumns).setValue("");
  timeSheet.getRange(nextRow,totalColumn).setValue(`=SUM(${hoursRange})`);
  timeSheet.getRange(nextRow,totalColumn).setNumberFormat('[h]:mm:ss');

}


function getSheetColumns(sheet) {
  let colOffset = 1;
  
  let range = sheet.getDataRange();
  let rows = range.getValues();
  let headerRow = rows[0];
  
  let clockInColumn = headerRow.indexOf(SHEET_COLUMNS.IN) + colOffset;
  let clockOutColumn = headerRow.indexOf(SHEET_COLUMNS.OUT) + colOffset;
  let hoursColumn = headerRow.indexOf(SHEET_COLUMNS.HOURS) + colOffset;
  let totalColumn = headerRow.indexOf(SHEET_COLUMNS.TOTAL) + colOffset;

  return {
    clockInColumn,
    clockOutColumn,
    hoursColumn,
    totalColumn,
  }
}
