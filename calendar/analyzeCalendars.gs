/** David Schraedel - 2024 */


/** TRIGGERS AUTOMATION */

function showSheetTriggers() {
  let triggers = ScriptApp.getProjectTriggers();
  console.log('Current project has ' + triggers.length + ' triggers.');
  for (const property in triggers) {
    console.log(triggers[property].getHandlerFunction());
  }
}

function removeSheetTriggers() {
  var triggers = ScriptApp.getProjectTriggers();

  for (const property in triggers) {
    if (triggers[property].getHandlerFunction() === "analyze") {
      console.log(triggers[property].getHandlerFunction(), ": deleted");
      ScriptApp.deleteTrigger(triggers[property]);
    }
  }
}

function setupSheetTriggers() {
  let triggers = ScriptApp.getProjectTriggers();
  let triggerStrings = "";
  for (const property in triggers) {
    if (triggers[property].getHandlerFunction() === "analyze" && triggers.length > 3) {
      throw new Error(`"${triggers[property].getHandlerFunction()}" trigger already setup.`);
    }
  }
  ScriptApp.newTrigger('analyze').timeBased().onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(23).create();
  ScriptApp.newTrigger('analyze').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(5).create();
  ScriptApp.newTrigger('analyze').timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(5).create();
  ScriptApp.newTrigger('analyze').timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(5).create();
  ScriptApp.newTrigger('analyze').timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(5).create();
  ScriptApp.newTrigger('analyze').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(5).create();
  ScriptApp.newTrigger('analyze').timeBased().onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(5).create();
  ScriptApp.newTrigger('analyze').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(12).create();
  triggerStrings += "analyze";
  console.log(`triggers: ${triggerStrings} were set up`);
}

function analyze() {
  // Determines when script was last run.
  today = new Date();

  analyzeCalendarTime();

  PropertiesService.getScriptProperties().setProperty('analyzeCalendar_LastRun', today);
}


/** LOG SHEET */

const TIME_LOG_COLUMNS = {
  WEEKSTART: "Week Start",
  WEEKEND: "Week End",
  DEFAULT: "Default",
  YOU: "You",
  WORK: "Work",
  RELATIONSHIP: "Relationship",
  TOTAL_PLANNED: "Total % Planned",
  LASTUPDATED: "Last Updated",
}

function getLogSheetVariables(sheet) {
  let colOffset = 1;
  
  let range = sheet.getDataRange();
  let rows = range.getValues();
  let headerRow = rows[0];
  
  let weekStartColumn = headerRow.indexOf(TIME_LOG_COLUMNS.WEEKSTART) + colOffset;
  let weekEndColumn = headerRow.indexOf(TIME_LOG_COLUMNS.WEEKEND) + colOffset;
  let defaultColumn = headerRow.indexOf(TIME_LOG_COLUMNS.DEFAULT) + colOffset;
  let youColumn = headerRow.indexOf(TIME_LOG_COLUMNS.YOU) + colOffset;
  let workColumn = headerRow.indexOf(TIME_LOG_COLUMNS.WORK) + colOffset;
  let relationshipColumn = headerRow.indexOf(TIME_LOG_COLUMNS.RELATIONSHIP) + colOffset;
  let totalPlannedColumn = headerRow.indexOf(TIME_LOG_COLUMNS.TOTAL_PLANNED) + colOffset;
  let lastUpdatedColumn = headerRow.indexOf(TIME_LOG_COLUMNS.LASTUPDATED) + colOffset;

  return {
    weekStartColumn,
    weekEndColumn,
    defaultColumn,
    youColumn,
    workColumn,
    relationshipColumn,
    totalPlannedColumn,
    lastUpdatedColumn,
  };
}

/** MAIN */

const SHEET_NAMES = {
  DEFAULT: "Default",
  YOU: "You",
  WORK: "Work",
  RELATIONSHIP: "Relationship",
  TIME_LOG: "TIME_LOG",
}

const CALENDAR_ID = {
  YOU:"", // ADD CALENDAR ID
  WORK:"", // ADD CALENDAR ID
  RELATIONSHIP:"", // ADD CALENDAR ID
}

const SHEETDOC_ID = {
  ANALYZE:"", // ADD SHEETS DOCUMENT ID
}

function analyzeCalendarTime() {
  /** Get calendar events */
  let app = CalendarApp;

  // calendar IDs
  const youCalID = CALENDAR_ID.YOU;
  const workCalID = CALENDAR_ID.WORK;
  const relationCalID = CALENDAR_ID.RELATIONSHIP;

  // get calendars
  let defaultCalendar = app.getDefaultCalendar();
  let youCalendar = app.getCalendarById(youCalID);
  let workCalendar = app.getCalendarById(workCalID);
  let relationshipCalendar = app.getCalendarById(relationCalID);

  // set week start and end day
  const daysInWeek = 7;
  const dayInMilliseconds = 1000 * 60 * 60 * 24;
  const today = new Date();
  const dayOfWeek = today.getDay();
  let weekDaysRemaining = daysInWeek - 1 - dayOfWeek;
  const dayStartTime = [0,0,0];
  const dayEndTime = [23,59,59];

  let weekStartMilli = today.getTime() - (dayInMilliseconds * dayOfWeek);
  let weekStart = new Date(weekStartMilli);
  weekStart.setHours(dayStartTime[0],dayStartTime[1],dayStartTime[2]);
  let weekEndMilli = today.getTime() + (dayInMilliseconds * weekDaysRemaining);
  let weekEnd = new Date(weekEndMilli);
  weekEnd.setHours(dayEndTime[0],dayEndTime[1],dayEndTime[2]);

  let weekDateLimits = { weekStart,weekEnd };

  // get events for each calendar
  var defaultEvents = getWeekEvents(defaultCalendar, weekDateLimits.weekStart, weekDateLimits.weekEnd);
  var youEvents = getWeekEvents(youCalendar, weekDateLimits.weekStart, weekDateLimits.weekEnd);
  var relationshipEvents = getWeekEvents(relationshipCalendar, weekDateLimits.weekStart, weekDateLimits.weekEnd);
  var workEvents = getWeekEvents(workCalendar, weekDateLimits.weekStart, weekDateLimits.weekEnd);

  /** events to google sheet */
  // get sheet
  let sheetApp = SpreadsheetApp;
  let spreadSheetDoc = sheetApp.openById(SHEETDOC_ID.ANALYZE);

  let defaultSheet = spreadSheetDoc.getSheetByName(SHEET_NAMES.DEFAULT);
  let youSheet = spreadSheetDoc.getSheetByName(SHEET_NAMES.YOU);
  let relationshipSheet = spreadSheetDoc.getSheetByName(SHEET_NAMES.RELATIONSHIP);
  let workSheet = spreadSheetDoc.getSheetByName(SHEET_NAMES.WORK);

  // populate event sheets
  let defaultResult = populateEventSheet(defaultSheet,defaultEvents);
  let youResult = populateEventSheet(youSheet,youEvents);
  let relateResult = populateEventSheet(relationshipSheet,relationshipEvents);
  let workResult = populateEventSheet(workSheet,workEvents);

  /** Log Sheet */
  // append percents to time log sheet
  let timeLogSheet = spreadSheetDoc.getSheetByName(SHEET_NAMES.TIME_LOG);
  let logWeekStart = weekDateLimits.weekStart;
  let logWeekEnd = weekDateLimits.weekEnd;
  let {
    weekStartColumn,
    weekEndColumn,
    defaultColumn,
    youColumn,
    workColumn,
    relationshipColumn,
    totalPlannedColumn,
    lastUpdatedColumn,
  } = getLogSheetVariables(timeLogSheet);

  console.log(timeLogSheet.getName());

  let nextRow = 2;

  let sheetWeekStart = timeLogSheet.getRange(nextRow,weekStartColumn).getValue();
  sheetWeekStart.setMilliseconds(0);
  logWeekStart.setMilliseconds(0);

  let isSameWeek = (sheetWeekStart.getTime() === logWeekStart.getTime());

  if (!isSameWeek) {
    timeLogSheet.insertRows(nextRow);
  }

  timeLogSheet.getRange(nextRow,weekStartColumn).setValue(logWeekStart); // week dates
  timeLogSheet.getRange(nextRow,weekEndColumn).setValue(logWeekEnd); // week dates
  timeLogSheet.getRange(nextRow,defaultColumn).setValue(defaultResult.percentage);
  timeLogSheet.getRange(nextRow,youColumn).setValue(youResult.percentage);
  timeLogSheet.getRange(nextRow,workColumn).setValue(workResult.percentage);
  timeLogSheet.getRange(nextRow,relationshipColumn).setValue(relateResult.percentage);

  timeLogSheet.getRange(nextRow,lastUpdatedColumn).setValue(today);

  // populate summed %s of week
  let defaultCell = timeLogSheet.getRange(nextRow,defaultColumn).getA1Notation()
  let youCell = timeLogSheet.getRange(nextRow,youColumn).getA1Notation()
  let workCell = timeLogSheet.getRange(nextRow,workColumn).getA1Notation()
  let relationshipCell = timeLogSheet.getRange(nextRow,relationshipColumn).getA1Notation()

  timeLogSheet.getRange(nextRow,totalPlannedColumn).setValue(`=SUM(${defaultCell},${youCell},${workCell},${relationshipCell})`);

  // format
  timeLogSheet.getRange(nextRow,defaultColumn).setNumberFormat("0.00%");
  timeLogSheet.getRange(nextRow,youColumn).setNumberFormat("0.00%");
  timeLogSheet.getRange(nextRow,workColumn).setNumberFormat("0.00%");
  timeLogSheet.getRange(nextRow,relationshipColumn).setNumberFormat("0.00%");
  timeLogSheet.getRange(nextRow,totalPlannedColumn).setNumberFormat("0.00%");

  let percentWeekPlanned = timeLogSheet.getRange(nextRow,totalPlannedColumn).getValue().toFixed(2);

  let completionMessage = `completed:\ndefault: ${defaultResult.isSuccess}, you: ${youResult.isSuccess}, relationship: ${relateResult.isSuccess}, work: ${workResult.isSuccess}\n\n\nYou planned ${percentWeekPlanned} of your week`;
  console.log(completionMessage);


}


/** EVENT FUNCTIONS */

function getWeekEvents(calendar,dateStart,dateEnd) {
  return calendar.getEvents(dateStart,dateEnd);
}

const EVENT_COLUMNS = {
  TITLE: "Title",
  START: "Start",
  END: "End",
  DESCRIPTION: "Description",
  OWNER: "Owner",
  DURATIONCALC: "CalculatedDuration",
  TOTAL: "TOTAL",
  PERCENT: "Percent",
};

function getEventSheetVariables(sheet) {
  let colOffset = 1;
  
  let range = sheet.getDataRange();
  let numRows = range.getNumRows();
  let rows = range.getValues();
  let headerRow = rows[0];
  
  let titleColumn = headerRow.indexOf(EVENT_COLUMNS.TITLE) + colOffset;
  let startTimeColumn = headerRow.indexOf(EVENT_COLUMNS.START) + colOffset;
  let endTimeColumn = headerRow.indexOf(EVENT_COLUMNS.END) + colOffset;
  let descriptionColumn = headerRow.indexOf(EVENT_COLUMNS.DESCRIPTION) + colOffset;
  let durationCalcColumn = headerRow.indexOf(EVENT_COLUMNS.DURATIONCALC) + colOffset;
  let totalColumn = headerRow.indexOf(EVENT_COLUMNS.TOTAL) + colOffset;
  let percentColumn = headerRow.indexOf(EVENT_COLUMNS.PERCENT) + colOffset;

  return {
    numRows,
    titleColumn,
    startTimeColumn,
    endTimeColumn,
    descriptionColumn,
    durationCalcColumn,
    totalColumn,
    percentColumn,
  };
}


function populateEventSheet(sheet,events) {
  const weekHoursConstant = 7.00; // numeric duration of a week in google sheets
  console.log(sheet.getName());
  let {
    numRows,
    titleColumn,
    startTimeColumn,
    endTimeColumn,
    descriptionColumn,
    durationCalcColumn,
    totalColumn,
    percentColumn,
  } = getEventSheetVariables(sheet);

  // clear previous value
  let startRow = 2;
  sheet.getRange(startRow,titleColumn,numRows,1).setValue("");
  sheet.getRange(startRow,startTimeColumn,numRows,1).setValue("");
  sheet.getRange(startRow,endTimeColumn,numRows,1).setValue("");
  sheet.getRange(startRow,descriptionColumn,numRows,1).setValue("");
  sheet.getRange(startRow,durationCalcColumn,numRows,1).setValue("");
  sheet.getRange(startRow,totalColumn,numRows,1).setValue("");
  sheet.getRange(startRow,percentColumn,numRows,1).setValue("");

  // iterate events
  let isSuccess = true;
  try {
    events.forEach((e) => {
      let title = e.getTitle();
      let startTime = e.getStartTime();
      let endTime = e.getEndTime();
      let description = e.getDescription();
      let isAllDay = e.isAllDayEvent();

      if (!isAllDay) {
        // populate calendar events
        let lastRow = sheet.getLastRow() + 1;
        sheet.getRange(lastRow,titleColumn).setValue(title);
        sheet.getRange(lastRow,startTimeColumn).setValue(startTime);
        sheet.getRange(lastRow,endTimeColumn).setValue(endTime);
        sheet.getRange(lastRow,descriptionColumn).setValue(description);

        // populate time difference calc
        let endTimeColumnLetter = sheet.getRange(lastRow,endTimeColumn).getA1Notation().split("")[0];
        let startTimeColumnLetter = sheet.getRange(lastRow,startTimeColumn).getA1Notation().split("")[0];
        sheet.getRange(lastRow,durationCalcColumn).setValue(`=${endTimeColumnLetter}${lastRow}-${startTimeColumnLetter}${lastRow}`);
      } else {
        let lastRow = sheet.getLastRow() + 1;
        sheet.getRange(lastRow,titleColumn).setValue(title);
        // sheet.getRange(lastRow,startTimeColumn).setValue(startTime);
        // sheet.getRange(lastRow,endTimeColumn).setValue(endTime);
        sheet.getRange(lastRow,descriptionColumn).setValue(description);
      }
      
    });

    // populate total time sum
    let lastRow = sheet.getLastRow();
    let durationStartCell = sheet.getRange(startRow,durationCalcColumn).getA1Notation();
    let durationEndCell = sheet.getRange(lastRow,durationCalcColumn).getA1Notation();
    
    // set formula in TOTAL col
    sheet.getRange(startRow,totalColumn).setValue(`=SUM(${durationStartCell}:${durationEndCell})`);

    // populate percentage
    let totalSum = sheet.getRange(startRow,totalColumn).getA1Notation();
    
    sheet.getRange(startRow,percentColumn).setValue(`=${totalSum}/${weekHoursConstant}`);
    sheet.getRange(startRow,percentColumn).setNumberFormat("0.00%");

    var percentage = sheet.getRange(startRow,percentColumn).getValues();

  } catch (e) {
    console.error('error while... : ', e.toString());
    isSuccess = false;
  }

  return {isSuccess, percentage};
}

