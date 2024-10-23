// David Schraedel, 2024
// automations for job applications google sheet

function showJobApplicationTriggers() {
  let triggers = ScriptApp.getProjectTriggers();
  console.log('Current project has ' + triggers.length + ' triggers.');
  for (const property in triggers) {
    console.log(`${triggers[property].getUniqueId()}, triggered by ${triggers[property].getTriggerSource().toString()}, calls ${triggers[property].getHandlerFunction()}()`);
    
    // console.log(triggers[property].getEventType().toString())
  }
}
function removeJobApplicationTriggers() {
  var triggers = ScriptApp.getProjectTriggers();

  for (const property in triggers) {
    if (triggers[property].getHandlerFunction() === "checkFollowupDates") {
      console.log(triggers[property].getHandlerFunction(), ": deleted");
      ScriptApp.deleteTrigger(triggers[property]);
    }
  }
}
function setupJobApplicationTriggers() {
  let triggers = ScriptApp.getProjectTriggers();
  let triggerStrings = "";
  for (const property in triggers) {
    if (triggers[property].getHandlerFunction() === "checkFollowupDates" && triggers.length > 3) {
      throw new Error(`"${triggers[property].getHandlerFunction()}" trigger already setup.`);
    }
  }
  ScriptApp.newTrigger('checkFollowupDates').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(1).create();
  ScriptApp.newTrigger('checkFollowupDates').timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(1).create();
  ScriptApp.newTrigger('checkFollowupDates').timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(1).create();
  ScriptApp.newTrigger('checkFollowupDates').timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(1).create();
  ScriptApp.newTrigger('checkFollowupDates').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(1).create();
  ScriptApp.newTrigger('checkFollowupDates').timeBased().onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(1).create();
  triggerStrings += "checkFollowupDates";
  console.log(`triggers: ${triggerStrings} were set up`);
}


const SHEET_NAMES = {
  APPLICATIONS: "APPLICATIONS",
}

const SHEET_COLUMNS = {
  FOLLOWUP_DATE: "FollowUp Date",
  DATE_APPLIED: "Date Applied",
  POSITION: "Position",
  COMPANY: "Company",
  RECRUITER: "Recruiter",
  REFERRER: "Referrer",
  LAST_INTERVIEW: "Last Interview",
  EXPECTED_REPLY_DATE: "Expected Reply Date",
  FOLLOWUP_DATE: "Followup Date"
}

const EMAIL_ADDRESS = "" // ADD EMAIL ADDRESS

function checkFollowupDates() {
  const Today = new Date();
  // const Now = Date.now()
  const MillisecondsInDay = 86400000;
  const ThreeDays = Math.round((MillisecondsInDay * 3) / MillisecondsInDay);
  const TwoDays = Math.round((MillisecondsInDay * 2) / MillisecondsInDay);

  // get sheet doc
  const sheetApp = SpreadsheetApp;
  const spreadSheet = sheetApp.getActiveSpreadsheet();

  // get sheet
  let applicationsSheet = spreadSheet.getSheetByName(SHEET_NAMES.APPLICATIONS);


  let {
      followupDate_Col,
      dateApplied_Col,
      position_Col ,
      company_Col ,
      recruiter_Col ,
      referrer_Col ,
      lastInterview_Col ,
      expectedReplyDate_Col
    } = getSheetColumns(applicationsSheet)


  // get cells
  let firstRow = 2;
  let numRows = applicationsSheet.getMaxRows();

  let followupRange = applicationsSheet.getRange(firstRow,followupDate_Col,numRows);
  let followupValues = followupRange.getValues();

  // populate list of relevant rows
  let jobsToFollowupOn = [];

  for (const [index, value] of followupValues.entries()) {
    if (value[0] !== "") {
      if (new Date(value[0]) > new Date(Today)) {
        let milliDifference = value[0].getTime() - Today.getTime();
        let daysDifference = Math.round(milliDifference / MillisecondsInDay)
        if (daysDifference <= ThreeDays && daysDifference >= TwoDays) {
          // only push jobs with text that is not striked through
          if (!applicationsSheet.getRange(index + 2,followupDate_Col).getTextStyle().isStrikethrough()) {
            jobsToFollowupOn.push(
            {
              followupDate: value[0].toLocaleDateString('en-us', { weekday:"long", year:"numeric", month:"short", day:"numeric"}),
              date_applied: applicationsSheet.getRange(index + 2, dateApplied_Col).getValue() === "" ? "Unknown Date" : applicationsSheet.getRange(index + 2, dateApplied_Col).getValue().toLocaleDateString('en-us', { weekday:"long", year:"numeric", month:"short", day:"numeric"}),
              position: applicationsSheet.getRange(index + 2, position_Col).getValue(),
              company: applicationsSheet.getRange(index + 2, company_Col).getValue(),
              recruiter: applicationsSheet.getRange(index + 2, recruiter_Col).getValue(),
              referrer: applicationsSheet.getRange(index + 2, referrer_Col).getValue(),
              last_interview: applicationsSheet.getRange(index + 2, lastInterview_Col).getValue() === "" ? "Not Yet Specified" : applicationsSheet.getRange(index + 2, lastInterview_Col).getValue().toLocaleDateString('en-us', { weekday:"long", year:"numeric", month:"short", day:"numeric"}),
              expected_reply_date: applicationsSheet.getRange(index + 2, expectedReplyDate_Col).getValue() === "" ? "Not Yet Specified" : applicationsSheet.getRange(index + 2, expectedReplyDate_Col).getValue().toLocaleDateString('en-us', { weekday:"long", year:"numeric", month:"short", day:"numeric"}),
            }
          )
          }
          
        }
      }
    }
  }


  var html = HtmlService.createTemplateFromFile("Email");

  // for each job to followup on...
  for (const job of jobsToFollowupOn) {
    // populate html values
    html.position = job.position
    html.company = job.company
    html.last_interview = job.last_interview
    html.expected_reply_date = job.expected_reply_date
    html.followup_date = job.followupDate
    html.recruiter = job.recruiter
    html.date_applied = job.date_applied
    html.referrer = job.referrer

    // create plain text format
    let emailText = `Follow up on ${job.position} at ${job.company} from your interview on ${job.last_interview}. You expected a reply by ${job.expected_reply_date}, and planned to follow up on ${job.followupDate}. \n\nRecruiters include the following: ${job.recruiter}. \nYou applied on ${job.date_applied},and if referred, were referred by ${job.referrer}. \nBest!`;

    // send email
    let htmlContent = html.evaluate().getContent();
    MailApp.sendEmail(EMAIL_ADDRESS, `Follow up 🛎 on ${job.position} position 💻`, emailText, { htmlBody: htmlContent });
  }
}



function getSheetColumns(sheet) {
  let colOffset = 1;
  
  let range = sheet.getDataRange();
  let rows = range.getValues();
  let headerRow = rows[0];
  
  let followupDate_Col = headerRow.indexOf(SHEET_COLUMNS.FOLLOWUP_DATE) + colOffset;
  let dateApplied_Col = headerRow.indexOf(SHEET_COLUMNS.DATE_APPLIED) + colOffset;
  let position_Col = headerRow.indexOf(SHEET_COLUMNS.POSITION) + colOffset;
  let company_Col = headerRow.indexOf(SHEET_COLUMNS.COMPANY) + colOffset;
  let recruiter_Col = headerRow.indexOf(SHEET_COLUMNS.RECRUITER) + colOffset;
  let referrer_Col = headerRow.indexOf(SHEET_COLUMNS.REFERRER) + colOffset;
  let lastInterview_Col = headerRow.indexOf(SHEET_COLUMNS.LAST_INTERVIEW) + colOffset;
  let expectedReplyDate_Col = headerRow.indexOf(SHEET_COLUMNS.EXPECTED_REPLY_DATE) + colOffset;

  return {
    followupDate_Col,
    dateApplied_Col,
    position_Col ,
    company_Col ,
    recruiter_Col ,
    referrer_Col ,
    lastInterview_Col ,
    expectedReplyDate_Col,
  }
}





