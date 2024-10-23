
function showTriggers() {
  let triggers = ScriptApp.getProjectTriggers();
  console.log('Current project has ' + ScriptApp.getProjectTriggers().length + ' triggers.');
}

function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
    console.log(i + 1, ": deleted");
  }
}

function setup() {
  let triggers = ScriptApp.getProjectTriggers();
  if (triggers.length > 0) {
    throw new Error('Triggers are already setup.');
  }
  ScriptApp.newTrigger('sync').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).create();
}

function sync() {
  let email = "" // ADD USER EMAIL

  // Determines when script was last run.
  today = new Date();
  let lastRun = PropertiesService.getScriptProperties().getProperty('lastRun');
  lastRun = lastRun ? new Date(lastRun) : null;

  duplicateCalendarWeek();

  PropertiesService.getScriptProperties().setProperty('lastRun', today);
  MailApp.sendEmail(`${email}`, "🗿 Calendars Synced 🗿", `your Google calendars have automatically synced:\n\n${today}`)
}


function duplicateCalendarWeek() {
  // constants
  let app = CalendarApp;
  const daysInWeek = 7;
  const today = new Date();
  const searchKeyword = "*auto*" // choose a keyword

  // calendar IDs
  const youCalID = ""; // ADD CALENDAR ID
  const workCalID = ""; // ADD CALENDAR ID
  const relationCalID = ""; // ADD CALENDAR ID

  // get calendars
  let defaultCalendar = app.getDefaultCalendar();
  let youCalendar = app.getCalendarById(youCalID);
  let workCalendar = app.getCalendarById(workCalID);
  let relationshipCalendar = app.getCalendarById(relationCalID);

  // set span of days
  let sevenDaysPrior = new Date()
  sevenDaysPrior.setDate(today.getDate() - daysInWeek);
  let periodOfDays = [sevenDaysPrior,today];

  // get events
  var defaultEvents = getEvents(defaultCalendar, periodOfDays[0], periodOfDays[1],{search: searchKeyword});
  var youEvents = getEvents(youCalendar, periodOfDays[0], periodOfDays[1],{search: searchKeyword});
  var relationshipEvents = getEvents(relationshipCalendar, periodOfDays[0], periodOfDays[1],{search: searchKeyword});
  var workEvents = getEvents(workCalendar, periodOfDays[0], periodOfDays[1],{search: searchKeyword});
  
  // duplicate events
  duplicateEvents(defaultCalendar,defaultEvents,daysInWeek);
  duplicateEvents(youCalendar,youEvents,daysInWeek);
  duplicateEvents(relationshipCalendar,relationshipEvents,daysInWeek);
  duplicateEvents(workCalendar,workEvents,daysInWeek);

}

function getEvents(calendar,dateStart,dateEnd, object) {
  return calendar.getEvents(dateStart,dateEnd, object);
}

function duplicateEvents(calendar, calEvents, daysInWeek) {
  // iterate events
  calEvents.forEach((event) => {
    // new times
    let newStartTime = event.getStartTime();
    newStartTime.setDate(event.getStartTime().getDate() + daysInWeek);
    let newEndTime = event.getEndTime();
    newEndTime.setDate(event.getEndTime().getDate() + daysInWeek);

    // copy other parameters
    let eventTitle = event.getTitle();
    let eventLocation = event.getLocation();
    let eventDescription = event.getDescription();

    // console.log(typeof newEndTime);
    
    // create New Event
    try {
      let newEvent = calendar.createEvent(eventTitle,newStartTime,newEndTime,
        {
          location: eventLocation,
          description: eventDescription,
        }
      );
      // make any other changes to event
      try {
        if (event.getColor() !== calendar.getColor() && event.getColor() !== "") {
          newEvent.setColor(event.getColor());
        }
      } catch (e) {
        console.error('error while changing color', e.toString());
      }
    } catch (e) {
      console.error('error while creating event', e.toString());
    }
  });

  return;
}







