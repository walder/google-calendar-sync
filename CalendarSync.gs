// Configuration - Modify these values as needed
const CONFIG = {
  sourceCalendarId: 'a-calendar@google.com',          // Source calendar ID
  targetCalendarId: 'b-calendar@google.com',          // Target calendar ID
  syncDays: 30,                                  // Days to sync forward
  syncPrefixSource: '[av]',                      // Sync event prefix
  syncPrefixTarget: '[al]',                      // Sync event prefix
  metaKeyForID: 'SOURCEID'                       // Key for metadatainformation of original calendar id
};

function syncCalendars() {
  console.log('Starting calendar synchronization...');
  var sourceEvents = getEventsFromCalendar(CONFIG.sourceCalendarId)
  console.log(`Source calendar ${CONFIG.sourceCalendarId} has ${sourceEvents.length} events`);

  var targetEvents = getEventsFromCalendar(CONFIG.targetCalendarId)
  console.log(`Target calendar ${CONFIG.targetCalendarId} has ${targetEvents.length} events`);

  var targetCalendar = CalendarApp.getCalendarById(CONFIG.targetCalendarId);
  var sourceCalendar = CalendarApp.getCalendarById(CONFIG.sourceCalendarId);

  cleanPreviouslySyncedEvents(targetEvents)
  copyEventsFromSourceCalendar(sourceEvents, sourceCalendar, targetCalendar, CONFIG.syncPrefixSource)

  cleanPreviouslySyncedEvents(sourceEvents)
  copyEventsFromSourceCalendar(targetEvents, targetCalendar, sourceCalendar, CONFIG.syncPrefixTarget)
}

function copyEventsFromSourceCalendar(sourceEvents, sourceCalendar, targetCalendar, prefix){
    for (let i = 0; i < sourceEvents.length; i++) {
    var sourceEvent = sourceEvents[i];
    if(sourceEvent.getTag(CONFIG.metaKeyForID) == null && sourceEvent.getMyStatus() != CalendarApp.GuestStatus.NO){
      console.log("Creating Event: "+prefix+" "+sourceEvent.getTitle())

   
      var timezone = sourceCalendar.getTimeZone()
      var targetEvent = targetCalendar.createEvent(prefix+" "+sourceEvent.getTitle(), sourceEvent.getStartTime(), sourceEvent.getEndTime())
      targetEvent.setDescription(sourceEvent.getDescription())
      targetEvent.setLocation(sourceEvent.getLocation())
      targetEvent.setTag(CONFIG.metaKeyForID, sourceEvent.getId())
      if(sourceEvent.isAllDayEvent()){
          targetEvent.setAllDayDates(sourceEvent.getAllDayStartDate(), sourceEvent.getAllDayEndDate())
      }
      //if(sourceEvent.isRecurringEvent())
        //if possible ad later date add code to better update recurring events
        //console.log("Found Recurring Event: "+sourceEvent.getTitle()+" ID:"+sourceEvent.getId())
    }
      }
}

function transformTimeToTimezone(time, timeZone){
   var eventStartTimeInLocalTimeOfCalendar = Utilities.formatDate(time, timeZone, 'MMMM dd, yyyy hh:mm:ss Z');
   console.log(eventStartTimeInLocalTimeOfCalendar)
  return new Date(eventStartTimeInLocalTimeOfCalendar)
}

function cleanPreviouslySyncedEvents(events){
    for (let i = 0; i < events.length; i++) {
      var event = events[i];
      if (event.getTag(CONFIG.metaKeyForID) != null) {
        console.log("deleting previously synced event: "+event.getTitle())
        event.deleteEvent()
      }
    }
}

function getEventsFromCalendar(calendarId){
  var calendar = CalendarApp.getCalendarById(calendarId);
  var timezone = calendar.getTimeZone();

  var syncStartDate = new Date();
  var calendarStartTime = Utilities.formatDate(syncStartDate, timezone, 'MMMM dd, yyyy 00:00:00 Z');
  var syncStartDate = new Date(calendarStartTime)

  var syncEndDate = new Date(syncStartDate.getTime() + (CONFIG.syncDays+1) * 24 * 60 * 60 * 1000);
  var calendarEndTime = Utilities.formatDate(syncEndDate, timezone, 'MMMM dd, yyyy 00:00:00 Z');
  var syncEndDate = new Date(calendarEndTime)

  console.log(`Sync range in Timezone of Calendar: ${calendarStartTime} to ${calendarEndTime}`);
  var sourceEvents = calendar.getEvents(syncStartDate, syncEndDate);
  return sourceEvents;
}
