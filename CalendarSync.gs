// Configuration - Modify these values as needed
const CONFIG = {
  sourceCalendarId: 'a-calendar@google.com',          // Source calendar ID
  targetCalendarId: 'b-calendar@google.com',          // Target calendar ID
  syncDays: 60,                                  // Days to sync forward
  syncPrefixSource: '[av]',                      // Sync event prefix
  syncPrefixTarget: '[al]',                      // Sync event prefix
  metaKeyForID: 'SOURCEID',                       // Key for metadatainformation of original calendar id
  metaLastUpdated: 'SOURCE_LAST_UPDATED'
};

function syncCalendars() {
  var targetCalendar = CalendarApp.getCalendarById(CONFIG.targetCalendarId);
  var sourceCalendar = CalendarApp.getCalendarById(CONFIG.sourceCalendarId);

  console.log('Starting calendar synchronization...');
  var sourceEvents = getEventsFromCalendar(sourceCalendar)
  console.log(`Source calendar ${CONFIG.sourceCalendarId} has ${sourceEvents.length} events`);

  var targetEvents = getEventsFromCalendar(targetCalendar)
  console.log(`Target calendar ${CONFIG.targetCalendarId} has ${targetEvents.length} events`);

  copyEventsFromSourceCalendar(sourceEvents, targetEvents, sourceCalendar, targetCalendar, CONFIG.syncPrefixSource)
  copyEventsFromSourceCalendar(targetEvents, sourceEvents, targetCalendar, sourceCalendar, CONFIG.syncPrefixTarget)
}

function copyEventsFromSourceCalendar(sourceEvents, targetEvents, sourceCalendar, targetCalendar, prefix){
    for (let i = 0; i < sourceEvents.length; i++) {
      var sourceEvent = sourceEvents[i];
      createOrUpdateTargetEvent(sourceEvent, targetEvents, sourceCalendar, targetCalendar, prefix)
    }
    for (let i = 0; i < targetEvents.length; i++) {
      var targetEvent = targetEvents[i];
      var isOrphan = true;
      var targetEventMetaKeyId = targetEvent.getTag(CONFIG.metaKeyForID)
      for (let j = 0; j < sourceEvents.length; j++) {
        var sourceEvent = sourceEvents[j];
        if (targetEventMetaKeyId == sourceEvent.getId() && isRegularEventOrExactInstanceOfReocurringEvent(sourceEvent, targetEvent)){
          isOrphan = false;
        }
      }
      if (isOrphan && targetEventMetaKeyId != null){
        console.log("Target Event is Orphan will be deleted: "+targetEvent.getTitle())
        targetEvent.deleteEvent()
      }
    }
}

function isRegularEventOrExactInstanceOfReocurringEvent(sourceEvent, targetEvent){
  if(sourceEvent.isRecurringEvent()){
    if(sourceEvent.getStartTime().getTime() == targetEvent.getStartTime().getTime() 
      && sourceEvent.getEndTime().getTime() == targetEvent.getEndTime().getTime()){
        return true;
    }
    else{
      return false;
    }
  }
  else{
    return true;
  }
}

function createOrUpdateTargetEvent(sourceEvent, targetEvents, sourceCalendar, targetCalendar, prefix){
  var sourceEventExists = false;
   for (let i = 0; i < targetEvents.length; i++) {
    var targetEvent = targetEvents[i]
    var isRegularOrRecurringEvent = isRegularEventOrExactInstanceOfReocurringEvent(sourceEvent, targetEvent)
    //console.log("Checking Event: " +targetEvent.getTitle()+" | Is Regular or Recurring: "+isRegularOrRecurringEvent)
    if (targetEvent.getTag(CONFIG.metaKeyForID) == sourceEvent.getId() && isRegularOrRecurringEvent){
      sourceEventExists = true
      var targetLastUpdateDate = targetEvent.getTag(CONFIG.metaLastUpdated)
      //console.log("Source Event last updated: "+sourceEvent.getLastUpdated().getTime()+" Target Event last Updated: "+targetLastUpdateDate)
      if(sourceEvent.getLastUpdated().getTime() == targetLastUpdateDate){
        if(sourceEvent.getMyStatus() == CalendarApp.GuestStatus.NO){
          console.log("Deleting event because declined in source calendar: "+targetEvent.getTitle())
          targetEvent.deleteEvent()
        }
      }
       else{
          console.log("this event needs updating: "+sourceEvent.getTitle())
          fillEventFields(sourceEvent, targetEvent, targetCalendar, sourceCalendar, prefix)
        }
      /*
      else{
        console.log("This Event does not need update: "+sourceEvent.getTitle())
      }*/
    }
   }
   if (!sourceEventExists && sourceEvent.getTag(CONFIG.metaKeyForID) == null){
    if(sourceEvent.getMyStatus() != CalendarApp.GuestStatus.NO){
      console.log("creating new event: "+prefix+" "+sourceEvent.getTitle())
      var targetEvent = targetCalendar.createEventFromDescription(sourceEvent.getTitle())
      fillEventFields(sourceEvent, targetEvent, targetCalendar, sourceCalendar, prefix)
    }
   }
}

function fillEventFields(sourceEvent, targetEvent, targetCalendar, sourceCalendar, prefix){
  var timezone = sourceCalendar.getTimeZone()
  targetEvent.setTitle(prefix+" "+sourceEvent.getTitle())
  targetEvent.setDescription(sourceEvent.getDescription())
  targetEvent.setTime(sourceEvent.getStartTime(), sourceEvent.getEndTime())
  targetEvent.setLocation(sourceEvent.getLocation())
  targetEvent.setTag(CONFIG.metaKeyForID, sourceEvent.getId())
  targetEvent.setTag(CONFIG.metaLastUpdated, sourceEvent.getLastUpdated().getTime())
  if(sourceEvent.isAllDayEvent()){
    targetEvent.setAllDayDates(sourceEvent.getAllDayStartDate(), sourceEvent.getAllDayEndDate())
  }
}

function transformTimeToTimezone(time, timeZone){
   var eventStartTimeInLocalTimeOfCalendar = Utilities.formatDate(time, timeZone, 'MMMM dd, yyyy hh:mm:ss Z');
   console.log(eventStartTimeInLocalTimeOfCalendar)
  return new Date(eventStartTimeInLocalTimeOfCalendar)
}

function deleteSyncedEventsOfCalendars(){
  var targetCalendar = CalendarApp.getCalendarById(CONFIG.targetCalendarId);
  var sourceCalendar = CalendarApp.getCalendarById(CONFIG.sourceCalendarId);

  console.log('Starting calendar synchronization...');
  var sourceEvents = getEventsFromCalendar(sourceCalendar)
  console.log(`Source calendar ${CONFIG.sourceCalendarId} has ${sourceEvents.length} events`);

  var targetEvents = getEventsFromCalendar(targetCalendar)
  console.log(`Target calendar ${CONFIG.targetCalendarId} has ${targetEvents.length} events`);

  cleanPreviouslySyncedEvents(targetEvents)
  cleanPreviouslySyncedEvents(sourceEvents)
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

function getEventsFromCalendar(calendar){
  var timezone = calendar.getTimeZone();

  var syncStartDate = new Date();
  var calendarStartTime = Utilities.formatDate(syncStartDate, timezone, 'MMMM dd, yyyy 00:00:00 Z');
  var syncStartDate = new Date(calendarStartTime)

  var syncEndDate = new Date(syncStartDate.getTime() + (CONFIG.syncDays) * 24 * 60 * 60 * 1000);
  var calendarEndTime = Utilities.formatDate(syncEndDate, timezone, 'MMMM dd, yyyy 00:00:00 Z');
  var syncEndDate = new Date(calendarEndTime)

  console.log(`Sync range in Timezone of Calendar: ${calendarStartTime} to ${calendarEndTime}`);
  var sourceEvents = calendar.getEvents(syncStartDate, syncEndDate);
  return sourceEvents;
}
