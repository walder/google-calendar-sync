// Configuration - Modify these values as needed
const CONFIG = {
  sourceCalendarId: 'a-calendar@google.com',    // Source calendar ID
  targetCalendarId: 'b-calendar@google.com',    // Target calendar ID
  syncDays: 60,                                 // Days to sync forward
  syncPrefixSource: '[av]',                     // Sync event prefix
  syncPrefixTarget: '[al]',                     // Sync event prefix
  metaKeyForID: 'SOURCE_ID',                    // Key for metadata information of original calendar id
  metaLastUpdated: 'SOURCE_LAST_UPDATED',
  metaIsRecurring: 'SOURCE_IS_RECURRING'
};

function syncCalendars(timeDelta) {
  const startTime = new Date().getTime();
  if (!timeDelta) {
    timeDelta = 0;
  }
  const targetCalendar = CalendarApp.getCalendarById(CONFIG.targetCalendarId);
  const sourceCalendar = CalendarApp.getCalendarById(CONFIG.sourceCalendarId);

  console.log('Starting calendar synchronization...');
  const sourceEvents = getEventsFromCalendar(sourceCalendar);
  console.log(`Source calendar ${CONFIG.sourceCalendarId} has ${sourceEvents.length} events`);

  const targetEvents = getEventsFromCalendar(targetCalendar);
  console.log(`Target calendar ${CONFIG.targetCalendarId} has ${targetEvents.length} events`);

  copyEventsFromSourceCalendar(sourceEvents, targetEvents, sourceCalendar, targetCalendar, CONFIG.syncPrefixSource, timeDelta);
  copyEventsFromSourceCalendar(targetEvents, sourceEvents, targetCalendar, sourceCalendar, CONFIG.syncPrefixTarget, timeDelta);
  const endTime = new Date().getTime();
  console.log("Run Time for Sync: " + (endTime - startTime) / 1000);
}

function syncCalendarsWithTimeDelta() {
  const currentTime = new Date().getTime();
  const lookBack = 600;
  const deltaTime = (currentTime - lookBack * 1000);
  console.log("Will Sync Eents that have changed in the last: " + lookBack / 60 + " Minutes");
  syncCalendars(deltaTime);
}

function createMapKeyForCalendarEvent(event, id, isRecurring) {
  let key = id;
  if (String(isRecurring).toLowerCase() == "true") {
    key += "-" + event.getStartTime() + "-" + event.getEndTime();
  }
  return key;
}

function copyEventsFromSourceCalendar(sourceEvents, targetEvents, sourceCalendar, targetCalendar, prefix, timeDelta) {
  const newTargetEvents = [];
  const orphanMap = new Map();
  const targetMap = new Map();
  for (let i = 0; i < targetEvents.length; i++) {
    const targetEvent = targetEvents[i];
    //we only add events that are copies from the source calendar, will speed up checks
    if (targetEvent.getTag(CONFIG.metaKeyForID) != null) {
      newTargetEvents.push(targetEvent);
      const key = createMapKeyForCalendarEvent(targetEvent, targetEvent.getTag(CONFIG.metaKeyForID), targetEvent.getTag(CONFIG.metaIsRecurring));
      // console.log("Adding to Orphans Map: "+key)
      orphanMap.set(key, targetEvent);
      targetMap.set(key, targetEvent);
    }
  }
  targetEvents = newTargetEvents;

  let startTime = new Date().getTime();
  for (let i = 0; i < sourceEvents.length; i++) {
    const sourceEvent = sourceEvents[i];
    //we only check original events ot the calendar, this speeds up runs after the initial sync. We only sync events that are less than timeDelta old.
    //if run on calendar update trigger this should reduce the amount of events we have to check significanly.
    if (sourceEvent.getTag(CONFIG.metaKeyForID) == null) {
      const key = createMapKeyForCalendarEvent(sourceEvent, sourceEvent.getId(), sourceEvent.isRecurringEvent());
      // console.log("Deleting from Orphans Map: "+key)
      orphanMap.delete(key);
      /*
      optimized sync, if last change newer than time delta, or we can't find the event in the already synced events
      happens when we move one day ahead in the check, and there are events that have been created a long time ago that now need syncing.
      */
      if (sourceEvent.getLastUpdated().getTime() >= timeDelta || !targetMap.get(key)) {
        createOrUpdateTargetEvent(sourceEvent, targetEvents, sourceCalendar, targetCalendar, prefix, targetMap);
      }
    }
  }
  let endTime = new Date().getTime();
  console.log("Sync Source events from " + sourceCalendar.getId() + " took: " + (endTime - startTime) / 1000 + " seconds");

  startTime = new Date().getTime();
  orphanMap.forEach((value, key) => {
    console.log("Deleting orphan: " + value.getTitle());
    try {
      value.deleteEvent();
    }
    catch {
      console.log(`Error while deleting orphan event "${value.getTitle()}" (${value.getId()}): ${e}`);
    }
  });
  endTime = new Date().getTime();
  console.log("Orphan clean up for " + targetCalendar.getId() + " took: " + (endTime - startTime) / 1000 + " seconds");
}

function createOrUpdateTargetEvent(sourceEvent, targetEvents, sourceCalendar, targetCalendar, prefix, targetMap) {
  let targetEvent = targetMap.get(createMapKeyForCalendarEvent(sourceEvent, sourceEvent.getId(), sourceEvent.isRecurringEvent()));
  if (!targetEvent) {
    if (sourceEvent.getMyStatus() != CalendarApp.GuestStatus.NO) {
      console.log("creating new event: " + prefix + " " + sourceEvent.getTitle());
      try {
        targetEvent = targetCalendar.createEventFromDescription(sourceEvent.getTitle());
        fillEventFields(sourceEvent, targetEvent, targetCalendar, sourceCalendar, prefix);
      }
      catch (e) {
        console.log(`Error while creating "${sourceEvent.getTitle()}" (${sourceEvent.getId()}): ${e}`);
      }
    }
  } else {
    const targetLastUpdateDate = Number(targetEvent.getTag(CONFIG.metaLastUpdated));
    if (sourceEvent.getLastUpdated().getTime() != targetLastUpdateDate) {
      console.log("this event needs updating: " + sourceEvent.getTitle());
      fillEventFields(sourceEvent, targetEvent, targetCalendar, sourceCalendar, prefix);
    }
    if (sourceEvent.getMyStatus() == CalendarApp.GuestStatus.NO) {
      console.log("Deleting event because declined in source calendar: " + targetEvent.getTitle());
      try {
        targetEvent.deleteEvent();
      }
      catch (e) {
        console.log(`Error while deleting "${sourceEvent.getTitle()}" (${sourceEvent.getId()}): ${e}`);
      }
    }
  }
}

function fillEventFields(sourceEvent, targetEvent, targetCalendar, sourceCalendar, prefix) {
  try {
    //checking every field for change improves performance, since every update is a network request
    if (targetEvent.getTitle() != (prefix + " " + sourceEvent.getTitle())) {
      // console.log("update Title")
      targetEvent.setTitle(prefix + " " + sourceEvent.getTitle());
    }
    if (targetEvent.getDescription() != sourceEvent.getDescription()) {
      // console.log("update desc")
      targetEvent.setDescription(sourceEvent.getDescription());
    }
    if (targetEvent.getStartTime().getTime() != sourceEvent.getStartTime().getTime() || targetEvent.getEndTime().getTime() != sourceEvent.getEndTime().getTime()) {
      // console.log("update start end time")
      targetEvent.setTime(sourceEvent.getStartTime(), sourceEvent.getEndTime());
    }
    if (targetEvent.getLocation() != sourceEvent.getLocation()) {
      // console.log("update location")
      targetEvent.setLocation(sourceEvent.getLocation());
    }
    if (targetEvent.getTag(CONFIG.metaKeyForID) != sourceEvent.getId()) {
      // console.log("update ID")
      targetEvent.setTag(CONFIG.metaKeyForID, sourceEvent.getId());
    }
    if (Number(targetEvent.getTag(CONFIG.metaLastUpdated)) != sourceEvent.getLastUpdated().getTime()) {
      // console.log("update lastupdated")
      targetEvent.setTag(CONFIG.metaLastUpdated, sourceEvent.getLastUpdated().getTime());
    }
    if (targetEvent.getTag(CONFIG.metaIsRecurring) != String(sourceEvent.isRecurringEvent())) {
      // console.log("update recurring")
      targetEvent.setTag(CONFIG.metaIsRecurring, String(sourceEvent.isRecurringEvent()));
    }
    if (sourceEvent.isAllDayEvent() != targetEvent.isAllDayEvent()) {
      // console.log("all day")
      targetEvent.setAllDayDates(sourceEvent.getAllDayStartDate(), sourceEvent.getAllDayEndDate());
    }
  }
  catch (e) {
    console.log(`Error while updating "${sourceEvent.getTitle()}" (${sourceEvent.getId()}): ${e}`);
  }
}

function deleteSyncedEventsOfCalendars() {
  const targetCalendar = CalendarApp.getCalendarById(CONFIG.targetCalendarId);
  const sourceCalendar = CalendarApp.getCalendarById(CONFIG.sourceCalendarId);

  console.log('Starting calendar synchronization...');
  const sourceEvents = getEventsFromCalendar(sourceCalendar);
  console.log(`Source calendar ${CONFIG.sourceCalendarId} has ${sourceEvents.length} events`);

  const targetEvents = getEventsFromCalendar(targetCalendar);
  console.log(`Target calendar ${CONFIG.targetCalendarId} has ${targetEvents.length} events`);

  cleanPreviouslySyncedEvents(targetEvents);
  cleanPreviouslySyncedEvents(sourceEvents);
}

function cleanPreviouslySyncedEvents(events) {
  for (let i = 0; i < events.length; i++) {
    const event = events[i];
    if (event.getTag(CONFIG.metaKeyForID) != null) {
      console.log("deleting previously synced event: " + event.getTitle());
      event.deleteEvent();
    }
  }
}

function getEventsFromCalendar(calendar) {
  const timezone = calendar.getTimeZone();

  let syncStartDate = new Date();
  const calendarStartTime = Utilities.formatDate(syncStartDate, timezone, 'MMMM dd, yyyy 00:00:00 Z');
  syncStartDate = new Date(calendarStartTime);

  let syncEndDate = new Date(syncStartDate.getTime() + (CONFIG.syncDays) * 24 * 60 * 60 * 1000);
  const calendarEndTime = Utilities.formatDate(syncEndDate, timezone, 'MMMM dd, yyyy 00:00:00 Z');
  syncEndDate = new Date(calendarEndTime);

  console.log(`Sync range in Timezone of Calendar: ${calendarStartTime} to ${calendarEndTime}`);
  let sourceEvents = [];
  try {
    sourceEvents = calendar.getEvents(syncStartDate, syncEndDate);
  }
  catch (e) {
    console.log(`Error while loading events from "${calendar.getId()}": ${e}`);
  }
  return sourceEvents;
}
