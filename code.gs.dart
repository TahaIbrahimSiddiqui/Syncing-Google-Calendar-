/**
 * Multi-Calendar Sync Script for Google Apps Script
 * Syncs events from multiple source calendars to one destination calendar as "Busy" events
 * CLEAN VERSION - No source info, prevents duplicates, only syncs new events
 */

// Configuration - Update these settings
const CONFIG = {
  // Array of source calendar IDs to sync FROM
  sourceCalendarIds: [
    'xyz@gmail.com',  // Replace with email ids to sync calendar from (be sure that you already shared the calendar with your work id, you're runing this script in your work id)
    'xyz@ashoka.edu.in',  // Replace with actual email id
    // Add more calendar email IDs as needed
  ],
  
  // Destination calendar ID to sync TO (use 'primary' for your main calendar)
  destinationCalendarId: 'primary', // or 'your-email@gmail.com'
  
  // Number of days to look ahead for events
  daysToSync: 30,
  
  // Whether to sync all-day events
  syncAllDayEvents: true,
  
  // Whether to update existing synced events if source changes
  updateExistingEvents: true
};

/**
 * Main function to sync calendars
 * Run this function manually or set up a trigger
 */
function syncCalendars() {
  try {
    console.log('Starting calendar sync...');
    
    const startDate = new Date();
    const endDate = new Date();
    endDate.setDate(startDate.getDate() + CONFIG.daysToSync);
    
    const destinationCalendar = CalendarApp.getCalendarById(CONFIG.destinationCalendarId);
    if (!destinationCalendar) {
      throw new Error('Destination calendar not found');
    }
    
    // Get existing "Busy" events to avoid duplicates
    const existingBusyEvents = getExistingBusyEvents(destinationCalendar, startDate, endDate);
    console.log(`Found ${existingBusyEvents.size} existing "Busy" events`);
    
    let totalEventsSynced = 0;
    let totalEventsSkipped = 0;
    
    // Process each source calendar
    CONFIG.sourceCalendarIds.forEach(calendarId => {
      try {
        const sourceCalendar = CalendarApp.getCalendarById(calendarId);
        if (!sourceCalendar) {
          console.warn(`Source calendar ${calendarId} not found or not accessible`);
          return;
        }
        
        console.log(`Processing calendar: ${sourceCalendar.getName()}`);
        
        const events = sourceCalendar.getEvents(startDate, endDate);
        let calendarEventsSynced = 0;
        let calendarEventsSkipped = 0;
        
        events.forEach(event => {
          const result = syncEvent(event, destinationCalendar, existingBusyEvents);
          if (result === 'synced') {
            calendarEventsSynced++;
          } else if (result === 'skipped') {
            calendarEventsSkipped++;
          }
        });
        
        console.log(`Calendar ${sourceCalendar.getName()}: ${calendarEventsSynced} synced, ${calendarEventsSkipped} skipped`);
        totalEventsSynced += calendarEventsSynced;
        totalEventsSkipped += calendarEventsSkipped;
        
      } catch (error) {
        console.error(`Error processing calendar ${calendarId}:`, error.message);
      }
    });
    
    console.log(`Sync completed: ${totalEventsSynced} new events synced, ${totalEventsSkipped} duplicates skipped`);
    
  } catch (error) {
    console.error('Calendar sync failed:', error.message);
    throw error;
  }
}

/**
 * Sync a single event from source to destination calendar
 */
function syncEvent(sourceEvent, destinationCalendar, existingBusyEvents) {
  try {
    // Skip if all-day events are disabled and this is an all-day event
    if (!CONFIG.syncAllDayEvents && sourceEvent.isAllDayEvent()) {
      return 'ignored';
    }
    
    const eventKey = generateEventKey(sourceEvent);
    const existingEvent = existingBusyEvents.get(eventKey);
    
    // If event already exists, check if we should update it
    if (existingEvent) {
      if (!CONFIG.updateExistingEvents) {
        return 'skipped';
      }
      
      // Only update if source event was modified after the existing event
      if (sourceEvent.getLastUpdated() <= existingEvent.getLastUpdated()) {
        return 'skipped';
      }
      
      // Update existing event
      updateExistingEvent(existingEvent, sourceEvent);
      console.log(`Updated: ${sourceEvent.getTitle()}`);
      return 'updated';
    } else {
      // Create new "Busy" event
      createBusyEvent(destinationCalendar, sourceEvent);
      console.log(`Created: ${sourceEvent.getTitle()}`);
      return 'synced';
    }
    
  } catch (error) {
    console.error(`Error syncing event "${sourceEvent.getTitle()}":`, error.message);
    return 'error';
  }
}

/**
 * Create a new "Busy" event with no source information
 */
function createBusyEvent(destinationCalendar, sourceEvent) {
  let newEvent;
  
  if (sourceEvent.isAllDayEvent()) {
    newEvent = destinationCalendar.createAllDayEvent(
      "Busy",
      sourceEvent.getAllDayStartDate(),
      sourceEvent.getAllDayEndDate(),
      {
        description: "" // Completely empty description
      }
    );
  } else {
    newEvent = destinationCalendar.createEvent(
      "Busy",
      sourceEvent.getStartTime(),
      sourceEvent.getEndTime(),
      {
        description: "" // Completely empty description
      }
    );
  }
  
  // Set the event to show as busy
  newEvent.setTransparency(CalendarApp.EventTransparency.OPAQUE);
  newEvent.setVisibility(CalendarApp.Visibility.DEFAULT);
}

/**
 * Update an existing "Busy" event
 */
function updateExistingEvent(existingEvent, sourceEvent) {
  // Keep the title as "Busy" and description empty
  existingEvent.setTitle("Busy");
  existingEvent.setDescription("");
  
  // Update the time/date
  if (sourceEvent.isAllDayEvent()) {
    existingEvent.setAllDayDates(
      sourceEvent.getAllDayStartDate(),
      sourceEvent.getAllDayEndDate()
    );
  } else {
    existingEvent.setTime(
      sourceEvent.getStartTime(),
      sourceEvent.getEndTime()
    );
  }
  
  // Ensure the updated event shows as busy
  existingEvent.setTransparency(CalendarApp.EventTransparency.OPAQUE);
  existingEvent.setVisibility(CalendarApp.Visibility.DEFAULT);
}

/**
 * Generate a unique key for an event based on time slots
 */
function generateEventKey(event) {
  let startTime, endTime;
  
  if (event.isAllDayEvent()) {
    startTime = event.getAllDayStartDate().toDateString();
    endTime = event.getAllDayEndDate().toDateString();
  } else {
    startTime = event.getStartTime().toISOString();
    endTime = event.getEndTime().toISOString();
  }
  
  // Key is based purely on time slot, not event title
  return `${startTime}|${endTime}`;
}

/**
 * Get existing "Busy" events to avoid duplicates
 * Only looks at time slots, not titles or descriptions
 */
function getExistingBusyEvents(calendar, startDate, endDate) {
  const events = calendar.getEvents(startDate, endDate);
  const busyEvents = new Map();
  
  events.forEach(event => {
    // Only consider events titled "Busy" with empty descriptions as our synced events
    if (event.getTitle() === "Busy" && (event.getDescription() === "" || event.getDescription() === null)) {
      
      // Create key based on time slot
      let startTime, endTime;
      
      if (event.isAllDayEvent()) {
        startTime = event.getAllDayStartDate().toDateString();
        endTime = event.getAllDayEndDate().toDateString();
      } else {
        startTime = event.getStartTime().toISOString();
        endTime = event.getEndTime().toISOString();
      }
      
      const key = `${startTime}|${endTime}`;
      busyEvents.set(key, event);
    }
  });
  
  return busyEvents;
}

/**
 * Setup automatic sync trigger (run this once to set up automation)
 */
function setupTrigger() {
  // Delete existing triggers for this function
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'syncCalendars') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger to run every hour
  ScriptApp.newTrigger('syncCalendars')
    .timeBased()
    .everyHours(1)
    .create();
    
  console.log('Trigger setup complete. Sync will run every hour.');
}

/**
 * Remove all "Busy" events with empty descriptions (cleanup function)
 */
function removeAllBusyEvents() {
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - 30); // Look back 30 days too
  const endDate = new Date();
  endDate.setDate(endDate.getDate() + CONFIG.daysToSync);
  
  const destinationCalendar = CalendarApp.getCalendarById(CONFIG.destinationCalendarId);
  const events = destinationCalendar.getEvents(startDate, endDate);
  
  let removedCount = 0;
  events.forEach(event => {
    // Remove "Busy" events with empty descriptions
    if (event.getTitle() === "Busy" && (event.getDescription() === "" || event.getDescription() === null)) {
      event.deleteEvent();
      removedCount++;
    }
  });
  
  console.log(`Removed ${removedCount} "Busy" events`);
}

/**
 * Test function to check calendar access and show sample events
 */
function testCalendarAccess() {
  console.log('Testing calendar access...');
  
  CONFIG.sourceCalendarIds.forEach((calendarId, index) => {
    try {
      const calendar = CalendarApp.getCalendarById(calendarId);
      if (calendar) {
        console.log(`✓ Source ${index + 1}: ${calendar.getName()} (${calendarId})`);
        
        // Test getting events
        const testStart = new Date();
        const testEnd = new Date();
        testEnd.setDate(testStart.getDate() + 7);
        const events = calendar.getEvents(testStart, testEnd);
        console.log(`  - Found ${events.length} events in next 7 days`);
        
        // Show first few events as examples
        events.slice(0, 3).forEach(event => {
          console.log(`    - "${event.getTitle()}" on ${event.getStartTime()}`);
        });
        
      } else {
        console.log(`✗ Cannot access calendar: ${calendarId}`);
      }
    } catch (error) {
      console.log(`✗ Error accessing calendar ${calendarId}: ${error.message}`);
    }
  });
  
  try {
    const destCalendar = CalendarApp.getCalendarById(CONFIG.destinationCalendarId);
    console.log(`✓ Destination calendar: ${destCalendar.getName()}`);
    
    // Count existing "Busy" events
    const testStart = new Date();
    const testEnd = new Date();
    testEnd.setDate(testStart.getDate() + 7);
    const existingEvents = destCalendar.getEvents(testStart, testEnd);
    const busyCount = existingEvents.filter(e => e.getTitle() === "Busy" && (e.getDescription() === "" || e.getDescription() === null)).length;
    console.log(`  - Currently has ${busyCount} "Busy" events`);
    
  } catch (error) {
    console.log(`✗ Cannot access destination calendar: ${error.message}`);
  }
}

/**
 * Show statistics about synced events
 */
function showSyncStats() {
  const startDate = new Date();
  const endDate = new Date();
  endDate.setDate(startDate.getDate() + CONFIG.daysToSync);
  
  const destinationCalendar = CalendarApp.getCalendarById(CONFIG.destinationCalendarId);
  const events = destinationCalendar.getEvents(startDate, endDate);
  
  const busyEvents = events.filter(e => e.getTitle() === "Busy" && (e.getDescription() === "" || e.getDescription() === null));
  
  console.log(`\n=== SYNC STATISTICS ===`);
  console.log(`Total events in destination: ${events.length}`);
  console.log(`"Busy" synced events: ${busyEvents.length}`);
  console.log(`Next 5 "Busy" events:`);
  
  busyEvents.slice(0, 5).forEach((event, index) => {
    const timeStr = event.isAllDayEvent() ? 
      `All day on ${event.getAllDayStartDate().toDateString()}` :
      `${event.getStartTime()} - ${event.getEndTime()}`;
    console.log(`  ${index + 1}. ${timeStr}`);
  });
  console.log(`========================\n`);
}