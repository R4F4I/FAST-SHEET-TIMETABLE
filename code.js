// copy-paste this in script.google.com
// Configuration: Replace with your actual Sheet ID and Calendar ID
const SHEET_ID = '1qQDZzc1lppWN220HVB7C3kwGtwbENvJsyummPbqzf_Y'; // Replace with your Spreadsheet ID
const CALENDAR_ID = 'c_ed42c601c32dd74fd62b775501cfd7affa0dafd895a9e48422796bf6c0aa2fd0@group.calendar.google.com'; // Replace with your Calendar ID

// --- Main Sync Function ---
function syncToCalendar() {
  Logger.log(`--- Sync process started at: ${new Date().toLocaleString()} ---`);

  // First, force the child sheet to recalculate its formulas to get the latest data.
  forceSheetRecalculation();

  const activeSpreadsheet = SpreadsheetApp.openById(SHEET_ID);
  // The original script used getSheets()[1]. If 'style2' is not the second sheet, adjust this index or use getSheetByName('style2').
  const sheet = activeSpreadsheet.getSheets()[1]; 
  if (!sheet) {
    Logger.log(`Error: Sheet at index 1 (or named 'style2') not found in Spreadsheet ID: ${SHEET_ID}. Halting sync.`);
    // SpreadsheetApp.getUi().alert("Error: Source sheet not found. Please check script configuration.");
    return;
  }
  Logger.log(`Using sheet: "${sheet.getName()}"`);

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) {
    Logger.log(`Error: Calendar with ID "${CALENDAR_ID}" not found. Halting sync.`);
    // SpreadsheetApp.getUi().alert(`Error: Calendar with ID "${CALENDAR_ID}" not found.`);
    return;
  }
  Logger.log(`Using calendar: "${calendar.getName()}"`);
  
  const data = sheet.getDataRange().getValues();
  
  // Fetch academic start and end dates from the sheet (cells G9 and H9)
  const academicTermStartDateString = sheet.getRange('G9').getDisplayValue(); // Use getDisplayValue for robustness with date formats
  const academicTermEndDateString = sheet.getRange('H9').getDisplayValue();

  if (!academicTermStartDateString || !academicTermEndDateString) {
    Logger.log('Error: Academic start date (G9) or end date (H9) is missing in the sheet. Halting sync.');
    // SpreadsheetApp.getUi().alert('Error: Academic start (G9) or end (H9) date is missing. Halting sync.');
    return;
  }

  const academicTermStartDate = new Date(academicTermStartDateString);
  const academicTermEndDate = new Date(academicTermEndDateString);

  if (isNaN(academicTermStartDate.getTime()) || isNaN(academicTermEndDate.getTime())) {
    Logger.log(`Error: Invalid academic start date ("${academicTermStartDateString}" from M15) or end date ("${academicTermEndDateString}" from O15) format in the sheet. Halting sync.`);
    // SpreadsheetApp.getUi().alert('Error: Invalid academic start or end date format in M15/O15. Halting sync.');
    return;
  }
  // Set end date to end of day for inclusive range
  academicTermEndDate.setHours(23, 59, 59, 999); 

  Logger.log(`Academic Term: Start Date = ${academicTermStartDate.toDateString()}, End Date = ${academicTermEndDate.toDateString()}`);

  // 1. Delete old events within the academic term
  deleteAllEventsInCalendar(calendar, academicTermStartDate, academicTermEndDate);

  // 2. Sync new events from the sheet to the calendar
  // Start from row 2 (index 1) if row 1 contains headers
  for (let i = 1; i < data.length; i++) {
    const rowNumber = i + 1; // For user-friendly logging (1-based)
    const rowData = data[i];

    if (!rowData || rowData.every(cell => cell === "")) {
        Logger.log(`Skipping empty or invalid row at sheet index ${i} (Row ${rowNumber}).`);
        continue;
    }

    // Destructure with care, expecting 5 columns: Day, Subject, Room, StartTime, EndTime
    const dayCellValue = rowData[0];
    const subjectCellValue = rowData[1];
    const roomCellValue = rowData[2];
    const startTimeCellValue = rowData[3];
    const endTimeCellValue = rowData[4];

    // --- Validate essential data for the row ---
    if (typeof dayCellValue !== 'string' || dayCellValue.trim() === '') {
        Logger.log(`Skipping Row ${rowNumber}: 'Day' (Column A) is missing, not a string, or empty. Value: "${dayCellValue}"`);
        continue;
    }
    if (!subjectCellValue || (typeof subjectCellValue === 'string' && subjectCellValue.trim() === '')) {
        Logger.log(`Skipping Row ${rowNumber} (Day: "${dayCellValue}"): 'Subject' (Column B) is missing or empty. Value: "${subjectCellValue}"`);
        continue;
    }
    if (!startTimeCellValue || (typeof startTimeCellValue !== 'string' && typeof startTimeCellValue !== 'number' && !(startTimeCellValue instanceof Date)) || String(startTimeCellValue).trim() === '') {
        Logger.log(`Skipping Row ${rowNumber} (Day: "${dayCellValue}", Subject: "${subjectCellValue}"): 'Start Time' (Column D) is missing or invalid. Value: "${startTimeCellValue}"`);
        continue;
    }
    if (!endTimeCellValue || (typeof endTimeCellValue !== 'string' && typeof endTimeCellValue !== 'number' && !(endTimeCellValue instanceof Date)) || String(endTimeCellValue).trim() === '') {
        Logger.log(`Skipping Row ${rowNumber} (Day: "${dayCellValue}", Subject: "${subjectCellValue}"): 'End Time' (Column E) is missing or invalid. Value: "${endTimeCellValue}"`);
        continue;
    }
    
    // --- Calculate event dates ---
    // Option 1: Use getNextDateOfWeekday (starts series from the upcoming instance of that weekday)
    const eventFirstOccurrenceStartDate = getNextDateOfWeekday(dayCellValue, startTimeCellValue);
    const eventFirstOccurrenceEndDate = getNextDateOfWeekday(dayCellValue, endTimeCellValue);

    // Option 2: Use getFirstEventDateOnOrAfter (starts series from the first instance on or after academicTermStartDate)
    // const eventFirstOccurrenceStartDate = getFirstEventDateOnOrAfter(academicTermStartDate, dayCellValue, startTimeCellValue);
    // const eventFirstOccurrenceEndDate = getFirstEventDateOnOrAfter(academicTermStartDate, dayCellValue, endTimeCellValue);

    if (!eventFirstOccurrenceStartDate || !eventFirstOccurrenceEndDate) {
      Logger.log(`Skipping event creation for Subject "${subjectCellValue}" on Row ${rowNumber} (Day: "${dayCellValue}") due to invalid date calculation (likely invalid time format or weekday name). StartTime Val: "${startTimeCellValue}", EndTime Val: "${endTimeCellValue}"`);
      continue;
    }
    
    if (eventFirstOccurrenceEndDate <= eventFirstOccurrenceStartDate) {
      Logger.log(`Skipping event for Subject "${subjectCellValue}" on Row ${rowNumber}: End time (${eventFirstOccurrenceEndDate}) is not after start time (${eventFirstOccurrenceStartDate}).`);
      continue;
    }

    // Ensure the first occurrence is not after the academic end date
    if (eventFirstOccurrenceStartDate > academicTermEndDate) {
        Logger.log(`Skipping event for Subject "${subjectCellValue}" on Row ${rowNumber}: Its first occurrence (${eventFirstOccurrenceStartDate.toLocaleString()}) is after the academic term end date (${academicTermEndDate.toLocaleString()}).`);
        continue;
    }
    
    // --- Create the event series ---
    try {
      Logger.log(`Attempting to create event series: Subject="${subjectCellValue}", Day="${dayCellValue}", First Start="${eventFirstOccurrenceStartDate.toLocaleString()}", First End="${eventFirstOccurrenceEndDate.toLocaleString()}", Recur Until="${academicTermEndDate.toDateString()}"`);
      calendar.createEventSeries(
        subjectCellValue.toString().trim(), // Ensure subject is a string
        eventFirstOccurrenceStartDate,
        eventFirstOccurrenceEndDate,
        CalendarApp.newRecurrence()
          .addWeeklyRule()
          .until(academicTermEndDate), // Recur until the end of the academic term
        {
          location: roomCellValue ? roomCellValue.toString().trim() : '',
          description: `Scheduled class for ${subjectCellValue.toString().trim()}`
        }
      );
      Logger.log(`Successfully created event series for "${subjectCellValue}" from Row ${rowNumber}.`);
    } catch (e) {
      Logger.log(`Error creating event series for Subject "${subjectCellValue}" from Row ${rowNumber}: ${e.toString()}. Details: ${e.stack ? e.stack : ''}`);
    }
  }
  Logger.log(`--- Sync process finished at: ${new Date().toLocaleString()} ---`);
  // SpreadsheetApp.getUi().alert('Calendar sync process completed. Check logs (View > Logs) for details.');
}

// --- Helper Function to Delete Events ---
function deleteAllEventsInCalendar(calendar, termStartDate, termEndDate) {
  Logger.log(`Starting deletion of events from ${termStartDate.toDateString()} to ${termEndDate.toDateString()}`);
  let deletedCount = 0;
  try {
    const events = calendar.getEvents(termStartDate, termEndDate);
    const deletedSeriesIds = new Set(); // To avoid trying to delete the same series multiple times

    Logger.log(`Found ${events.length} event instances within the date range to process for deletion.`);
    for (const event of events) {
      try {
        const series = event.getEventSeries?.(); // Optional chaining
        
        if (series) {
          const seriesId = series.getId();
          if (!deletedSeriesIds.has(seriesId)) {
            Logger.log(`Attempting to delete series: "${event.getTitle()}" (Series ID: ${seriesId})`);
            series.deleteEventSeries();
            deletedSeriesIds.add(seriesId);
            deletedCount++;
            Logger.log(`Deleted series: "${event.getTitle()}"`);
          }
        } else {
          // This is a non-recurring event or a single instance modified and detached
          Logger.log(`Attempting to delete single event: "${event.getTitle()}" (ID: ${event.getId()}) starting ${event.getStartTime().toLocaleString()}`);
          event.deleteEvent();
          deletedCount++;
          Logger.log(`Deleted single event: "${event.getTitle()}"`);
        }
      } catch (e) {
        Logger.log(`Error during deletion of event "${event.getTitle()}": ${e.toString()}. Details: ${e.stack ? e.stack : ''}. Continuing...`);
      }
    }
    Logger.log(`Finished deleting events. ${deletedCount} events/series were targeted for deletion.`);
  } catch (e) {
    Logger.log(`Major error during event deletion process: ${e.toString()}. Details: ${e.stack ? e.stack : ''}`);
  }
}

function forceSheetRecalculation() {
  try {
    const sheet = SpreadsheetApp.openById(CHILD_SHEET_ID).getSheets()[1];
    if (sheet) {
      const range = sheet.getRange('A1');
      const value = range.getValue();
      range.setValue(value);
      SpreadsheetApp.flush(); 
      Logger.log(`Forced recalculation of sheet: "${sheet.getName()}"`);
    }
  } catch (e) {
    Logger.log(`Error forcing sheet recalculation: ${e.toString()}`);
  }
}


// --- Helper Function to Calculate Next Weekday Date ---
function getNextDateOfWeekday(weekdayString, timeValue) {
  // Defensive check for the weekdayString parameter
  if (typeof weekdayString !== 'string' || weekdayString.trim() === '') {
    Logger.log(`Error in getNextDateOfWeekday: Received invalid or empty weekdayString. Value: "${weekdayString}"`);
    return null; // Return null to indicate an error
  }

  const days = {
    'SUNDAY': 0, 'MONDAY': 1, 'TUESDAY': 2, 'WEDNESDAY': 3,
    'THURSDAY': 4, 'FRIDAY': 5, 'SATURDAY': 6
  };

  const normalizedWeekday = weekdayString.trim().toUpperCase();
  const targetDay = days[normalizedWeekday];

  if (typeof targetDay === 'undefined') {
    Logger.log(`Error in getNextDateOfWeekday: Invalid weekday input "${weekdayString}" (normalized to "${normalizedWeekday}"). Not found in days mapping.`);
    return null; 
  }

  const today = new Date(); 
  const resultDate = new Date(today); // Start with today
  // Calculate how many days to add to get to the targetDay
  // If today is Monday (1) and target is Wednesday (3), diff is 2.
  // If today is Wednesday (3) and target is Monday (1), diff is (1-3+7)%7 = 5.
  const dayDifference = (targetDay - today.getDay() + 7) % 7;
  resultDate.setDate(today.getDate() + dayDifference);

  let hour, minute;

  if (timeValue instanceof Date) { // If time from sheet is already a Date object (e.g., if cell is formatted as time)
    hour = timeValue.getHours();
    minute = timeValue.getMinutes();
  } else {
    // Ensure timeValue is something that can be converted to a string
    if (typeof timeValue !== 'string' && typeof timeValue !== 'number') {
        Logger.log(`Error in getNextDateOfWeekday: timeValue is not a string or number for ${normalizedWeekday}. Value: "${timeValue}", Type: ${typeof timeValue}`);
        return null;
    }
  
    // Remove all characters except digits and the colon.
    // const timeStr = String(timeValue).replace(/[^0-9:]/g, '').trim();
    const timeStr = String(timeValue).replace(/[^0-9:]/g, '').trim().replace(/:$/, '');

    const timeComponents = timeStr.split(':');

    if (timeComponents.length !== 2) {
      Logger.log(`Error in getNext-DateOfWeekday: Invalid time format "${timeStr}" for ${normalizedWeekday} (expected H:MM or HH:MM).`);
      return null;
    }

    hour = parseInt(timeComponents[0], 10);
    minute = parseInt(timeComponents[1], 10);

    if (isNaN(hour) || isNaN(minute) || hour < 0 || hour > 23 || minute < 0 || minute > 59) {
      Logger.log(`Error in getNextDateOfWeekday: Parsed invalid hour or minute from time "${timeStr}" for ${normalizedWeekday} (Hour: ${hour}, Minute: ${minute}).`);
      return null;
    }

    // Apply AM/PM inference based on your sheet's specific convention:
    // Hours 1-7 are considered PM (e.g., 1:00 is 13:00, 7:00 is 19:00).
    // Hours 8-11 are AM.
    // Hour 12 is PM (noon, 12:00).
    // Hour 0 is AM (midnight, 00:00).
    // Hours 13-23 are already in 24-hour PM.
    if (hour >= 1 && hour <= 7) { 
        hour += 12; // Convert 1:xx PM to 7:xx PM to 24-hour format
    }
    // Note: 12:xx is already correctly handled as 12 PM (noon) by default in 24h.
    // 8:xx, 9:xx, 10:xx, 11:xx are correctly handled as AM.
  }
  
  resultDate.setHours(hour, minute, 0, 0); // Set seconds and milliseconds to 0
  return resultDate;
}

// --- OPTIONAL: Helper Function to Calculate First Event Date On or After Academic Start ---
function getFirstEventDateOnOrAfter(academicStartDate, weekdayString, timeValue) {
  if (typeof weekdayString !== 'string' || weekdayString.trim() === '') {
    Logger.log(`Error in getFirstEventDateOnOrAfter: Received invalid or empty weekdayString. Value: "${weekdayString}"`);
    return null;
  }

  const days = {
    'SUNDAY': 0, 'MONDAY': 1, 'TUESDAY': 2, 'WEDNESDAY': 3,
    'THURSDAY': 4, 'FRIDAY': 5, 'SATURDAY': 6
  };
  const normalizedWeekday = weekdayString.trim().toUpperCase();
  const targetDayOfWeek = days[normalizedWeekday];

  if (typeof targetDayOfWeek === 'undefined') {
    Logger.log(`Error in getFirstEventDateOnOrAfter: Invalid weekday input "${weekdayString}".`);
    return null;
  }

  let resultDate = new Date(academicStartDate); // Start from the academic term start date
  
  // Adjust resultDate to the first targetDayOfWeek on or after academicStartDate
  let currentDayOfWeek = resultDate.getDay();
  let dayDifference = (targetDayOfWeek - currentDayOfWeek + 7) % 7;
  resultDate.setDate(resultDate.getDate() + dayDifference);

  // Time parsing (same logic as getNextDateOfWeekday)
  let hour, minute;
  if (timeValue instanceof Date) {
    hour = timeValue.getHours();
    minute = timeValue.getMinutes();
  } else {
    if (typeof timeValue !== 'string' && typeof timeValue !== 'number') {
        Logger.log(`Error in getFirstEventDateOnOrAfter: timeValue is not a string or number for ${normalizedWeekday}. Value: "${timeValue}", Type: ${typeof timeValue}`);
        return null;
    }
    const timeStr = String(timeValue).trim();
    const timeComponents = timeStr.split(':');
    if (timeComponents.length !== 2) {
      Logger.log(`Error in getFirstEventDateOnOrAfter: Invalid time format "${timeStr}" for ${normalizedWeekday}.`);
      return null;
    }
    hour = parseInt(timeComponents[0], 10);
    minute = parseInt(timeComponents[1], 10);

    if (isNaN(hour) || isNaN(minute) || hour < 0 || hour > 23 || minute < 0 || minute > 59) {
      Logger.log(`Error in getFirstEventDateOnOrAfter: Parsed invalid hour/minute from "${timeStr}" for ${normalizedWeekday}.`);
      return null;
    }
    if (hour >= 1 && hour <= 7) { // 1 PM to 7 PM
        hour += 12;
    }
  }
  resultDate.setHours(hour, minute, 0, 0);
  return resultDate;
}


// --- Trigger Management Functions ---
function createWeeklyTrigger() {
  // Delete any existing triggers for 'syncToCalendar' to avoid duplicates
  deleteTriggers('syncToCalendar');

  // Create a new time-driven trigger
  ScriptApp.newTrigger('syncToCalendar')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.SUNDAY) // Runs every Sunday
    .atHour(21) // At 9 PM (adjust as needed, based on your timezone settings in Apps Script project)
    .create();
  Logger.log('Weekly trigger for syncToCalendar created for Sunday at 9 PM.');
  // SpreadsheetApp.getUi().alert('Weekly trigger for syncToCalendar created/updated for Sunday at 9 PM.');
}

function deleteTriggers(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  for (const trigger of triggers) {
    if (functionName && trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
      Logger.log(`Deleted trigger for function "${functionName}" (ID: ${trigger.getUniqueId()})`);
    } else if (!functionName) { // If no functionName provided, delete all triggers for this project
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
      Logger.log(`Deleted project trigger (ID: ${trigger.getUniqueId()})`);
    }
  }
  if (deletedCount > 0) {
    Logger.log(`${deletedCount} trigger(s) processed for deletion.`);
  } else {
    Logger.log(`No triggers found matching criteria for deletion (Function: ${functionName ? functionName : 'Any'}).`);
  }
}

// --- Utility function to manually run or test deleting all triggers ---
function deleteAllProjectTriggers() {
    deleteTriggers(null); // Pass null to delete all triggers for this project
    // SpreadsheetApp.getUi().alert('All project triggers have been deleted.');
}

// --- Optional: onChange trigger (can be very resource-intensive) ---
// function onChange(e) {
//   // This function will run every time an edit is made to the spreadsheet.
//   // It can be very frequent and might hit quota limits if not handled carefully.
//   // Consider if you really need real-time sync or if a periodic sync (e.g., weekly) is sufficient.
//   Logger.log(`onChange event detected: ${e.changeType}`);
//   // Add conditions here if you only want to sync on specific changes
//   // For example, if (e.changeType === 'EDIT' || e.changeType === 'INSERT_ROW' || e.changeType === 'REMOVE_ROW') {
//   //   syncToCalendar();
//   // }
//   // For now, let's call it directly, but be cautious.
//   // syncToCalendar(); 
// }
