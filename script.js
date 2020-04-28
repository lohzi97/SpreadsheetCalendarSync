/**
 * Person-in-charge:
 * - Loh Zi Jian, lohzi97@gmail.com
 */

/**
 * Define parameter.
 * - All these need to be defined by user.
 */
var spreadSheetURL = "https://docs.google.com/spreadsheets/d/1XdYDfoyke-NvJOv9T_4iK6hq3pWtUu74-Sl5WJ7dHZs/edit#gid=0";
var calendarID = "ttj425k3bbo0ktrns2tsvq8ln8@group.calendar.google.com";
var headerColor = "#4a86e8";
var sheetNames = ["Sheet1", "Sheet2"];
var headerString = [
  {
    'date': "Date",
    'time': "Time",
    'recurrence': "Recurrence",
    'titles': ["Blog Title", "FB Post Title 1"],
    'ids': ['CalendarEventID - Blog', 'CalendarEventID - FB'],
    'prefix': ["Blog", "FB"]
  },
  {
    'date': "D",
    'time': "T",
    'recurrence': "Re",
    'titles': ["Articles", "Stories"],
    'ids': ['CalendarEventID - Arti', 'CalendarEventID - Sto'],
    'prefix': ["Ar", "St"]
  }
];
var noSyncString = "NOSYNC";
var syncPeriod = {
  'yearBefore': 1,
  'yearAfter': 3
};
var syncMinutes = 15;
var identificationString = "Calendar Sync";

/**
 * A special function that runs when the spreadsheet is first
 * opened or reloaded. onOpen() is used to add custom menu
 * items to the spreadsheet.
 */
function onOpen() {

  // Get the UI element to create menu.
  let ui = SpreadsheetApp.getUi();

  // Check whether the auto sync trigger has been created.
  let allTriggers = ScriptApp.getProjectTriggers();
  if (allTriggers.length === 0) {
    ui.createMenu('Sync Calendar')
      .addItem('Sync To', 'syncToWrapper')
      .addItem('Sync From', 'syncFrom')
      .addSeparator()
      .addItem('Clear All', 'clearAll')
      .addItem('Enable Auto Sync', 'enableTrigger')
      .addToUi();
    return;
  }
  else {
    for (const trigger of allTriggers) {
      if (trigger.getHandlerFunction() !== "syncTo_") {
        ui.createMenu('Sync Calendar')
          .addItem('Sync To', 'syncToWrapper')
          .addItem('Sync From', 'syncFrom')
          .addSeparator()
          .addItem('Clear All', 'clearAll')
          .addItem('Enable Auto Sync', 'enableTrigger')
          .addToUi();
        return;
      }
    }
  }

  ui.createMenu('Sync Calendar')
    .addItem('Sync To', 'syncToWrapper')
    .addItem('Sync From', 'syncFrom')
    .addSeparator()
    .addItem('Clear All', 'clearAll')
    .addItem('Disable Auto Sync', 'deleteTrigger')
    .addToUi();
    
}

/**
 * Function that create the auto sync trigger. It will sync even if user has closed the spreadsheet.
 */
function createTrigger_() {
  ScriptApp.newTrigger('syncTo_')
    .timeBased()
    .everyMinutes(syncMinutes)
    .create();
}

/**
 * Function that deletes the auto sync trigger.
 */
function deleteTrigger() {

  // Loop through all the available trigger
  let allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    if (trigger.getHandlerFunction() === "syncTo_") {

      // delete the trigger.
      ScriptApp.deleteTrigger(trigger);

      // update the menu
      let ui = SpreadsheetApp.getUi();
      ui.createMenu('Sync Calendar')
        .addItem('Sync To', 'syncToWrapper')
        .addItem('Sync From', 'syncFrom')
        .addSeparator()
        .addItem('Clear All', 'clearAll')
        .addItem('Enable Auto Sync', 'enableTrigger')
        .addToUi();

      // Show a dialog box to indicate auto sync disabled.
      SpreadsheetApp.getUi().alert(`Auto sync disabled.`);

      return;

    }
  }

}

function enableTrigger() {
  try {

    // Create the auto sync trigger.
    createTrigger_();

    // Update the menu
    let ui = SpreadsheetApp.getUi();
    ui.createMenu('Sync Calendar')
      .addItem('Sync To', 'syncToWrapper')
      .addItem('Sync From', 'syncFrom')
      .addSeparator()
      .addItem('Clear All', 'clearAll')
      .addItem('Disable Auto Sync', 'deleteTrigger')
      .addToUi();

    // Show a dialog box to indicate auto sync enabled.
    SpreadsheetApp.getUi().alert(`Auto sync enabled.`);

  } catch (error) {
    
    SpreadsheetApp.getUi().alert(error);

  }
}

/**
 * Wrapper function that wrap up the syncTo function with try catch and add UI to it
 */
function syncToWrapper() {

  try {

    // Run the syncTo function.
    syncTo_();
    
    // Show a dialog box to indicate sync has completed.
    SpreadsheetApp.getUi().alert(`Sync Complete.`);

  } catch (error) {

    SpreadsheetApp.getUi().alert(error);

  }

}

/**
 * Function that update the calendar with spreadsheet.
 */
function syncTo_() {

  // Get the calendar
  let calendar = CalendarApp.getCalendarById(calendarID);
  if (calendar === null) {
    throw new Error(`Failed to get your calendar: ${calendarID}.`);
  }

  // Get an array of event that previously has set in calendar
  let calendarEvents = getCalenderEvents_(calendar);
  
  // Get an array of event that defined in all the sheets.
  let sheetEvents = [];
  for (let s = 0; s < sheetNames.length; s++) {
   
    // Get the sheet
    let sheet = SpreadsheetApp.openByUrl(spreadSheetURL).getSheetByName(sheetNames[s]);
    if (sheet === null) {
      throw new Error(`Unable to find ${sheetName} in ${spreadSheetURL}.`);
    }
    
    // Get an array of event that defined in the sheet
    let eventArray = getSheetEvents_(sheet, s);
    sheetEvents = sheetEvents.concat(eventArray);

  }
    
  // Compare sheetEvents and calendarEvents, extract those isn't exactly same.
  let diffEvents = compareEvents_(sheetEvents, calendarEvents);

  // From the diffEvents, perform the following to sync sheet to calendar:
  // - create event if "belong" field is "sheet"
  // - delete event if "belong" field is "calendar"
  // - update event if "belong" field is "sheet&"
  // - delete the event series and recreate an event if "belong" field is "sheet&DR"
  // - delete the event and recreate an event series if "belong" field is "sheetAR"
  let createdEvents = [];
  let deletedEvents = [];
  for (let i = 0; i < diffEvents.length; i++) {
    let event = diffEvents[i];
    if (event.belong === "sheet") {
      if (event.recurrence === "null") {
        let createdE = calendar.createEvent(
          event.title,
          event.startTime,
          event.endTime,
          {'description': event.description}
        ).setTag('identification', `Created by "${identificationString}" spreadsheet.`);
        createdEvents.push(calanderToEvent_(createdE));
      }
      else {
        let createdE = calendar.createEventSeries(
          event.title,
          event.startTime,
          event.endTime,
          formRecurrenceRule_(event),
          {'description': event.description}
        ).setTag('identification', `Created by "${identificationString}" spreadsheet.`).setTag('recurrence', JSON.stringify(event.recurrence));
        let sT = new Date(event.startTime.getTime() - (5 * 60 * 1000));
        let eT = new Date(event.endTime.getTime() + (5 * 60 * 1000));
        let searchStr = event.title.split(" - ")[1].split(" ")[0];
        let firstE = calendar.getEvents(sT, eT, {'search': searchStr});
        createdEvents.push(calanderToEvent_(firstE[0]));
      }
    }
    else if (event.belong === "calendar") {
      
      // Find the event or event series and delete it.
      if (event.recurrence === "null") {
        let calEvent = calendar.getEventById(event.id);
        calEvent.deleteEvent();
      }
      else {
        let calEvent = calendar.getEventById(event.id);
        let calEventSeries = calEvent.getEventSeries();
        calEventSeries.deleteEventSeries();
      }
      deletedEvents.push(event);

    }
    else if (event.belong === "sheet&") {
      if (event.recurrence === "null") {
        let calEvent = calendar.getEventById(event.id);
        let eObj = calanderToEvent_(calEvent);
        if (event.title !== eObj.title) {
          calEvent.setTitle(event.title);
        }
        if (event.startTime.getTime() !== eObj.startTime.getTime() || event.endTime.getTime() !== eObj.endTime.getTime()) {
          calEvent.setTime(event.startTime, event.endTime);
        }
      }
      else {
        let calEvent = calendar.getEventById(event.id);
        let calEventSeries = calEvent.getEventSeries();
        let eSeriesObj = calanderToEvent_(calEvent);
        if (event.title !== eSeriesObj.title) {
          calEventSeries.setTitle(event.title);
        }
        if (
          event.startTime.getTime() !== eSeriesObj.startTime.getTime() || 
          event.endTime.getTime() !== eSeriesObj.endTime.getTime() ||
          event.recurrence.rule !== eSeriesObj.recurrence.rule ||
          event.recurrence.repeatTimes !== eSeriesObj.recurrence.repeatTimes ||
          event.recurrence.end !== eSeriesObj.recurrence.end ||
          event.recurrence.endTimes !== eSeriesObj.recurrence.endTimes ||
          event.recurrence.endDate.year !== eSeriesObj.recurrence.endDate.year ||
          event.recurrence.endDate.month !== eSeriesObj.recurrence.endDate.month ||
          event.recurrence.endDate.day !== eSeriesObj.recurrence.endDate.day ||
          event.recurrence.endTimes !== eSeriesObj.recurrence.endTimes ||
          event.recurrence.repeatMode !== eSeriesObj.recurrence.repeatMode ||
          !arrayIsEqual_(event.recurrence.repeatOn, eSeriesObj.recurrence.repeatOn)
        ) {

          // Supposingly we should use this line of code to update it. 
          // But seems like it cannot properly delete the old event when we tried to update its start date. 
          // So use back the delete then create method.
          // This line put here as a referrence. If future this issue has been fixed by google. This should be the preferred method,
          // because we actually have limits on creating events, and i believe update will always be faster than delete and create.

          // calEventSeries.setRecurrence(formRecurrenceRule_(event), event.startTime, event.endTime);
          
          calEventSeries.deleteEventSeries();
          deletedEvents.push(event);

          let createdE = calendar.createEventSeries(
            event.title,
            event.startTime,
            event.endTime,
            formRecurrenceRule_(event),
            {'description': event.description}
          ).setTag('identification', `Created by "${identificationString}" spreadsheet.`).setTag('recurrence', JSON.stringify(event.recurrence));
          let sT = new Date(event.startTime.getTime() - (5 * 60 * 1000));
          let eT = new Date(event.endTime.getTime() + (5 * 60 * 1000));
          let searchStr = event.title.split(" - ")[1].split(" ")[0];
          let firstE = calendar.getEvents(sT, eT, {'search': searchStr});
          createdEvents.push(calanderToEvent_(firstE[0]));

        }
      }
    }
    else if (event.belong === "sheet&DR") {
      let calEvent = calendar.getEventById(event.id);
      let calEventSeries = calEvent.getEventSeries();
      calEventSeries.deleteEventSeries();
      deletedEvents.push(event);
      let createdE = calendar.createEvent(
        event.title,
        event.startTime,
        event.endTime,
        {'description': event.description}
      ).setTag('identification', `Created by "${identificationString}" spreadsheet.`);
      createdEvents.push(calanderToEvent_(createdE));
    }
    else if (event.belong === "sheet&AR") {
      let calEvent = calendar.getEventById(event.id);
      calEvent.deleteEvent();
      deletedEvents.push(event);
      let createdE = calendar.createEventSeries(
        event.title,
        event.startTime,
        event.endTime,
        formRecurrenceRule_(event),
        {'description': event.description}
      ).setTag('identification', `Created by "${identificationString}" spreadsheet.`).setTag('recurrence', JSON.stringify(event.recurrence));
      let sT = new Date(event.startTime.getTime() - (5 * 60 * 1000));
      let eT = new Date(event.endTime.getTime() + (5 * 60 * 1000));
      let searchStr = event.title.split(" - ")[1].split(" ")[0];
      let firstE = calendar.getEvents(sT, eT, {'search': searchStr});
      createdEvents.push(calanderToEvent_(firstE[0]));
    }
  }

  // Begin modify the spreadsheet. It includes:
  // - adding Calendar Event IDs
  for (let s = 0; s < sheetNames.length; s++) {

    // Get the sheet
    let sheet = SpreadsheetApp.openByUrl(spreadSheetURL).getSheetByName(sheetNames[s]);

    // Find the header row, which is defined by a specific background color.
    let headerRow = getHeaderRow_(sheet);

    // Find the column that contain the info of the calendar events.
    let infoCol = getInfoColumn_(sheet, s, headerRow);

    // Get all the relevant data in the spreadsheet.
    let dataA1Notation = (headerRow+1).toString() + ":" + sheet.getLastRow().toString();
    let dataRange = sheet.getRange(dataA1Notation);
    let data = dataRange.getValues();

    // Loop through the data row by row, and check if that row requires edit or not.
    for (let i = 0; i < data.length; i++) {
      
      // If the Event ID matched one of the deleted event's id, then delete the Event ID.
      for (let j = 0; j < infoCol.ids.length; j++) {
        for (const event of deletedEvents) {
          if (data[i][infoCol.ids[j]] === event.id) {
            let cellA1Notation = lettersFromIndex_(infoCol.ids[j]) + (i+headerRow+1).toString();
            let cell = sheet.getRange(cellA1Notation);
            cell.clearContent();
            break;
          } 
        }
      }

      // If the title, startTime, endTime and recurrent of that row is same with one of the created event,
      // add event.id to the Event ID column.
      let eventFromRow = sheetRowToEvent_(data[i], infoCol, i, s);
      for (const e of eventFromRow) {
        for (const event of createdEvents) {
          if (
            (
              e.title === event.title &&
              e.startTime.getTime() === event.startTime.getTime() &&
              e.endTime.getTime() === event.endTime.getTime() &&
              e.recurrence === "null" && event.recurrence === "null"
            ) 
            ||
            (
              e.title === event.title &&
              e.startTime.getTime() === event.startTime.getTime() &&
              e.endTime.getTime() === event.endTime.getTime() &&
              e.recurrence.rule === event.recurrence.rule &&
              e.recurrence.repeatTimes === event.recurrence.repeatTimes &&
              e.recurrence.end === event.recurrence.end &&
              e.recurrence.endTimes === event.recurrence.endTimes &&
              e.recurrence.endDate.year === event.recurrence.endDate.year &&
              e.recurrence.endDate.month === event.recurrence.endDate.month &&
              e.recurrence.endDate.day === event.recurrence.endDate.day &&
              e.recurrence.repeatMode === event.recurrence.repeatMode &&
              arrayIsEqual_(e.recurrence.repeatOn, event.recurrence.repeatOn)
            )
          ) {
            // Find the column of title. 
            let titleCol;
            let splitedTitle = event.title.split(" - ");
            if (splitedTitle.length > 2) {
              splitedTitle[1] = splitedTitle.slice(1).join(' - ')
            }
            for (let j = 0; j < data[i].length; j++) {
              if (data[i][j] === splitedTitle[1]) {
                titleCol = j;
              }
            }
            // Check which index in headerString.titles that title is in.
            let idx;
            for (let j = 0; j < infoCol.titles.length; j++) {
              if (infoCol.titles[j] === titleCol) {
                idx = j;
              }
            }
            // Add the event id to the id column.
            let idCol = infoCol.ids[idx];
            let cellA1Notation = lettersFromIndex_(idCol) + (i+headerRow+1).toString();
            let cell = sheet.getRange(cellA1Notation);
            cell.setValue(event.id);
            
          }
        }
      }

    }

    protectIDColumn_(sheet, headerRow, infoCol);

  }
  
  return;
  
}

/**
 * Function that update the spreadsheet with calendar.
 */
function syncFrom() {
  
  // Show prompt to user indicating that this will clear everything, and ask them to confirm on continue.
  let ui = SpreadsheetApp.getUi();
  let response = ui.alert('WARNING! \nSync From calendar will overwrite certain portion of your sheet. \nAre you sure you want to continue?', ui.ButtonSet.YES_NO);

  if (response === ui.Button.NO) {
    return;
  }

  try {

    // Get the calendar
    let calendar = CalendarApp.getCalendarById(calendarID);
    if (calendar === null) {
      throw new Error(`Failed to get your calendar: ${calendarID}.`);
    }

    // Get an array of event that previously has set in calendar
    let calendarEvents = getCalenderEvents_(calendar);

    // Get an array of event that defined in the sheets.
    let sheetEvents = [];
    for (let s = 0; s < sheetNames.length; s++) {
     
      // Get the sheet
      let sheet = SpreadsheetApp.openByUrl(spreadSheetURL).getSheetByName(sheetNames[s]);
      if (sheet === null) {
        throw new Error(`Unable to find ${sheetName} in ${spreadSheetURL}.`);
      }

      // Get an array of event that defined in the sheet
      let eventArray = getSheetEvents_(sheet, s);
      sheetEvents = sheetEvents.concat(eventArray);

    }

    // Compare sheetEvents and calendarEvents, extract those isn't exactly same.
    let diffEvents = compareEvents_(sheetEvents, calendarEvents);

    for (let s = 0; s < sheetNames.length; s++) {

      // Get the sheet
      let sheet = SpreadsheetApp.openByUrl(spreadSheetURL).getSheetByName(sheetNames[s]);

      // Find the header row, which is defined by a specific background color.
      let headerRow = getHeaderRow_(sheet);

      // Find the column that contain the info of the calendar events.
      let infoCol = getInfoColumn_(sheet, s, headerRow);

      // From the diffEvents, perform the following to sync calendar to sheet:
      // - create a new row and add info to it if "belong" field is "calendar"
      // - add noSyncString to title if "belong" field is "sheet"
      // - update the row info if "belong" field is "calendar&", "calendar&AR" "calendar&DR"
      for (let z = 0; z < diffEvents.length; z++) {
        
        let event = diffEvents[z];
        
        if (event.belong === "calendar") {

          // Convert the event object into a spreadsheet row. 
          let rowOfEvent = eventToSheetRow_(event, infoCol, s);

          // Check if the title column is empty. If it is empty, means that this is not the sheet that it should append to.
          // Skip it.
          if (titlesIsEmpty(rowOfEvent, infoCol)) {
            continue;
          }

          // Append it to the end of the spreadsheet.
          sheet.appendRow(rowOfEvent);

        }
        else if (event.belong === "sheet") {

          // Find the index that we need to use to set the Title and EventID.
          let splitedTitle = event.title.split(" - ");
          if (splitedTitle.length > 2) {
            splitedTitle[1] = splitedTitle.slice(1).join(' - ');
          }
          let idx = headerString[s].prefix.indexOf(splitedTitle[0]);

          // Find the row that the event is in, by looping through all the data and comparing the event title. We does not compare
          // all the info because the generated rowOfEvent may be a bit different than the original one. But still use title because
          // it is the least modified one. 
          let dataA1Notation = (headerRow+1).toString() + ":" + sheet.getLastRow().toString();
          let data = sheet.getRange(dataA1Notation).getValues();
          for (let i = 0; i < data.length; i++) {
            if (data[i][infoCol.titles[idx]] === splitedTitle[1]) {

              // Add the noSyncString in front of the title.
              let cellA1Notation = lettersFromIndex_(infoCol.titles[idx]) + (i+headerRow+1).toString();
              let cell = sheet.getRange(cellA1Notation);
              let cellValue = cell.getValue()
              cell.setValue(noSyncString + " " + cellValue);
              break;

            }
          }

        }
        else if (
          event.belong === "calendar&" ||
          event.belong === "calendar&AR" ||
          event.belong === "calendar&DR"
        ) {

          // Convert the event object into a spreadsheet row. 
          let rowOfEvent = eventToSheetRow_(event, infoCol, s);

          // Find the index that we need to use to set the Title and EventID.
          let splitedTitle = event.title.split(" - ");
          if (splitedTitle.length > 2) {
            splitedTitle[1] = splitedTitle.slice(1).join(' - ')
          }
          let idx = headerString[s].prefix.indexOf(splitedTitle[0]);

          // Find the row that the event is in, by looping through all the data and comparing the event id.
          let dataA1Notation = (headerRow+1).toString() + ":" + sheet.getLastRow().toString();
          let data = sheet.getRange(dataA1Notation).getValues();
          for (let i = 0; i < data.length; i++) {
            if (data[i][infoCol.ids[idx]] === event.id) {

              // Replace everything in the row with the generated rowOfEvent content.
              for (let j = 0; j < rowOfEvent.length; j++) {
                if (rowOfEvent[j] !== "") {
                  let cellA1Notation = lettersFromIndex_(j) + (i+headerRow+1).toString();
                  let cell = sheet.getRange(cellA1Notation);
                  cell.setValue(rowOfEvent[j]);   
                }
              }
            }
          }

        }
      }

      protectIDColumn_(sheet, headerRow, infoCol);

    }
    
    // show a dialog box to indicate sync has completed.
    SpreadsheetApp.getUi().alert(`Sync Complete.`);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(error);
  }
  return;

  function titlesIsEmpty(rowOfEvent, infoCol) {

    for (const titleCol of infoCol.titles) {
      if (rowOfEvent[titleCol]) {
        return false;
      }
    }
    return true;

  }

}

/**
 * Function that delete all the calendar event that is created by this spreadsheet.
 * Basically this is a hard refrest type of thing. If anything goes wrong, clear all, then resycn it with the canlendar.
 */
function clearAll() {

  // Show prompt to user indicating that this will clear everything, and ask them to confirm on continue.
  let ui = SpreadsheetApp.getUi();
  let response = ui.alert('WARNING! IRREVERSIBLE ACTION! \nThis will delete all the auto generated event and clear the EventID columns content in all defined sheets. \nAre you sure you want to continue?', ui.ButtonSet.YES_NO);

  if (response === ui.Button.NO) {
    return;
  }

  try {

    // Get the calendar
    let calendar = CalendarApp.getCalendarById(calendarID);
    if (calendar === null) {
      throw new Error(`Failed to get your calendar: ${calendarID}.`);
    }

    // Get all the event and delete them.
    let events = getCalenderEvents_(calendar);
    for (const event of events) {
      if (event.recurrence === "null") {
        let calEvent = calendar.getEventById(event.id);
        calEvent.deleteEvent();
      }
      else {
        let calEvent = calendar.getEventById(event.id);
        let calEventSeries = calEvent.getEventSeries();
        calEventSeries.deleteEventSeries();
      }
    }

    for (let s = 0; s < sheetNames.length; s++) {

      // Get the sheet
      let sheet = SpreadsheetApp.openByUrl(spreadSheetURL).getSheetByName(sheetNames[s]);
      if (sheet === null) {
        throw new Error(`Unable to find ${sheetName} in ${spreadSheetURL}.`);
      }

      // Erase all everything in the CalendarEventID column(s).
      let headerRow = getHeaderRow_(sheet); // Find the header.
      let infoCol = getInfoColumn_(sheet, s, headerRow); // Find the CalendarEventID column.
      let dataA1Notation = lettersFromIndex_(infoCol.ids[0]) + (headerRow+1).toString() + ":" + lettersFromIndex_(infoCol.ids[infoCol.ids.length-1]);
      sheet.getRange(dataA1Notation).clearContent();

    }

    // show a dialog box to indicate clear has completed.
    ui.alert(`Clear Complete.`);

  } catch (error) {
    ui.alert(error);
  }
  return;
}

/**
 * Helper function that capture all the event in spreadsheet.
 */
function getSheetEvents_(sheet, sheetIdx) {
  
  // Create an events array to store all the event.
  let events = [];

  // Find the header row, which is defined by a specific background color.
  let headerRow = getHeaderRow_(sheet);

  // Find the column that contain the info of the calendar events.
  let infoCol = getInfoColumn_(sheet, sheetIdx, headerRow);
  
  // Get all the relevant data in the spreadsheet.
  let dataA1Notation = (headerRow+1).toString() + ":" + sheet.getLastRow().toString();
  let data = sheet.getRange(dataA1Notation).getValues();
  
  // Loop through the all the data row by row.
  // If there is a valid event, make an object to push into events array.
  for (let i = 0; i < data.length; i++) {

    // Skip rows that does not have date, time and any title info.
    if (data[i][infoCol.date] === "" || data[i][infoCol.time] === "" || !containValidTitlesData(data[i], infoCol.titles)) {
      continue;
    }

    let formedEvents = sheetRowToEvent_(data[i], infoCol, i, sheetIdx);
    events = events.concat(formedEvents);

  }

  return events;

  /**
   * Function to check whether or not the row contain any titles data, and check whether the data is all marked with noSyncString.
   * @param {[]} data1D row of the data
   * @param {[]} titlesCol an array which represent the column index of each title data.
   * @returns {boolean} Return true if there is at least contain one title data. Otherwise false.
   */
  function containValidTitlesData(data1D, titlesCol) {
    let noSyncRegEx = new RegExp(noSyncString + '(\\w|\\d|\\s)*');
    let count = 0;
    for (const titleCol of titlesCol) {
      if (data1D[titleCol] !== "" && !noSyncRegEx.test(data1D[titleCol])) {
        count++;
      }
    }
    if (count === 0) {
      return false;
    }
    else {
      return true;
    }
  }
  
}

/**
 * Helper function that capture all the event in calander.
 */
function getCalenderEvents_(calendar) {

  // Create an events array to store all the event.
  let events = [];

  // Get all the events that occur within defined sync period.
  let syncStartDate = new Date();
  syncStartDate.setFullYear(syncStartDate.getFullYear() - syncPeriod.yearBefore);
  let syncEndDate = new Date();
  syncEndDate.setFullYear(syncEndDate.getFullYear() + syncPeriod.yearAfter);
  let calEvents = calendar.getEvents(syncStartDate, syncEndDate);

  // Loop through all the calendar events. Extract the event that is created by this script and place it into events array.
  for (const e of calEvents) {

    // SKip the events that does not contain `Created by "${identificationString}" spreadsheet.` tag.
    let tagKeys = e.getAllTagKeys();
    if (!tagKeys.includes('identification')) {
      continue;
    }
    else {
      let tagStr = `Created by "${identificationString}" spreadsheet.`;
      if (e.getTag('identification') !== tagStr) {
        continue;
      }
    }

    // Check if the event is part of a recurring event.
    if (e.isRecurringEvent()) {
      let repeated = false;
      for (let i = 0; i < events.length; i++) {
        let title = e.getTitle();
        if (events[i].title === title) {
          repeated = true;
          break;
        }
      }
      if (repeated) {
        continue;
      }
    }

    events.push(calanderToEvent_(e));
    
  }

  return events;

}

/**
 * Helper function that compare the events array.
 */
function compareEvents_(sheetEvents, calendarEvents) {

  // Create an diffEvents array to store all different event.
  let diffEvents = [];

  // Generate an array of IDs of events that is unique in its kind.
  let IDsArr = [];
  for (const sheetEvent of sheetEvents) {
    IDsArr.push(sheetEvent.id);
  }
  for (const calendarEvent of calendarEvents) {
    IDsArr.push(calendarEvent.id);
  }
  let uniqueID = Array.from(new Set(IDsArr));

  // Loop through the uniqueID array to check does the event:
  // - only exist in sheetEvents 
  // - only exist in calendarEvents
  // - exsit in both sheetEvents and calendar Events.
  for (const id of uniqueID) {

    // Find where does the event located in sheetEvents array and calendarEvents array.
    let sheetEventsIdx;
    let calendarEventsIdx;
    for (let i = 0; i < sheetEvents.length; i++) {
      if (id === sheetEvents[i].id) {
        sheetEventsIdx = i;
        break;
      }
    }
    for (let i = 0; i < calendarEvents.length; i++) {
      if (id === calendarEvents[i].id) {
        calendarEventsIdx = i;
        break;
      }
    }

    if (typeof(sheetEventsIdx) !== "undefined" && typeof(calendarEventsIdx) !== "undefined") {
      if (sheetEvents[sheetEventsIdx].recurrence === "null" && calendarEvents[calendarEventsIdx].recurrence !== "null") {
        diffEvents.push({
          "belong": "sheet&DR",
          "id": sheetEvents[sheetEventsIdx].id,
          "title": sheetEvents[sheetEventsIdx].title,
          "description": sheetEvents[sheetEventsIdx].description,
          "startTime": sheetEvents[sheetEventsIdx].startTime,
          "endTime": sheetEvents[sheetEventsIdx].endTime,
          "recurrence": sheetEvents[sheetEventsIdx].recurrence
        });
        diffEvents.push({
          "belong": "calendar&AR",
          "id": calendarEvents[calendarEventsIdx].id,
          "title": calendarEvents[calendarEventsIdx].title,
          "description": calendarEvents[calendarEventsIdx].description,
          "startTime": calendarEvents[calendarEventsIdx].startTime,
          "endTime": calendarEvents[calendarEventsIdx].endTime,
          "recurrence": calendarEvents[calendarEventsIdx].recurrence
        });
      }
      else if (sheetEvents[sheetEventsIdx].recurrence !== "null" && calendarEvents[calendarEventsIdx].recurrence === "null") {
        diffEvents.push({
          "belong": "sheet&AR",
          "id": sheetEvents[sheetEventsIdx].id,
          "title": sheetEvents[sheetEventsIdx].title,
          "description": sheetEvents[sheetEventsIdx].description,
          "startTime": sheetEvents[sheetEventsIdx].startTime,
          "endTime": sheetEvents[sheetEventsIdx].endTime,
          "recurrence": sheetEvents[sheetEventsIdx].recurrence
        });
        diffEvents.push({
          "belong": "calendar&DR",
          "id": calendarEvents[calendarEventsIdx].id,
          "title": calendarEvents[calendarEventsIdx].title,
          "description": calendarEvents[calendarEventsIdx].description,
          "startTime": calendarEvents[calendarEventsIdx].startTime,
          "endTime": calendarEvents[calendarEventsIdx].endTime,
          "recurrence": calendarEvents[calendarEventsIdx].recurrence
        });
      }
      else if (
        (
          (
            sheetEvents[sheetEventsIdx].recurrence === "null" && calendarEvents[calendarEventsIdx].recurrence === "null"
          ) &&
          (
            sheetEvents[sheetEventsIdx].title !== calendarEvents[calendarEventsIdx].title ||
            sheetEvents[sheetEventsIdx].startTime.getTime() !== calendarEvents[calendarEventsIdx].startTime.getTime() ||
            sheetEvents[sheetEventsIdx].endTime.getTime() !== calendarEvents[calendarEventsIdx].endTime.getTime()
          )
        ) || 
        (
          (
            sheetEvents[sheetEventsIdx].recurrence !== "null" && calendarEvents[calendarEventsIdx].recurrence !== "null"
          ) &&
          (
            sheetEvents[sheetEventsIdx].title !== calendarEvents[calendarEventsIdx].title ||
            sheetEvents[sheetEventsIdx].startTime.getTime() !== calendarEvents[calendarEventsIdx].startTime.getTime() ||
            sheetEvents[sheetEventsIdx].endTime.getTime() !== calendarEvents[calendarEventsIdx].endTime.getTime() ||
            sheetEvents[sheetEventsIdx].recurrence.rule !== calendarEvents[calendarEventsIdx].recurrence.rule ||
            sheetEvents[sheetEventsIdx].recurrence.repeatTimes !== calendarEvents[calendarEventsIdx].recurrence.repeatTimes ||
            sheetEvents[sheetEventsIdx].recurrence.end !== calendarEvents[calendarEventsIdx].recurrence.end ||
            sheetEvents[sheetEventsIdx].recurrence.endTimes !== calendarEvents[calendarEventsIdx].recurrence.endTimes ||
            sheetEvents[sheetEventsIdx].recurrence.endDate.year !== calendarEvents[calendarEventsIdx].recurrence.endDate.year ||
            sheetEvents[sheetEventsIdx].recurrence.endDate.month !== calendarEvents[calendarEventsIdx].recurrence.endDate.month ||
            sheetEvents[sheetEventsIdx].recurrence.endDate.day !== calendarEvents[calendarEventsIdx].recurrence.endDate.day ||
            sheetEvents[sheetEventsIdx].recurrence.repeatMode !== calendarEvents[calendarEventsIdx].recurrence.repeatMode ||
            !arrayIsEqual_(sheetEvents[sheetEventsIdx].recurrence.repeatOn, calendarEvents[calendarEventsIdx].recurrence.repeatOn)
          )
        )
      ) {
        diffEvents.push({
          "belong": "sheet&",
          "id": sheetEvents[sheetEventsIdx].id,
          "title": sheetEvents[sheetEventsIdx].title,
          "description": sheetEvents[sheetEventsIdx].description,
          "startTime": sheetEvents[sheetEventsIdx].startTime,
          "endTime": sheetEvents[sheetEventsIdx].endTime,
          "recurrence": sheetEvents[sheetEventsIdx].recurrence
        });
        diffEvents.push({
          "belong": "calendar&",
          "id": calendarEvents[calendarEventsIdx].id,
          "title": calendarEvents[calendarEventsIdx].title,
          "description": calendarEvents[calendarEventsIdx].description,
          "startTime": calendarEvents[calendarEventsIdx].startTime,
          "endTime": calendarEvents[calendarEventsIdx].endTime,
          "recurrence": calendarEvents[calendarEventsIdx].recurrence
        });
      }
    }
    else if (typeof(sheetEventsIdx) !== "undefined") {
      diffEvents.push({
        "belong": "sheet",
        "id": sheetEvents[sheetEventsIdx].id,
        "title": sheetEvents[sheetEventsIdx].title,
        "description": sheetEvents[sheetEventsIdx].description,
        "startTime": sheetEvents[sheetEventsIdx].startTime,
        "endTime": sheetEvents[sheetEventsIdx].endTime,
        "recurrence": sheetEvents[sheetEventsIdx].recurrence
      });
    }
    else {
      diffEvents.push({
        "belong": "calendar",
        "id": calendarEvents[calendarEventsIdx].id,
        "title": calendarEvents[calendarEventsIdx].title,
        "description": calendarEvents[calendarEventsIdx].description,
        "startTime": calendarEvents[calendarEventsIdx].startTime,
        "endTime": calendarEvents[calendarEventsIdx].endTime,
        "recurrence": calendarEvents[calendarEventsIdx].recurrence
      });
    }
    
  }

  return diffEvents;

}

/**
 * Helper function that take a whole row of into the event obejct that is used throughout this script.
 */
var selfGenID = 0;
function sheetRowToEvent_(rowOfData, infoCol, rowNumber, sheetIdx) {

  let formedEvents = [];

  let dateCol = infoCol.date;
  let timeCol = infoCol.time;
  let recurCol = infoCol.recurrence;
  let titlesCol = infoCol.titles;
  let idsCol = infoCol.ids;

  let date = typeof(rowOfData[dateCol]) === "number" ? rowOfData[dateCol].toString() : rowOfData[dateCol];
  let time = typeof(rowOfData[timeCol]) === "number" ? rowOfData[timeCol].toString() : rowOfData[timeCol];
  let recur = rowOfData[recurCol];

  // Verify the date, time and recurrence input, and then
  // format the date, time and recurrence into Javascript Date object, for setting the startTime and endTime of event.
  let dateRegEx = /\d{8}/;
  let timeRegEx = /\d{4}/;
  let recurRegEx1 = /Re:every\d+(Day|Week|Month|Year)/i;
  let recurRegEx2_1 = /On:\[\D*?\]/i;
  let recurRegEx2_2 = /Mon|Tues|Wed|Thurs|Fri|Sat|Sun/gi;
  let recurRegEx3 = /With:(Date|Week)/i;
  let recurRegEx4 = /End(On:\d{8}|After:\d+times)/i;
  let recurrence = "null";
  let start = {};
  let end = {};
  let dateArr = date.split("-");
  let timeArr = time.split("-");
  let recurArr = recur.split(" ");

  // Process date.
  if (dateArr.length > 2) {
    throw new Error(`Error: Invalid date format.\nMore than 1 hyphen detected in date coloum of row ${rowNumber+1}, counted from header.`);
  }
  for (const d of dateArr) {
    if (!dateRegEx.test(d)) {
      throw new Error(`Error: Invalid date format.\nDate should be in YYYYMMDD format. Eg: 20200401. Error found in date column of row ${rowNumber+1}, counted from header.`);
    }
  }
  start.year = parseInt(dateArr[0].slice(0,4));
  start.month = parseInt(dateArr[0].slice(4,6));
  start.day = parseInt(dateArr[0].slice(6));
  if (dateArr.length === 1) {
    end.year = start.year;
    end.month = start.month;
    end.day = start.day;
  }
  else {
    end.year = parseInt(dateArr[1].slice(0,4));
    end.month = parseInt(dateArr[1].slice(4,6));
    end.day = parseInt(dateArr[1].slice(6));
  }

  // Process time.
  if (timeArr.length > 2) {
    throw new Error(`Error: Invalid time format.\nMore than 1 hyphen detected in time coloum of row ${rowNumber+1}, counted from header.`);
  }
  for (const t of timeArr) {
    if (!timeRegEx.test(t)) {
      throw new Error(`Error: Invalid time format.\nTime should be in 24 Hour format. Eg: 0800. Error found in date column of row ${rowNumber+1}, counted from header.`);
    }
  }
  start.hours = parseInt(timeArr[0].slice(0,2));
  start.minutes = parseInt(timeArr[0].slice(2));
  if (timeArr.length === 1) {
    end.hours = start.hours + 1;
    end.minutes = start.minutes;
  }
  else {
    end.hours = parseInt(timeArr[1].slice(0,2));
    end.minutes = parseInt(timeArr[1].slice(2));
  }

  let startTime = new Date(start.year, start.month-1, start.day, start.hours, start.minutes);
  let endTime = new Date(end.year, end.month-1, end.day, end.hours, end.minutes);

  // Skip this event(s) if it is outside of the sync period.
  let syncStartDate = new Date();
  syncStartDate.setFullYear(syncStartDate.getFullYear() - syncPeriod.yearBefore);
  let syncEndDate = new Date();
  syncEndDate.setFullYear(syncEndDate.getFullYear() + syncPeriod.yearAfter);
  if (endTime.getTime() < syncStartDate.getTime() || startTime.getTime() > syncEndDate.getTime()) {
    return [];
  } 

  // Process recurrence.
  if (recur !== "") {
    recurrence = {
      'rule': "null",
      'repeatTimes': "null",
      'end': "null",
      'endTimes': "null",
      'endDate': {
        'year': "null",
        'month': "null",
        'day': "null"
      },
      'repeatMode': "null",
      'repeatOn': ["null"]
    };
    if (recurArr.length > 3) {
      throw new Error(`Error: Invalid recurrence format.\nMore than 3 section detected in recurrence coloum of row ${rowNumber+1}, counted from header.`);
    }
    if (!recurRegEx1.test(recurArr[0])) {
      throw new Error(`Error: Invalid recurrence format.\nPlease follow the predefined format. Error found in recurrence column of row ${rowNumber+1}, counted from header.`);
    }
    let recurRuleIdx = recurArr[0].search(/Day|Week|Month|Year/i);
    recurrence.rule = recurArr[0].slice(recurRuleIdx).toLowerCase();
    recurrence.repeatTimes = parseInt(recurArr[0].slice(8, recurRuleIdx));
    if (recurrence.rule === "week") {
      if (!recurRegEx2_1.test(recurArr[1])) {
        throw new Error(`Error: Invalid recurrence format.\nPlease follow the predefined format for recurrence week. Error found in recurrence column of row ${rowNumber+1}, counted from header.`);
      }
      recurrence.repeatOn = recurArr[1].match(recurRegEx2_2);
      if (recurrence.repeatOn === null) {
        throw new Error(`Error: Invalid recurrence format.\nUnable to identify which day(s) of the week to repeat. Error found in recurrence column of row ${rowNumber+1}, counted from header.`);
      }
      if (recurArr.length === 3) {
        if (!recurRegEx4.test(recurArr[2])) {
          throw new Error(`Error: Invalid recurrence format.\nPlease follow the predefined format for recurrence week. Error found in recurrence column of row ${rowNumber+1}, counted from header.`);
        }
        let tempArr = recurArr[2].split(':');
        recurrence.end = tempArr[0];
        if (tempArr[0].toLowerCase() === "endafter") {
          recurrence.endTimes = parseInt(tempArr[1].slice(0,-5));
        }
        else {
          recurrence.endDate.year = parseInt(tempArr[1].slice(0,4));
          recurrence.endDate.month = parseInt(tempArr[1].slice(4,6));
          recurrence.endDate.day = parseInt(tempArr[1].slice(6));
          let endOn = new Date(recurrence.endDate.year, recurrence.endDate.month-1, recurrence.endDate.day);
          if (endTime.getTime() >= endOn.getTime()) {
            throw new Error(`Error: Invalid recurrence value. \nThe recurrence endOn date must be greater than the event end date.`);
          }
        }
      }
    }
    else if (recurrence.rule === "month") {
      if (!recurRegEx3.test(recurArr[1])) {
        throw new Error(`Error: Invalid recurrence format.\nPlease follow the predefined format for recurrence month. Error found in recurrence column of row ${rowNumber+1}, counted from header.`);
      }
      recurrence.repeatMode = recurArr[1].slice(5);
      if (recurArr.length === 3) {
        if (!recurRegEx4.test(recurArr[2])) {
          throw new Error(`Error: Invalid recurrence format.\nPlease follow the predefined format for recurrence month. Error found in recurrence column of row ${rowNumber+1}, counted from header.`);
        }
        let tempArr = recurArr[2].split(':');
        recurrence.end = tempArr[0];
        if (tempArr[0].toLowerCase() === "endafter") {
          recurrence.endTimes = parseInt(tempArr[1].slice(0,-5));
        }
        else {
          recurrence.endDate.year = parseInt(tempArr[1].slice(0,4));
          recurrence.endDate.month = parseInt(tempArr[1].slice(4,6));
          recurrence.endDate.day = parseInt(tempArr[1].slice(6));
          let endOn = new Date(recurrence.endDate.year, recurrence.endDate.month-1, recurrence.endDate.day);
          if (endTime.getTime() >= endOn.getTime()) {
            throw new Error(`Error: Invalid recurrence value. \nThe recurrence endOn date must be greater than the event end date.`);
          }
        }
      }
    }
    else {
      if (recurArr.length === 3) {
        throw new Error(`Error: Invalid recurrence format.\nPlease follow the predefined format for recurrence day/year. Error found in recurrence column of row ${rowNumber+1}, counted from header.`);
      }
      else if (recurArr.length === 2) {
        if (!recurRegEx4.test(recurArr[1])) {
          throw new Error(`Error: Invalid recurrence format.\nPlease follow the predefined format for recurrence day/year. Error found in recurrence column of row ${rowNumber+1}, counted from header.`);
        }
        let tempArr = recurArr[1].split(':');
        recurrence.end = tempArr[0];
        if (tempArr[0].toLowerCase() === "endafter") {
          recurrence.endTimes = parseInt(tempArr[1].slice(0,-5));
        }
        else {
          recurrence.endDate.year = parseInt(tempArr[1].slice(0,4));
          recurrence.endDate.month = parseInt(tempArr[1].slice(4,6));
          recurrence.endDate.day = parseInt(tempArr[1].slice(6));
          let endOn = new Date(recurrence.endDate.year, recurrence.endDate.month-1, recurrence.endDate.day);
          if (endTime.getTime() >= endOn.getTime()) {
            throw new Error(`Error: Invalid recurrence value. \nThe recurrence endOn date must be greater than the event end date.`);
          }
        }
      }
    }
  }

  // Prepare the title and description of the event, and then push all the event into into the formedEvents array.
  let eventDescription = `Created by "${identificationString}" spreadsheet.`
  for (let i = 0; i < titlesCol.length; i++) {

    // Skip create event if the title column does not have any value, or it has noSyncString at the begining of the title.
    let noSyncRegEx = new RegExp(noSyncString + '(\\w|\\d|\\s)*');
    if (rowOfData[titlesCol[i]] === "" || noSyncRegEx.test(rowOfData[titlesCol[i]])) {
      continue;
    }

    let eventTitle = headerString[sheetIdx].prefix[i] + " - " + rowOfData[titlesCol[i]].slice(0,1000);
    let eventID;
    if (rowOfData[idsCol[i]] !== "") { 
      eventID = rowOfData[idsCol[i]]; 
    }
    else {
      eventID = selfGenID;
      selfGenID++;
    }

    formedEvents.push({
      "id": eventID,
      "title": eventTitle,
      "description": eventDescription,
      "startTime": startTime,
      "endTime": endTime,
      "recurrence": recurrence
    });

  }

  return formedEvents;

}

/**
 * Helper function that take a convert calendar event into event object that is used throughout this script.
 */
function calanderToEvent_(calendarEvent) {
  let id = calendarEvent.getId();
  let title = calendarEvent.getTitle();
  let description = calendarEvent.getDescription();
  let startTime = calendarEvent.getStartTime();
  let endTime = calendarEvent.getEndTime();
  let event = {
    'id': id,
    'title': title,
    'description': description,
    'startTime': startTime,
    'endTime': endTime,
    'recurrence': "null"
  };
  if (calendarEvent.isRecurringEvent()) {
    let recurrence = calendarEvent.getTag('recurrence');
    event.recurrence = JSON.parse(recurrence);
  }
  return event;
}

/**
 * Helper function that generate recurrence rule for calendar.
 */
function formRecurrenceRule_(event) {
  let recurRule;
  if (event.recurrence.rule === "day") {
    if (event.recurrence.end.toLowerCase() === "endafter") {
      recurRule = CalendarApp.newRecurrence().addDailyRule().interval(event.recurrence.repeatTimes).times(event.recurrence.endTimes);
    }
    else if (event.recurrence.end.toLowerCase() === "endon") {
      recurRule = CalendarApp.newRecurrence().addDailyRule().interval(event.recurrence.repeatTimes).until(new Date(event.recurrence.endDate.year, event.recurrence.endDate.month-1, event.recurrence.endDate.day));
    }
    else {
      recurRule = CalendarApp.newRecurrence().addDailyRule().interval(event.recurrence.repeatTimes);
    }
  }
  else if (event.recurrence.rule === "week") {
    let repeatOn = [];
    for (const week of event.recurrence.repeatOn) {
      if (week.toLowerCase() === "mon") { repeatOn.push(CalendarApp.Weekday.MONDAY); }
      else if (week.toLowerCase() === "tues") { repeatOn.push(CalendarApp.Weekday.TUESDAY); }
      else if (week.toLowerCase() === "wed") { repeatOn.push(CalendarApp.Weekday.WEDNESDAY); }
      else if (week.toLowerCase() === "thurs") { repeatOn.push(CalendarApp.Weekday.THURSDAY); }
      else if (week.toLowerCase() === "fri") { repeatOn.push(CalendarApp.Weekday.FRIDAY); }
      else if (week.toLowerCase() === "sat") { repeatOn.push(CalendarApp.Weekday.SATURDAY); }
      else { repeatOn.push(CalendarApp.Weekday.SUNDAY); }
    }
    if (event.recurrence.end.toLowerCase() === "endafter") {
      recurRule = CalendarApp.newRecurrence().addWeeklyRule().interval(event.recurrence.repeatTimes).onlyOnWeekdays(repeatOn).times(event.recurrence.endTimes);
    }
    else if (event.recurrence.end.toLowerCase() === "endon") {
      recurRule = CalendarApp.newRecurrence().addWeeklyRule().interval(event.recurrence.repeatTimes).onlyOnWeekdays(repeatOn).until(new Date(event.recurrence.endDate.year, event.recurrence.endDate.month-1, event.recurrence.endDate.day));
    }
    else {
      recurRule = CalendarApp.newRecurrence().addWeeklyRule().interval(event.recurrence.repeatTimes).onlyOnWeekdays(repeatOn);
    }
  }
  else if (event.recurrence.rule === "month") {
    if (event.recurrence.repeatMode.toLowerCase() === "date") {
      let dayOfMonth = event.startTime.getDate();
      if (event.recurrence.end.toLowerCase() === "endafter") {
        recurRule = CalendarApp.newRecurrence().addMonthlyRule().interval(event.recurrence.repeatTimes).onlyOnMonthDay(dayOfMonth).times(event.recurrence.endTimes);
      }
      else if (event.recurrence.end.toLowerCase() === "endon") {
        recurRule = CalendarApp.newRecurrence().addMonthlyRule().interval(event.recurrence.repeatTimes).onlyOnMonthDay(dayOfMonth).until(new Date(event.recurrence.endDate.year, event.recurrence.endDate.month-1, event.recurrence.endDate.day));
      }
      else {
        recurRule = CalendarApp.newRecurrence().addMonthlyRule().interval(event.recurrence.repeatTimes).onlyOnMonthDay(dayOfMonth);
      }
    }
    else {
      let dayOfWeek;
      switch (event.startTime.getDay()) {
        case 0: { dayOfWeek = CalendarApp.Weekday.SUNDAY; break; }
        case 1: { dayOfWeek = CalendarApp.Weekday.MONDAY; break; }
        case 2: { dayOfWeek = CalendarApp.Weekday.TUESDAY; break; }
        case 3: { dayOfWeek = CalendarApp.Weekday.WEDNESDAY; break; }
        case 4: { dayOfWeek = CalendarApp.Weekday.THURSDAY; break; }
        case 5: { dayOfWeek = CalendarApp.Weekday.FRIDAY; break; }
        case 5: { dayOfWeek = CalendarApp.Weekday.SATURDAY; break; }
        default:
          break;
      }
      let numbers = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31];
      let dayOfMonth = event.startTime.getDate();
      let mod = dayOfMonth % 7 - 1;
      let good_days = numbers.slice(dayOfMonth - mod, dayOfMonth - mod + 7);
      if (event.recurrence.end.toLowerCase() === "endafter") {
        recurRule = CalendarApp.newRecurrence().addWeeklyRule().interval(event.recurrence.repeatTimes).onlyOnWeekday(dayOfWeek).onlyOnMonthDays(good_days).times(event.recurrence.endTimes);
      }
      else if (event.recurrence.end.toLowerCase() === "endon") {
        recurRule = CalendarApp.newRecurrence().addWeeklyRule().interval(event.recurrence.repeatTimes).onlyOnWeekday(dayOfWeek).onlyOnMonthDays(good_days).until(new Date(event.recurrence.endDate.year, event.recurrence.endDate.month-1, event.recurrence.endDate.day));
      }
      else {
        recurRule = CalendarApp.newRecurrence().addWeeklyRule().interval(event.recurrence.repeatTimes).onlyOnWeekday(dayOfWeek).onlyOnMonthDays(good_days);
      }
    } 
  }
  else {
    if (event.recurrence.end.toLowerCase() === "endafter") {
      recurRule = CalendarApp.newRecurrence().addYearlyRule().interval(event.recurrence.repeatTimes).times(event.recurrence.endTimes);
    }
    else if (event.recurrence.end.toLowerCase() === "endon") {
      recurRule = CalendarApp.newRecurrence().addYearlyRule().interval(event.recurrence.repeatTimes).until(new Date(event.recurrence.endDate.year, event.recurrence.endDate.month-1, event.recurrence.endDate.day));
    }
    else {
      recurRule = CalendarApp.newRecurrence().addYearlyRule().interval(event.recurrence.repeatTimes);
    }
  }
  return recurRule;
}

/**
 * Helper function that find the header row in spreadsheet, which is defined by a specific background color.
 */
function getHeaderRow_(sheet) {
  let headerRow;
  for (let i = 1; i < sheet.getLastRow(); i++) {
    let r = sheet.getRange(i.toString() + ":" + i.toString());
    if (r.getBackground() === headerColor) {  
      headerRow = i;
      break;
    }
  }
  if (typeof(headerRow) === "undefined") {
    throw new Error(`Error: Invalid sheet format.\nUnable to find header row. Please use ${headerColor} as the background color of your header.`);
  }
  return headerRow;
}

/**
 * Helper function that find the column index that contain all the important infomation. 
 * Note that it does not return column index in A1 notation. Use lettersFromIndex(index) to convert it.
 * @returns {Object} Object that holds all the index. 
 */
function getInfoColumn_(sheet, sheetIdx, headerRow) {

  let dateCol;
  let timeCol;
  let recurCol;
  let titlesCol = [];
  let idsCol = [];
  let headerData = sheet.getRange(headerRow.toString() + ":" + headerRow.toString()).getValues();
  for (let i = 0; i < headerData[0].length; i++) {
    if (headerData[0][i] === headerString[sheetIdx].date) { dateCol = i; }
    else if (headerData[0][i] === headerString[sheetIdx].time) { timeCol = i; }
    else if (headerData[0][i] === headerString[sheetIdx].recurrence) { recurCol = i; }
    else {
      let contFlag = false;
      for (let j = 0; j < headerString[sheetIdx].titles.length; j++) {
        if (headerData[0][i] === headerString[sheetIdx].titles[j]) { 
          titlesCol[j] = i; 
          contFlag = true; 
          break; 
        }
      }
      if (contFlag) { continue; }
      for (let j = 0; j < headerString[sheetIdx].ids.length; j++) {
        if (headerData[0][i] === headerString[sheetIdx].ids[j]) { 
          idsCol[j] = i; 
          break; 
        }
      }
    }
  }
  titlesCol = fillArrayWithNull(titlesCol, headerString[sheetIdx].titles.length);
  idsCol = fillArrayWithNull(idsCol, headerString[sheetIdx].ids.length);
  if (
    typeof(dateCol) === "undefined" || 
    typeof(timeCol) === "undefined" || 
    typeof(recurCol) === "undefined" || 
    titlesCol.length !== headerString[sheetIdx].titles.length || titlesCol.includes("null") ||
    idsCol.length !== headerString[sheetIdx].ids.length || idsCol.includes("null")
  ) {
    let errMsg = "";
    if (typeof(dateCol) === "undefined") { errMsg = errMsg + `Unable to find ${headerString[sheetIdx].date} in header of ${sheetNames[sheetIdx]}.\n` }
    if (typeof(timeCol) === "undefined") { errMsg = errMsg + `Unable to find ${headerString[sheetIdx].time} in header of ${sheetNames[sheetIdx]}.\n` }
    if (typeof(recurCol) === "undefined") { errMsg = errMsg + `Unable to find ${headerString[sheetIdx].recurrence} in header of ${sheetNames[sheetIdx]}.\n` }
    if (titlesCol.includes("null")) {
      for (let i = 0; i < titlesCol.length; i++) {
        if (titlesCol[i] === "null") { errMsg = errMsg + `Unable to find ${headerString[sheetIdx].titles[i]} in header of ${sheetNames[sheetIdx]}.\n` }
      }
    }
    if (idsCol.includes("null")) {
      for (let i = 0; i < idsCol.length; i++) {
        if (idsCol[i] === "null") { errMsg = errMsg + `Unable to find ${headerString[sheetIdx].ids[i]} in header of ${sheetNames[sheetIdx]}.\n` }
      }
    }
    throw new Error(`Error: Invalid sheet format.\n${errMsg}`);
  }

  return {
    'date': dateCol,
    'time': timeCol,
    'recurrence': recurCol,
    'titles': titlesCol,
    'ids': idsCol
  }

  function fillArrayWithNull(arr, length) {
    if (length < arr.length) {
      throw new Error(`Invalid parameter in fillArrayWithEmptyStr function. "length" parameter must be greater than the length of "arr"`);
    }
    let newArr = [];
    for (let i = 0; i < length; i++) {
      if (arr[i]) { newArr[i] = arr[i]; }
      else { newArr[i] = "null"; }
    }
    return newArr;
  }

}

/**
 * Helper function that convert event object into an array that can be directly appended to the end of the spreadsheet.
 */
function eventToSheetRow_(event, infoCol, sheetIdx) {

  let sheetRow = [];

  // Find the index that we need to use to set the Title and EventID.
  let splitedTitle = event.title.split(" - ");
  if (splitedTitle.length > 2) {
    splitedTitle[1] = splitedTitle.slice(1).join(' - ')
  }
  let idx = headerString[sheetIdx].prefix.indexOf(splitedTitle[0]);

  // Convert the event object into spreadsheet cell value. 
  let title = splitedTitle[1];
  let id = event.id;
  let date = "";
  let time = "";
  let recurrence = "";

  // Make this start and end object to make everything easier to read.
  let start = {
    'year': event.startTime.getFullYear(),
    'month': event.startTime.getMonth() + 1,
    'day': event.startTime.getDate(),
    'hours': event.startTime.getHours(),
    'minutes': event.startTime.getMinutes()
  }
  let end = {
    'year': event.endTime.getFullYear(),
    'month': event.endTime.getMonth() + 1,
    'day': event.endTime.getDate(),
    'hours': event.endTime.getHours(),
    'minutes': event.endTime.getMinutes()
  }
  
  // Form the date value.
  date = start.year.toString().padStart(4,"0") + start.month.toString().padStart(2,"0") + start.day.toString().padStart(2,"0");
  if (
    start.year !== end.year ||
    start.month !== end.month ||
    start.day !== end.day
  ) {
    date = date + "-" + end.year.toString().padStart(4,"0") + end.month.toString().padStart(2,"0") + end.day.toString().padStart(2,"0");
  }

  // Form the time value.
  time = start.hours.toString().padStart(2,"0") + start.minutes.toString().padStart(2,"0");
  if (
    start.hours !== end.hours - 1 ||
    start.minutes !== end.minutes
  ) {
    time = time + "-" + end.hours.toString().padStart(2,"0") + end.minutes.toString().padStart(2,"0");
  }

  // Form the recurrence value, if required.
  if (event.recurrence !== "null") {
    recurrence = "Re:every" + event.recurrence.repeatTimes.toString() + capitalizeFirstLetter(event.recurrence.rule);
    if (event.recurrence.rule === "week") {
      recurrence = recurrence + " On:[" + event.recurrence.repeatOn.join(',') + "]";
    }
    else if (event.recurrence.rule === "month") {
      recurrence = recurrence + " With:" + event.recurrence.repeatMode;
    }
    if (event.recurrence.end.toLowerCase() === "endafter") {
      recurrence = recurrence + " " + event.recurrence.end + ":" + event.recurrence.endTimes + "times"
    }
    else {
      recurrence = recurrence + " " + event.recurrence.end + ":" + event.recurrence.endDate.year.toString().padStart(4,"0") + event.recurrence.endDate.month.toString().padStart(2,"0") + event.recurrence.endDate.day.toString().padStart(2,"0");
    }
  }

  // Put everything into the sheetRow array.
  sheetRow[infoCol.ids[idx]] = id;
  sheetRow[infoCol.titles[idx]] = title;
  sheetRow[infoCol.date] = date;
  sheetRow[infoCol.time] = time;
  sheetRow[infoCol.recurrence] = recurrence;

  // fill all the undefined section with empty string.
  for (let i = 0; i < sheetRow.length; i++) {
    if (typeof(sheetRow[i]) === "undefined") {
      sheetRow[i] = "";
    }
  }

  return sheetRow;

  /**
   * Helper function that capitalize the first letter for the whole string. 
   * @param {string} str String that want to be capitalized.
   * @returns Original string but with the first letter capitalized. Eg: "day" -> "Day", "day of week" -> "Day of week"
   */
  function capitalizeFirstLetter(str) {
    return str.charAt(0).toUpperCase() + str.substring(1);
  }

} 

/**
 * Helper function set the eventID column to be protected.
 */
function protectIDColumn_(sheet, headerRow, infoCol) {

  for (const idCol of infoCol.ids) {
    let A1Notation = lettersFromIndex_(idCol) + (headerRow+1).toString() + ":" + lettersFromIndex_(idCol);
    let idRange = sheet.getRange(A1Notation);
    idRange.protect().setWarningOnly(true);
  }

  return;

}

/**
 * Helper function that check whether two array are similar or not.
 * @param {[]} arr1 - First array.
 * @param {[]} arr2 - Second array.
 * @returns {bool} True if both of the array's element are completely similar. Otherwise false.
 */
function arrayIsEqual_(arr1, arr2) {
  if (arr1.length !== arr2.length) { return false; }
  for (let i = 0; i < arr1.length; i++) {
    if (arr1[i] !== arr2[i]) {
      return false;
    }
  }
  return true;
}

/**
 * Helper function that convert index to column of A1 notation.
 * @param {number} index - Index that need to be converted.
 * @param {string} curResult - (Optional) Recursive result.
 * @param {number} i - (Optional) Times to perform recursive.
 * @returns {string} The A1 notation of the given index. Eg: 0 -> A, 24 -> Y, 31 -> AF, 9007199254740990 -> BKTXHSOGHKKE
 */
function lettersFromIndex_(index, curResult, i) {

  if (i == undefined) i = 11; //enough for Number.MAX_SAFE_INTEGER
  if (curResult == undefined) curResult = "";

  let factor = Math.floor(index / Math.pow(26, i)); //for the order of magnitude 26^i

  if (factor > 0 && i > 0) {
    curResult += String.fromCharCode(64 + factor);
    curResult = lettersFromIndex_(index - Math.pow(26, i) * factor, curResult, i - 1);

  } else if (factor == 0 && i > 0) {
    curResult = lettersFromIndex_(index, curResult, i - 1);

  } else {
    curResult += String.fromCharCode(65 + index % 26);

  }
  return curResult;
}
