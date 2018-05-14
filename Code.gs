function onInstall(e) {
  onOpen(e);

  // Perform additional setup as needed.
}

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createAddonMenu();
  var props = readConfig();

  if (e && e.authMode == ScriptApp.AuthMode.NONE || !props.isInitialized) {
    menu.addItem('Create New', 'showCreateNewSidebar');
    menu.addItem('Create With ...', 'showCreateFromExistingSidebar');
  } else {
    menu.addItem('Refresh', 'refresh');
    menu.addItem('Purge', 'purge');
    menu.addSeparator();
    menu.addItem('Options', 'showEditSidebar');
  }
  menu.addToUi();
}

/**
 * Opens a sidebar. The sidebar structure is described in the CreateNewSidebar.html
 * project file.
 */
function showCreateNewSidebar() {
  var ui = HtmlService.createTemplateFromFile('CreateNewSidebar')
    .evaluate()
    .setTitle('Create New');
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Create new roster fill up sheet with new look up data
 *
 * @param {String} sheetname - required.
 * @param {String} frequency. Frequency of the inverval - required.
 * @param {String} startDate. Start date of the interval 
 * @param {String} endDate. End date of the interval
 * @param {Array} daysInWeek. The days in a week for weekly frequency
 * @param {String} customSheetname. The sheet name of custom range.
 * @param {String} customRange. The A1 notion of the data rows.
 */
function createNew(sheetname, frequency, startDate, endDate, daysInWeek, customSheetname, customRange) {

  if (!sheetname) {
    throw 'Invalid sheet name, ' + sheetname;
  }
  var newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname);
  if (newSheet != null) {
    throw 'A sheet with name ' + sheetname + ' existed. Please use another name.';
  }

  try {
    switch (frequency) {
      case 'd':
        {
          // validate number of days
          var validatedDates = validateStartEndDates(startDate, endDate);
          startDate = validatedDates.startDate;
          endDate = validatedDates.endDate;

          var daysCount = getDaysBetween(startDate, endDate);
          if (daysCount <= 0) {
            // minimum 1 day
            daysCount = 1;
          }

          newSheet = SpreadsheetApp.getActive().insertSheet(sheetname);

          // configure name column
          var personNameRange = newSheet.getRange(1, 1);
          personNameRange.setValue('Name');

          // configure timeslot headers
          var timeslotRange = newSheet.getRange(1, 2, 1, daysCount);
          updateTimetableHeaders(newSheet.getSheetName(), timeslotRange.getA1Notation(), getDatesForDaily(daysCount));

          // configure submitted on column
          var submittedOnRange = newSheet.getRange(1, daysCount + 1, 1, 1);
          submittedOnRange.setValue('Submitted On');

          var newConfig = {
            sheet_name: newSheet.getName(),
            range: {
              person_name: personNameRange.getA1Notation(),
              timeslot: timeslotRange.getA1Notation(),
              timestamp: submittedOnRange.getA1Notation()
            },
            start_date: startDate,
            end_date: endDate,
            frequency: 'd'
          };

          saveConfig(newConfig);

          break;
        }
      case 'w':
        {
          // validate number of days
          var validatedDates = validateStartEndDates(startDate, endDate);
          startDate = validatedDates.startDate;
          endDate = validatedDates.endDate;

          // validate days in week
          var validDaysInWeek = filterDaysInWeek(daysInWeek);
          if (!validDaysInWeek || validDaysInWeek.length <= 0) {
            throw 'Invalid days in week';
          }

          var daysCount = getDaysBetweenForWeek(startDate, endDate, validDaysInWeek);

          newSheet = SpreadsheetApp.getActive().insertSheet(sheetname);

          // configure name column
          var personNameRange = newSheet.getRange(1, 1);
          personNameRange.setValue('Name');

          // configure timetable headers
          var timeslotRange = newSheet.getRange(1, 2, 1, daysCount);
          updateTimetableHeaders(newSheet.getSheetName(), timeslotRange.getA1Notation(), getDatesForWeekly(daysCount, validDaysInWeek));

          // configure submitted on column
          var submittedOnRange = newSheet.getRange(1, daysCount + 1, 1, 1);
          submittedOnRange.setValue('Submitted On');

          var newConfig = {
            sheet_name: newSheet.getName(),
            range: {
              person_name: personNameRange.getA1Notation(),
              timeslot: timeslotRange.getA1Notation(),
              timestamp: submittedOnRange.getA1Notation()
            },
            start_date: startDate,
            end_date: endDate,
            frequency: 'w',
            days_in_week: validDaysInWeek
          };

          saveConfig(newConfig);

          break;
        }
      case 'c':
        {
          // validate custom range
          if (!customSheetname || !customRange) {
            throw 'Please select the custom range with dates';
          }
          var range = getRangeFromSheetA1Notation(customSheetname, customRange);
          var isSingleRow = isSingleRowRange(range);
          var isSingleColumn = isSingleColumnRange(range);
          if (!isSingleRow && !isSingleColumn) {
            throw 'Provide custom dates in a single row or column only';
          }

          var validDates = getDatesFromCustomRange(range);
          if (!validDates || validDates.length <= 0) {
            throw 'Empty custom dates';
          }

          var daysCount = validDates.length;
          newSheet = SpreadsheetApp.getActive().insertSheet(sheetname);

          // configure name column
          var personNameRange = newSheet.getRange(1, 1);
          personNameRange.setValue('Name');

          // configure timetable headers
          var timeslotRange = newSheet.getRange(1, 2, 1, daysCount);
          updateTimetableHeaders(newSheet.getSheetName(), timeslotRange.getA1Notation(), validDates);

          // configure submitted on column
          var submittedOnRange = newSheet.getRange(1, daysCount + 1, 1, 1);
          submittedOnRange.setValue('Submitted On');

          var newConfig = {
            sheet_name: newSheet.getName(),
            range: {
              person_name: personNameRange.getA1Notation(),
              timeslot: timeslotRange.getA1Notation(),
              timestamp: submittedOnRange.getA1Notation()
            },
            start_date: startDate,
            end_date: endDate,
            frequency: 'c',
            custom_dates: {
              sheet_name: customSheetname,
              range: customRange
            }
          };

          saveConfig(newConfig);

          break;
        }
      default:
        {
          throw 'Unsupported frequency, ' + frequency;
        }
    }
  } catch (e) {
    Logger.log('Create New exception: %s', e);
    if (newSheet != null) {
      Logger.log('Create New roll back sheet creation');
      removeConfig(newSheet.getName());
      SpreadsheetApp.getActive().deleteSheet(newSheet);
    }
    throw e;
  }

  showCreateFromExistingSidebar();
}

function readConfigForActiveSheet() {
  var sheet = SpreadsheetApp.getActiveSheet();
  return readConfigForSheet(sheet.getName());
}

function readConfigForUserSheet() {
  // TODO: read sheetname from user property, else use active sheet
  return readConfigForSheet('');
}

function readConfigForSheet(sheetname) {
  if (!sheetname || typeof sheetname !== 'string') {
    throw 'Sheetname is required';
  }
  return readConfig(sheetname);
}

/**
 * Opens a sidebar. The sidebar structure is described in the CreateFromExistingSidebar.html
 * project file.
 */
function showCreateFromExistingSidebar() {
  var ui = HtmlService.createTemplateFromFile('CreateFromExistingSidebar')
    .evaluate()
    .setTitle('Create With ...');
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Create new roster fill up sheet from existing look up data
 */
function createFromExisting(sheet_name, range_person_name, range_timeslot, range_timestamp,
  lookup_sheet_name, lookup_range_person_name, lookup_range_timeslot, lookup_range_timestamp, lookup_data_retention_days) {

  // validate fillup sheet
  if (!sheet_name) {
    throw 'Sheetname for fillup is required';
  }
  var sheetFillup = SpreadsheetApp.getActive().getSheetByName(sheet_name);
  if (sheetFillup == null) {
    throw 'Invalid sheet [' + sheet_name + ']';
  }

  // validate fillup range
  var rangePersonName = getRangeFromSheetA1Notation(sheet_name, range_person_name);
  if (!isSingleColumnRange(rangePersonName)) {
    throw 'Range of person name (fillup) can only be a single column';
  }
  if (!rangePersonName.canEdit()) {
    throw 'Insufficent permission to update range of person name (fillup)';
  }
  var rangeTimeslot = getRangeFromSheetA1Notation(sheet_name, range_timeslot);
  if (!isSingleColumnRange(rangeTimeslot)) {
    throw 'Range of timeslot (fillup) can only be a single column';
  }
  if (!rangeTimeslot.canEdit()) {
    throw 'Insufficent permission to update range of timeslot (fillup)';
  }
  var rangeTimestamp = getRangeFromSheetA1Notation(sheet_name, range_timestamp);
  if (!isSingleColumnRange(rangeTimestamp)) {
    throw 'Range of timestamp (fillup) can only be a single column';
  }
  if (!rangeTimestamp.canEdit()) {
    throw 'Insufficent permission to update range of timestamp (fillup)';
  }

  // validate lookup sheet
  if (!lookup_sheet_name) {
    throw 'Sheetname for lookup is required';
  }
  var sheetLookup = SpreadsheetApp.getActive().getSheetByName(lookup_sheet_name);
  if (sheetLookup == null) {
    throw 'Invalid sheet [' + lookup_sheet_name + ']';
  }

  // validate lookup range
  var lookupRangePersonName = getRangeFromSheetA1Notation(lookup_sheet_name, lookup_range_person_name);
  if (!isSingleColumnRange(lookupRangePersonName)) {
    throw 'Range of person name (lookup) can only be a single column';
  }
  if (!lookupRangePersonName.canEdit()) {
    throw 'Insufficent permission to update range of person name (lookup)';
  }
  var lookupRangeTimeslot = getRangeFromSheetA1Notation(lookup_sheet_name, lookup_range_timeslot);
  if (!isSingleColumnRange(lookupRangeTimeslot)) {
    throw 'Range of timeslot (lookup) can only be a single column';
  }
  if (!lookupRangeTimeslot.canEdit()) {
    throw 'Insufficent permission to update range of timeslot (lookup)';
  }
  var lookupRangeTimestamp = getRangeFromSheetA1Notation(lookup_sheet_name, lookup_range_timestamp);
  if (!isSingleColumnRange(lookupRangeTimestamp)) {
    throw 'Range of timestamp (lookup) can only be a single column';
  }
  if (!lookupRangeTimestamp.canEdit()) {
    throw 'Insufficent permission to update range of timestamp (lookup)';
  }
}

/**
 * Opens a sidebar. The sidebar structure is described in the EditOptionsSidebar.html
 * project file.
 */
function showEditSidebar() {
  var ui = HtmlService.createTemplateFromFile('EditSidebar')
    .evaluate()
    .setTitle('Edit Roster');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function update() {}

function refresh() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.alert('You clicked the first menu item!');
}

function purge() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.alert('You clicked the first menu item!');
}

/**
 * Helper function to validate start date and end dates
 * return validated startDate & endDate in a object
 */
function validateStartEndDates(startDate, endDate) {
  if (!startDate) {
    throw 'Start date is required';
  }
  if (startDate.constructor !== Date) {
    startDate = new Date(startDate);
  }
  if (isNaN(startDate.getTime())) {
    throw 'Start date is invalid';
  }
  if (!endDate) {
    throw 'End date is required';
  }
  if (endDate.constructor !== Date) {
    endDate = new Date(endDate);
  }
  if (isNaN(endDate.getTime())) {
    throw 'End date is invalid';
  }
  if (startDate > endDate) {
    throw 'End date cannot be earlier than start date';
  }

  return {
    startDate: startDate,
    endDate: endDate
  };
}

/**
 * Helper function to populate dates for daily frequency
 */
function getDatesForDaily(daysDisplay, startDate) {
  if (!startDate || startDate.constructor !== Date) {
    startDate = new Date();
  }

  var dates = [];
  for (var i = 0; i < daysDisplay; i++) {
    dates.push(updateDate(startDate, 'd', i));
  }
  return dates;
}

/**
 * Helper function to populate dates for weekly frequency
 */
function getDatesForWeekly(daysDisplay, daysInWeek, startDate) {
  if (!startDate || startDate.constructor !== Date) {
    startDate = new Date();
  }

  // Calculate nearest future day w.r.t the given start date
  // e.g: starts from Monday
  var nextDay = 1;
  var nextDayIndex = 0;
  var nextDayMin = 7;
  var startDay = startDate.getDay();
  for (var i = 0; i < daysInWeek.length; i++) {
    // difference from start day to other days
    var otherDay = parseInt(daysInWeek[i]);
    var diff = otherDay - startDay;
    if (diff == 0) {
      // same day as current day
      nextDay = otherDay;
      nextDayIndex = i;
      nextDayMin = diff;
      Logger.log('Next day is today');
      Logger.log('startDay:' + startDay + ', otherDay:' + otherDay + ', diff:' + diff);
      break;
    } else if (diff > 0) {
      // future day within same week
      // if min is the same, always take future day
      Logger.log('Next day is located later in the current week');
      Logger.log('startDay:' + startDay + ', otherDay:' + otherDay + ', diff:' + diff);
      if (diff <= nextDayMin) {
        nextDay = otherDay;
        nextDayIndex = i;
        nextDayMin = diff;
      }
    } else {
      // future day out of same week
      var newDiff = diff + 7;
      Logger.log('Next day is located in the upcoming week');
      Logger.log('startDay:' + startDay + ', otherDay:' + otherDay + ', diff:' + newDiff);
      if (newDiff <= nextDayMin) {
        nextDay = otherDay;
        nextDayIndex = i;
        nextDayMin = newDiff;
      }
    }
  }
  Logger.log('Calculation of Next Day');
  Logger.log('nextDay:' + nextDay + ', nextDayIndex:' + nextDayIndex + ', nextDayMin:' + nextDayMin);

  // Calculate days between dates
  var daysBetween = [];
  var daysBetweenPrevious;
  var daysBetweenCurrent;
  if (daysInWeek.length <= 1) {
    // different in days to next week
    daysBetween.push(7);
  } else {
    for (var i = 0; i < daysInWeek.length; i++) {
      daysBetweenCurrent = parseInt(daysInWeek[i]);
      if (i > 0) {
        // difference in days withint same week
        daysBetween.push(daysBetweenCurrent - daysBetweenPrevious);
      }
      if (i == daysInWeek.length - 1) {
        // difference in days to next week
        daysBetween.push(parseInt(daysInWeek[0]) + 7 - daysBetweenCurrent);
      }
      daysBetweenPrevious = daysBetweenCurrent;
    }
  }

  Logger.log('Days Between Dates');
  Logger.log(JSON.stringify(daysBetween));

  var dates = [];
  var daysBetweenIndex = nextDayIndex;
  var dateBegin = updateDate(startDate, 'd', nextDayMin);
  for (var i = 0; i < daysDisplay; i++) {
    if (i > 0) {
      dateBegin = updateDate(dateBegin, 'd', daysBetween[daysBetweenIndex]);
      daysBetweenIndex++;
      daysBetweenIndex %= daysBetween.length;
    }
    dates.push(dateBegin);
  }
  return dates;
}

/**
 * Helper function to populate roster timeable headers
 */
function updateTimetableHeaders(sheetname, A1Notation, dates) {

  Logger.log('UpdateTimeTableHeaders');

  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetname);
  if (sheet == null) {
    throw 'No such sheet:[' + sheetname + ']';
  }

  var range = sheet.getRange(A1Notation);
  if (!isSingleRowRange(range)) {
    throw 'Range give for timable headers must only be a single row.';
  }

  if (!dates || dates.constructor !== Array || !dates.length) {
    Logger.log(JSON.stringify(dates));
    throw 'Insufficent dates given for timetable headers';
  }

  var numColumns = range.getNumColumns();
  if (numColumns !== dates.length) {
    Logger.log('Range: ' + A1Notation);
    Logger.log('Dates:' + JSON.stringify(dates));
    throw 'Dates do not match with the give range';
  }

  var datesInText = [];
  var datesFormat = [];
  for (var i = 0; i < numColumns; i++) {

    var dateFormat = 'ddd (d-mmm)';
    var dateValue = '';
    if (dates[i] && dates[i].constructor === Date) {
      dateValue = new Date(dates[i]);
    }

    datesFormat.push(dateFormat);
    datesInText.push(dateValue);
  }

  range.setNumberFormats([datesFormat]);
  range.setValues([datesInText]);
}

/*
 **************** HELPER FUNCTIONS SECTION ****************
 */

/**
 * Helper function to get validated and sorted dates from custom range
 */
function getDatesFromCustomRange(customRange) {
  var rawDates = customRange.getValues();
  var validDates = [];
  var validDatesMap = {};
  for (var r = 0; r < customRange.getNumRows(); r++) {
    for (var c = 0; c < customRange.getNumColumns(); c++) {
      var raw = Date.parse(rawDates[r][c]);
      if (!isNaN(raw)) {
        var rawDate = new Date(raw);
        var rawDateTime = getStartOfDayDate(rawDate).getTime();
        if (!validDatesMap[rawDateTime]) {
          // only push if raw date time is not added before
          validDates.push(rawDate);
          validDatesMap[rawDateTime] = true;
        }
      }
    }
  }

  // ascending sort
  validDates.sort(function(a, b) {
    return a - b;
  });
  return validDates;
}

/**
 * Helper function to filter dates that is later than the given date
 * @param {Array} dates. array of dates to filter
 * @param {Date} pivotDate. the date filter to apply
 * @param {int} maxDays, maximum days in the filter result, -1 indicates no limit - optional.
 */
function filterDates(dates, pivotDate, maxDays) {

  if (!dates || dates.constructor !== Array) {
    throw 'Invalid dates given';
  }

  usePivot = true;
  if (!pivotDate || pivotDate.constructor !== Date) {
    usePivot = false;
  }

  if (maxDays === undefined || isNaN(maxDays)) {
    maxDays = -1;
  }

  var closestDateIndex = 0;
  if (usePivot) {
    // look for pivot date
    var pivotDay = getStartOfDayDate(pivotDate);
    var possibleDateIndex = binaryIndexOf.call(dates, pivotDay);

    Logger.log('pivot day: ' + pivotDay);
    Logger.log('possible index: ' + possibleDateIndex);

    if ((pivotDay - getStartOfDayDate(dates[possibleDateIndex])) === 0) {
      Logger.log('closest: ' + dates[possibleDateIndex]);
      // pivot day found
      closestDateIndex = possibleDateIndex;
    } else if (possibleDateIndex + 1 < dates.length && (today - getStartOfDayDate(dates[possibleDateIndex + 1])) === 0) {
      Logger.log('closest + 1: ' + dates[possibleDateIndex + 1]);
      // the day after pivot day
      closestDateIndex = possibleDateIndex + 1;
    } else if (possibleDateIndex - 1 >= 0 && (today - getStartOfDayDate(dates[possibleDateIndex - 1])) === 0) {
      Logger.log('closest - 1: ' + dates[possibleDateIndex - 1]);
      // the day before pivot day
      closestDateIndex = possibleDateIndex - 1;
    } else {
      throw 'Pivot date not found in the given dates';
    }
  }

  var filteredDates = [];
  var filteredDateIndex = closestDateIndex;
  for (var i = 0; i < dates.length && (maxDays < 0 || i < maxDays); i++) {
    if (filteredDateIndex < dates.length) {
      filteredDates.push(new Date(dates[filteredDateIndex]));
      filteredDateIndex++;
    }
  }
  return filteredDates;
}

/*
 * Helper function to get selected range in current active sheet
 */
function getSelectedRange() {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    return sheet.getActiveRange().getA1Notation();
  } catch (e) {
    Logger.log('Get Selected Range exception: %s', e);
    throw 'No range selected.';
  }
}

/*
 * Helper function to get selected range with sheet name in current active sheet
 */
function getSelectedRangeWithSheetname() {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    return {
      range: sheet.getActiveRange().getA1Notation(),
      sheetname: sheet.getName()
    };
  } catch (e) {
    Logger.log('Get Selected Range with Sheet Name exception: %s', e);
    throw 'No range selected.';
  }
}

/*
 * Helper function to get range via A1 Notation in current active sheet
 */
function getRangeFromA1Notation(A1Notation) {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    return sheet.getRange(A1Notation);
  } catch (e) {
    Logger.log('Get Range From A1 Notation exception: %s', e);
    throw 'Invalid A1 Notation [' + A1Notation + '] for range.';
  }
}

/*
 * Helper function to get range via A1 Notation in the given sheet nane
 */
function getRangeFromSheetA1Notation(sheetname, A1Notation) {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName(sheetname);
    return sheet.getRange(A1Notation);
  } catch (e) {
    Logger.log('Get Range From Sheet A1 Notation exception: %s', e);
    throw 'Invalid A1 Notation [' + A1Notation + '] for sheet [' + sheetname + '].';
  }
}

/**
 * Helper function to determine if given range contains only 1 column
 */
function isSingleColumnRange(range) {
  var rangeColumn = range.getNumColumns();
  return rangeColumn === 1;
}

/**
 * Helper function to determine if given range contains only 1 row
 */
function isSingleRowRange(range) {
  var rangeRow = range.getNumRows();
  return rangeRow === 1;
}

/**
 * Helper function to filter out invalid days in week.
 * @output {Array} array of days in week, values (1-7)
 */
function filterDaysInWeek(daysInWeek) {
  var days = [];
  if (typeof daysInWeek === 'string') {
    days = daysInWeek.split(',');
  } else if (daysInWeek.constructor === Array) {
    days = daysInWeek;
  }

  var validatedDays = {};
  for (var d in days) {
    var number = parseInt(days[d]);
    if (number > 0 && number < 8) {
      validatedDays[number] = 1;
    }
  }

  var validated = [];
  for (var d in validatedDays) {
    validated.push(d);
  }

  return validated;
}

MINUTE_IN_SECONDS = 60;
HOUR_IN_SECONDS = 3600;
DAY_IN_SECONDS = 86400;
/**
 * Helper function to add duration to JavaScript datetime
 * supports only 's' seconds, 'm' minutes, 'h' hours, 'd' days
 */
function updateDate(date, time_unit, time_unit_scalar) {
  if (!date || date.constructor !== Date) {
    throw 'Invalid date';
  }

  var time_unit_seconds = 0;
  if (time_unit === 's') {
    time_unit_seconds = 1;
  } else if (time_unit === 'm') {
    time_unit_seconds = MINUTE_IN_SECONDS;
  } else if (time_unit === 'h') {
    time_unit_seconds = HOUR_IN_SECONDS;
  } else if (time_unit === 'd') {
    time_unit_seconds = DAY_IN_SECONDS;
  } else {
    throw 'Unsupported time unit [' + time_unit + ']';
  }

  if (isNaN(time_unit_scalar)) {
    throw 'Invalid scalar number for time unit';
  }

  var newMiliseconds = date.getTime() + (time_unit_scalar * time_unit_seconds * 1000);
  return new Date(newMiliseconds);
}

/**
 * Helper function to calculate days between start & end dates with days in week to filter
 */
function getDaysBetweenForWeek(startDate, endDate, daysInWeek) {
  if (!startDate || startDate.constructor !== Date) {
    startDate = new Date();
  }

  if (!endDate || endDate.constructor !== Date) {
    endDate = new Date();
  }

  if (!daysInWeek || daysInWeek.constructor !== Array || daysInWeek.length <= 0) {
    return getDaysBetween(startDate, endDate);
  }

  var startDay = startDate.getDay();
  var endDay = endDate.getDay();
  var betweenDays = getDaysBetween(startDate, endDate);

  var startToWeekEnd = 7 - startDay;
  var endToWeekStart = endDay;
  var days = 0;
  for (var i in daysInWeek) {
    if (daysInWeek[i] >= startToWeekEnd) {
      days++;
    }
  }
  for (var i in daysInWeek) {
    if (daysInWeek[i] <= endToWeekStart) {
      days++;
    }
  }
  var remainingDays = betweenDays - days;
  if (remainingDays > 0) {
    var numberOfWeeks = Math.ceil(remainingDays / 7);
    days += (numberOfWeeks * daysInWeek.length);
  }

  return days;
}

/**
 * Helper function to calculate days between start & end dates
 */
function getDaysBetween(startDate, endDate) {
  if (!startDate || startDate.constructor !== Date) {
    startDate = new Date();
  }

  if (!endDate || endDate.constructor !== Date) {
    endDate = new Date();
  }

  // Take the difference between the dates and divide by milliseconds per day.
  // Round to nearest whole number to deal with DST.
  return Math.round((endDate - startDate) / (1000 * DAY_IN_SECONDS));
}

/**
 * Helper function to get start of the day date (00:00:00.000)
 */
function getStartOfDayDate(date) {
  if (!date || date.constructor !== Date) {
    throw 'Invalid date for start of day';
  }

  var startOfDay = new Date(date);
  startOfDay.setHours(0);
  startOfDay.setMinutes(0);
  startOfDay.setSeconds(0);
  startOfDay.setMilliseconds(0);
  return startOfDay;
}

/**
 * Helper function to get end of the day date (23:59:59.999)
 */
function getEndOfDayDate(date) {
  if (!date || date.constructor !== Date) {
    throw 'Invalid date for end of day';
  }

  var endOfDay = new Date(date);
  endOfDay.setHours(23);
  endOfDay.setMinutes(59);
  endOfDay.setSeconds(59);
  endOfDay.setMilliseconds(999);
  return endOfDay;
}

/**
 * https://oli.me.uk/2013/06/08/searching-javascript-arrays-with-a-binary-search/
 *
 * Performs a binary search on the host array. This method can either be
 * injected into Array.prototype or called with a specified scope like this:
 * binaryIndexOf.call(someArray, searchElement);
 *
 * @param {*} searchElement The item to search for within the array.
 * @return {Number} The index of the element which defaults to -1 when not found.
 */
function binaryIndexOf(searchElement) {
  'use strict';

  var minIndex = 0;
  var maxIndex = this.length - 1;
  var currentIndex;
  var currentElement;

  while (minIndex <= maxIndex) {
    currentIndex = (minIndex + maxIndex) / 2 | 0;
    currentElement = this[currentIndex];

    if (currentElement < searchElement) {
      minIndex = currentIndex + 1;
    } else if (currentElement > searchElement) {
      maxIndex = currentIndex - 1;
    } else {
      return currentIndex;
    }
  }

  // return -1;
  /* Return last visited index */
  return currentIndex;
}

/*
 * Helper function to get the default configurations for the given sheet (default to current active sheet)
 * Note: Sheet refers to fill up
 */
function getDefaultConfig(sheetname) {
  if (!sheetname) {
    var sheet = SpreadsheetApp.getActiveSheet();
    sheetname = sheet.getName();
    Logger.log('getDefaultConfig: sheetname not given, using current active sheet[' + sheetname + ']');
  } else {
    var givenSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname);
    if (givenSheet == null) {
      throw 'Sheetname[' + sheetname + '] given not found';
    }
    sheetname = givenSheet.getName();
  }

  return {
    sheet_name: sheetname,
    range_person_name: '',
    range_timeslot: '',
    range_timestamp: '',

    start_date: undefined,
    end_date: undefined,
    frequency: '',
    days_in_week: [],

    custom_dates_sheet_name: '',
    custom_dates_range: '',

    lookup_sheet_name: '',
    lookup_range_person_name: '',
    lookup_range_timeslot: '',
    lookup_range_timestamp: '',
    lookup_data_retention_days: -1
  };
}

function getDefaultConfigNames() {
  return {
    sheet_name: 'FILLUP_SHEETNAME',
    range_person_name: 'FILLUP_RANGE_PERSON_NAME',
    range_timeslot: 'FILLUP_RANGE_TIMESLOT',
    range_timestamp: 'FILLUP_RANGE_TIMESTAMP',

    start_date: 'FILLUP_START_DATE',
    end_date: 'FILLUP_END_DATE',
    frequency: 'FILLUP_FREQUENCY',

    days_in_week: 'FILLUP_DAYS_IN_WEEK',
    custom_dates_sheet_name: 'FILLUP_CUSTOM_DATES_SHEET_NAME',
    custom_dates_range: 'FILLUP_CUSTOM_DATES_RANGE',

    lookup_sheet_name: 'LOOKUP_SHEET_NAME',
    lookup_range_person_name: 'LOOKUP_RANGE_PERSON_NAME',
    lookup_range_timeslot: 'LOOKUP_RANGE_TIMESLOT',
    lookup_range_timestamp: 'LOOKUP_RANGE_TIMESTAMP',
    lookup_data_retention_days: 'LOOKUP_DATE_RETENTION_DAYS'
  };
}

function getDefaultSheetConfigNames(sheetname) {
  var config = getDefaultConfigNames();
  for (var c in config) {
    config[c] = sheetname + '_' + config[c];
  }
  return config;
}

/*
 * Helper function to read the configurations from Document properties service
 */
function readConfig(sheetname) {

  var config = getDefaultConfig(sheetname);
  var configNames = getDefaultSheetConfigNames(sheetname);
  var outputConfig = {};

  try {
    var props = PropertiesService.getDocumentProperties();

    for (var c in config) {
      // Logger.log(c + '@' + config[c] + '@' + configNames[c]);
      var configProp = props.getProperty(configNames[c]);
      if (configProp == null) {
        // read from default config
        outputConfig[c] = config[c];
      } else {
        outputConfig[c] = configProp;
      }

      switch (c) {
        case 'start_date':
        case 'end_date':
          {
            var dateISO = outputConfig[c];
            if (dateISO) {
              outputConfig[c] = new Date(dateISO);
            }
            break;
          }
        case 'days_in_week':
          {
            var days_in_week_string = outputConfig[c];
            var days_in_week_arr = [];
            if (typeof days_in_week_string === 'string' && days_in_week_string) {
              days_in_week_string.split(',').forEach(function(v, i) {
                days_in_week_arr.push(v);
              });
            }
            outputConfig[c] = days_in_week_arr;
            break;
          }

      }
    }
  } catch (e) {
    Logger.log('Read Config exception: %s', e);
    throw 'Unable to read config for the sheet.';
  }
  return outputConfig;
}

/*
 * Helper function to remove the configurations in Document properties service
 */
function removeConfig(config) {
  var sheetname = '';
  if (typeof config === 'string') {
    sheetname = config;
  } else {
    sheetname = config.sheet_name;
  }
  var sheetConfigNames = getDefaultSheetConfigNames(sheetname);
  var props = PropertiesService.getDocumentProperties();
  for (var n in sheetConfigNames) {
    props.deleteProperty(sheetConfigNames[n]);
  }
}

/*
 * Helper function to save the configurations to Document properties service
 */
function saveConfig(config) {

  var sheetConfig = readConfig(config.sheet_name);
  var configNames = getDefaultSheetConfigNames(config.sheet_name);

  try {

    // map all values in given config to existing config
    for (var c in config) {
      if (config[c] === undefined || configNames[c] === undefined) {
        // skip unknown config
        continue;
      }
      sheetConfig[c] = config[c];

      switch (c) {
        case 'start_date':
        case 'end_date':
          {
            var dateISO = sheetConfig[c];
            if (dateISO) {
              sheetConfig[c] = new Date(sheetConfig[c]);
            }
            break;
          }
        case 'days_in_week':
          {
            var days_in_week_string = '';
            if (sheetConfig[c] && sheetConfig[c].constructor === Array) {
              days_in_week_string = sheetConfig[c].join(',');
            }
            sheetConfig[c] = days_in_week_string;
            break;
          }
      }
    }

    var props = PropertiesService.getDocumentProperties();
    var unsavedProps = {};

    for (var c in sheetConfig) {
      unsavedProps[configNames[c]] = sheetConfig[c];
    }

    props.setProperties(unsavedProps);

  } catch (e) {
    Logger.log('Save Config exception: %s', e);
    throw 'Unable to save config for the sheet.';
  }
}

function sheetConfigProperty(sheetname, propertyName) {
  return sheetname + '_' + propertyName;
}
