/**
* Get sheet object from current active spreadsheet
*
* @param {string} sheetName the name of the sheet
* @return {sheet} the sheet object
*/
function getSheet(sheetName){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(sheetName);
}

/**
* Check if the non-empty data range of the sheet is editable
*
* @param {sheet} sheet the sheet object
* @return {boolean} boolean that indicates if the sheet is editable
*/
function isSheetEditable(sheet){
  return sheet.getDataRange().canEdit();
}

/**
* Purge old data of the sheet
*
* @param {sheet} sheet the sheet object
* @param {number} offsetRow number of rows to offset (Usually is the number of rows for headers)
* @param {number} timestampCol the column where the timestamp is located at
* @param {number} expiryDays days to keep the data before it expires
*/
function purgeOldData(sheet, offsetRow, timestampCol, expiryDays){

  // offset starts from 1,2,3..
  offsetRow = offsetRow || 1;  
  timestampCol = timestampCol || 3;
  expiryDays = expiryDays || -1;
  
  var moment = Moment.load();
  var today = new Date();
  
  if(sheet && expiryDays >= 0){
    Logger.log(expiryDays);
    sheet.sort(timestampCol);
    
    var expiredData = sheet.getDataRange().getValues();
    var expiredRowIndex = -1;
    
    // rows
    for(var u=0; u<expiredData.length; u++){
      
      if(u < offsetRow){
        //skip for headers
        continue;
      } 
      
      var filledTimestamp = expiredData[u][timestampCol-1];
      var filledDatetime = new Date(filledTimestamp);
      var filledMoment = moment(filledDatetime);
      
      // get the row index just now the filled data expired
      var todayMoment = moment(moment(today).format('DD-MM-YYYY') + ' 00:00', 'DD-MM-YYYY');
      var isFilledExpired = moment(todayMoment).add(-expiryDays, 'days').isBefore(filledMoment);
      
      //Logger.log("today:"+todayMoment.format('DD-MM-YYYY hh:mm')+" filled:"+filledMoment.format('DD-MM-YYYY hh:mm'));
      //Logger.log("expired:"+isFilledExpired+" expDays:"+expiryDays);

      if(isFilledExpired){
        expiredRowIndex = u;
        break;
      }
    }
    
    if(expiredRowIndex > 0){
      try{
        Logger.log("Deleting row "+ (offsetRow +1) + " to row " + ( expiredRowIndex+1));
        // deleteRows starts from 1, 2,..
        sheet.deleteRows(offsetRow + 1, expiredRowIndex+1);
      }catch(err){
        Logger.log(err);
      }
    }
  }
}

/**
* Populate unavailabilitis from 'lookup' sheet into 'fillup' sheet
*
* <p>
* Look up names & availabilities in 'lookup' sheet
* and fill up 'fillup' sheet with attandance status (On Time, Late, Unavailable)
* </p>
* <p>
* It will also perform purgeOldData()
* </p>
*
* <p>
* Currently the availabities must be in the combination of '[Status] On [Day]'
* </p>
* <p>
* For e.g Unavailable On Monday, Late On Tuesday, ...
* </p>
* 
* <pre>
* Look Up
* {string} sheetName the name of the sheet
* {number} nameCol column of name in the sheet
* {number} availabilityCol column of availability in the sheet
* {number} timestampCol column of timestamp in the sheet
* {number} rowOffset number of rows to offset (Usually is the number of rows for headers)
* {number} expiryDays days to keep the data before it expires (-1 - keep forever)
* </pre>
*
* <pre>
* Fill Up
* {string} sheetName the name of the sheet
* {number} namecol column of name in the sheet
* {Array<number>} days array of the days which attendance will be filled in (1-Mon, ... ,7-Sun)
* {Array<number>} daysCol array of the columns for the days
* {number} rowOffset number of rows to offset (Usually is the number of rows for headers)
* </pre>
*
* Example
* <pre>
*	populateUnavailability(
*    {
*      lookup:
*      {
*        sheetName: "Raid Unavailability",
*        nameCol: 1,
*        availabilityCol: 2,
*        timestampCol: 3,
*        rowOffset: 1,
*        expiryDays: -1
*      },
*      fillup:
*      {
*        sheetName: "Roster",
*        namecol: 2,
*        days: [1,2,4],
*        dayscol: [4,5,6],
*        rowOffset: 2,
*        rowCount: 21
*      }
*    }
*  );
* </pre>
*/
function populateUnavailability(options){
  
  var moment = Moment.load();
  
  options = options || {};
  
  var fillup = options.fillup || {};
  var lookup = options.lookup || {};
  
  var sheetFillUp = getSheet(fillup.sheetName || "");
  var sheetLookUp = getSheet(lookup.sheetName || "");
  
  if(sheetFillUp == null || sheetLookUp == null){
    Logger.log("Invalid sheets [" + (sheetLookUp ? "fill up" : "look up") + "]");
    return;
  }else if(!isSheetEditable(sheetFillUp) || !isSheetEditable(sheetLookUp)){
    Logger.log("Insufficient permission to edit the sheets");
    return;
  }
  
  purgeOldData(sheetLookUp, lookup.rowOffset, lookup.timestampCol, lookup.expiryDays);
  
  sheetLookUp = getSheet(lookup.sheetName || "");
  
  if(!sheetLookUp || !sheetFillUp){
    Logger.log("Invalid sheet [" + (sheetLookUp ? "fill up" : "look up") + "]");
    return;
  }else if(!isSheetEditable(sheetFillUp) || !isSheetEditable(sheetLookUp)){
    Logger.log("Insufficient permission to edit the sheets");
    return;
  }
  
  // defaults
  var lookupsheet = {
    namecol: lookup.nameCol || 1,
    availabilitycol: lookup.availabilityCol || 2,
    timestampcol: lookup.timestampCol || 3,
    rowoffset: lookup.rowOffset || 1,
    expirydays: lookup.expiryDays || -1
  };
  
  var fillupsheet = {
    namecol: fillup.namecol || 2,
    dayscol: fillup.dayscol || [4,5,6],
    // must be ascending - 0, 1, 2 ..
    days: fillup.days || [1,2,4],
    rowoffset: fillup.rowOffset || 2,
    rowcount: fillup.rowCount || 21
  };
  
  // remove duplicates
  var uniqueDays = [];
  for(var d in fillupsheet.days){
    if(uniqueDays.indexOf(fillupsheet.days[d]) === -1){
      uniqueDays.push(fillupsheet.days[d]);
    }
  }
  // sort days
  uniqueDays.sort();
  fillupsheet.days = uniqueDays;

  var dayCount = fillupsheet.days.length;
  var dayColCount = fillupsheet.dayscol.length;
  
  if(dayCount <= 0){
    Logger.log("No days configured");
    return;
  }

  if(dayColCount != dayCount){
    Logger.log("Days and DaysCol must be the same length");
    return;
  }
  
  var today = new Date();
  var todayMoment = moment(today);
  
  // Sun -> 0, Mon -> 1
  var dayOfWeekToday = today.getDay();
  
  
  // refresh sheetFillUp column names
  // change when today is one of the raid days
  
  var upcomingDay = -1;
  var upcomingDayIndex = -1;
  for(var d=0; d<dayCount; d++){
    
    if(fillupsheet.days[d] >= dayOfWeekToday){
      // within this week
      upcomingDay = fillupsheet.days[d];
      upcomingDayIndex = d;
      break;
    }
  }
  
  if(upcomingDay < 0){
    // next week
    upcomingDay = fillupsheet.days[0];
    upcomingDayIndex = 0;
  }
  

  for(var c=0; c<dayColCount; c++){
    // loop through each day column
    var daysIndex = (upcomingDayIndex + c ) % dayCount;
    var day = fillupsheet.days[daysIndex];

    var col = fillupsheet.dayscol[c];

    var daysToAdd = day - dayOfWeekToday;
    if(daysToAdd < 0){
      daysToAdd += 7;
    }
    
    // Logger.log("today:"+dayOfWeekToday+" day:"+day+" add:"+daysToAdd);

    var dayMoment = moment(today).add(daysToAdd, 'days');
    
    if(fillupsheet.rowoffset == 1){
      sheetFillUp.getRange(1, col).setValue(dayMoment.format("Do MMMM"));
    }else if(fillupsheet.rowoffset > 1){
      sheetFillUp.getRange(1, col).setValue(dayMoment.format("dddd"));
      sheetFillUp.getRange(2, col).setValue(dayMoment.format("Do MMMM"));
    }
  }

  var sheetFillUpAllData = sheetFillUp.getDataRange();
  var sheetLookUpAllData = sheetLookUp.getDataRange();

  // get sortale range of lookupsheet, e.g. skip sorting for offseted rows
  var lastRow = Math.max(lookupsheet.rowoffset+1, sheetLookUpAllData.getNumRows());
  var lastCol = Math.max(1, sheetLookUpAllData.getNumColumns());
  var sortableRange = sheetLookUp.getRange(lookupsheet.rowoffset+1,1, lastRow , lastCol);
  
  //Logger.log(lookupsheet.rowoffset +"-"+lastRow +"-"+lastCol);
  
  // sort timestampcol in lookup sheet
  sortableRange.sort(lookupsheet.timestampcol);
  
  var sheetFillUpData = sheetFillUpAllData.getValues();
  var sheetLookUpData = sheetLookUpAllData.getValues();
  
  var lookupRegexUnavail = [];
  var lookupRegexLate = [];

  for(var d in fillupsheet.days){
    var dayOfWeek = moment().weekday(fillupsheet.days[d]).format('dddd');
    lookupRegexUnavail.push("\\bUnavailable On "+dayOfWeek+"\\b");
    lookupRegexLate.push("\\bLate On "+dayOfWeek+"\\b");
  }

  // sheetFillUp row
  var startRow = fillupsheet.rowoffset >= 0 ? fillupsheet.rowoffset : 0;
  var endRow = startRow + Math.abs(fillupsheet.rowcount);
  
  for(var r=startRow ; r<Math.min(endRow, sheetFillUpData.length); r++){
    
    //Logger.log("fillup row "+r+" of "+sheetFillUpData.length);
    
    var sheetFillUpCharName = sheetFillUpData[r][fillupsheet.namecol-1];
    
    // Logger.log(sheetFillUpCharName);

    if(!sheetFillUpCharName || sheetFillUpCharName == ""){
      // skip invalid character name
      Logger.log("Invalid look up name ["+sheetFillUpCharName+"]");
      continue;
    }
    
    var isUnavailFilled = false;
    
    // unavail rows
    for(var u=0; u<sheetLookUpData.length; u++){
      
      // unavail columns 
      if(u < lookupsheet.rowoffset){
        //skip 0 for headers
        continue;
      }
      
      var unavailCharName = sheetLookUpData[u][lookupsheet.namecol-1];

      if(!unavailCharName || unavailCharName == ""){
        // skip invalid character name
        Logger.log("Invalid filled name ["+unavailCharName+"]");
        continue;
      }
      
      //Logger.log(unavailCharName);
      
      var charNameRegex = new RegExp(sheetFillUpCharName,"i");

      if(charNameRegex.test(unavailCharName)) {

        //Logger.log("filled");
        
        isUnavailFilled = true;
        
        var dates = sheetLookUpData[u][lookupsheet.availabilitycol-1];
        var filledTimestamp = sheetLookUpData[u][lookupsheet.timestampcol-1];
        var filledDatetime = new Date(filledTimestamp);
        var filledMoment = moment(filledDatetime);
        var startMoment = moment(filledMoment.format('DD-MM-YYYY'), 'DD-MM-YYYY');
        var endMoment = moment(startMoment).add(8,'days');
        
        // filled data only valid for the following 1 week
        var todayMoment = moment(today);
        var isFilledValid = todayMoment.isAfter(startMoment) && todayMoment.isBefore(endMoment);
        
        Logger.log(sheetFillUpCharName+" filled ["+dates+"]");
        
        for(var dCol=0; dCol<dayColCount; dCol++){

          var daysIndex = (upcomingDayIndex + dCol ) % dayCount;
          var regexU = new RegExp(lookupRegexUnavail[daysIndex], "i");
          var regexL = new RegExp(lookupRegexLate[daysIndex], "i");

          // Logger.log("valid:" + isFilledValid +" regex:"+lookupRegexUnavail[daysIndex]);
          
          var status = "On Time";
          if(isFilledValid && regexU.test(dates)){
            status = "Unavailable";
          }else if(isFilledValid && regexL.test(dates)){
            status = "Late";
          }

          sheetFillUp.getRange(r+1, fillupsheet.dayscol[dCol]).setValue(status);
        }
      }
    }
    
    // Unavail not filled
    if(!isUnavailFilled){

      for(var dCol=0; dCol<dayColCount; dCol++){
        sheetFillUp.getRange(r+1, fillupsheet.dayscol[dCol]).setValue("On Time");
      }
    }
  }
}

// TODO permenant status for name

function test_(){
  populateUnavailability(
    {
      lookup:
      {
        sheetName: "Raid Unavailability",
        nameCol: 1,
        availabilityCol: 2,
        timestampCol: 3,
        rowOffset: 1,
        expiryDays: -1
      },
      fillup:
      {
        sheetName: "Roster",
        namecol: 2,
        dayscol: [4,5,6],
        days: [1,2,4],
        rowOffset: 2,
        rowCount: 20
      }
    });
}
