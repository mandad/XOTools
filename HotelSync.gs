const monthString = {1: "JAN",
                    2: "FEB",
                    3: "MAR",
                    4: "APR",
                    10: "OCT",
                    11: "NOV",
                    12: "DEC"};
const timeZoneShift = 10; //hours
var nameRes = "";
var nameCal = "";

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Sync Tools').
  addItem('Mark Hotel Actions', 'markHotelDates').addToUi();
}

function markHotelDates() {
  var trackerSheet = SpreadsheetApp.getActiveSpreadsheet()
  var hotelSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1S2n18o8XyPAB6Am__upM3fYKxqFa_kV6SgdxNAiFJTk/edit");
  var reservationSheet = hotelSheet.getSheetByName("Sette");

  let row = 32;
  var name = "Test";

  //Clear existing marks first to update
  clearSheets(trackerSheet);
  
  //Loop through all the reservations
  while (name != "END") {
    //Get one reservation
    name = reservationSheet.getRange(row, 1).getValue();
    //skip blank rows
    if (name != "" && name != "END") {
      //Get all the supporting dates from the reservation sheet
      //Need to shift into UTC time zone
      let startDate = new Date(reservationSheet.getRange(row, 2).getValue().getTime() + timeZoneShift * 60 * 60 * 1000);
      let endDate = new Date(reservationSheet.getRange(row, 3).getValue().getTime() + timeZoneShift * 60 * 60 * 1000);
      let startSheetString = monthString[startDate.getMonth()+1];
      let monthSheetStart = trackerSheet.getSheetByName(monthString[startDate.getMonth()+1]);
      let monthSheetEnd = trackerSheet.getSheetByName(monthString[endDate.getMonth()+1]);
      
      //find row in personnel tracker on that month
      let calName = nameAssoc(trackerSheet, name);
      if (calName != "") {
        markBorder(monthSheetStart, startDate, calName, true);
        markBorder(monthSheetEnd, endDate, calName, false);
      }
      else {
        Logger.log("Name not found: " + name);
      }
    }
    row = row + 1;
  }

}

function markBorder(monthSheet, markDate, calName, begin = true) {
  for (let nameRow = 4; nameRow < monthSheet.getLastRow()+1; nameRow++) {
    let testName = monthSheet.getRange(nameRow, 2).getValue();
    if (testName == calName) {
      //Zero indexed within month, day 1 is column 3 ("C")
      let markRange = monthSheet.getRange(nameRow, markDate.getDate() + 2);
      if (begin) {
        markRange.setBorder(null, true, null, null, null, null, '#38761d', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      } else {
        markRange.setBorder(null, null, null, true, null, null, '#85200c', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        //markRange.setBorder(null, null, null, true, null, null, '#85200c', SpreadsheetApp.BorderStyle.DOUBLE);
      }
      return;
    }
  }
}

function nameAssoc(trackerSheet, thisResName) {
  //Check if the name has already been found
  if (nameRes != thisResName) {
    let nameMapRange = trackerSheet.getSheetByName("Names").getRange("A2").getDataRegion().getValues();
    //Loop through the map range
    for (let row = 1; row < nameMapRange.length; row++) {
      if (nameMapRange[row][1] == thisResName) {
        nameRes = thisResName;
        nameCal = nameMapRange[row][0];
        return nameCal;
      }
    }
  } else {
    //If it was found in a previous loop
    return nameCal;
  }
  //If not matching name was found
  return "";
}

function clearSheets(trackerSheet) {
  for (var sheetName in monthString) {
    let clearSheet = trackerSheet.getSheetByName(monthString[sheetName]);
    let clearRange = clearSheet.getRange(4,3,clearSheet.getLastRow()-4,clearSheet.getLastColumn()-3);
    clearRange.setBorder(null, false, null, false, false, null);
  }
}

function monthSheet(date) {
  return Utilities.formatDate(date, timeZone, 'mmm').toUpperCase();
}
