let oblogCosts = null;
const marsTopOffset = 4;  // Offset from marsTrans.getDataRange() to beginning of actual data
const marsDataRangeOffset = 1;  //Offset from top of sheet to marsTrans.getDataRange()
const oblogDataStart = 9;   // First data row of OBLOG table

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('OBLOG Tools').
  addItem('Reconcile MARS', 'gatherInput').addToUi();
}

function gatherInput() {
  var ui = SpreadsheetApp.getUi();

  /*
  * Runs on a different OBLOG sheet
  * Mostly for initial testing/running script without sheet

  //Get OBLOG Path
  var result = ui.prompt(
      'Oblog Google Sheet Path (include /edit):',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var oblogPath = result.getResponseText();
  if (button == ui.Button.CLOSE || button == ui.Button.CANCEL) {
    return;
  }
  */
  
  //Get MARS Path
  result = ui.prompt(
    'MARS Transactions Google Sheet Path (include /edit):',
    ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  button = result.getSelectedButton();
  var marsPath = result.getResponseText();
  if (button == ui.Button.CLOSE || button == ui.Button.CANCEL) {
    return;
  }

  //Get Start Date
  result = ui.prompt(
    'Start Date in MARS for Reconcile (MM/DD/YYYY):',
    ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  button = result.getSelectedButton();
  var userDate = result.getResponseText();
  if (button == ui.Button.CLOSE || button == ui.Button.CANCEL) {
    return;
  }
  //No checking of formats or paths at this point
  var startDate = new Date(userDate);

  reconcileMARSReport(marsPath, startDate/*, oblogPath*/);
}

//Test or use this as a backup
function reconcileTest() {
  // Month is zero indexed
  reconcileMARSReport("https://docs.google.com/spreadsheets/d/1Izpedb1SZ-DunT8sZ8dJqEsNdL2NFuRXEBw9mJ11Xnc/edit", new Date(2022, 0, 1), "https://docs.google.com/spreadsheets/d/1mUyWZDMFOqjvaZz3XonzEHgZ03l_qANQP3wBHZo5824/edit");
}

function reconcileMARSReport(marsPath, startDate, oblogPath = null) {
  var oblogSheet = SpreadsheetApp.getActiveSpreadsheet()
  if (oblogPath != null) {
    oblogSheet = SpreadsheetApp.openByUrl(oblogPath);
  } 
  
  var marsSheet = SpreadsheetApp.openByUrl(marsPath);


  if (marsSheet != null && oblogSheet != null) {
    var marsTrans = marsSheet.getSheetByName("Transaction Report");
    var oblogTrans = oblogSheet.getSheetByName("OBLOG");

    //Add match column for marking
    marsTrans.getRange("N4").setValue("OBLOG Matched");

    //Add filter to MARS and sort
    //Right now assume if it has a filter it is correct
    marsFilter = marsTrans.getFilter()
    if (marsFilter == null) {
      var filterRange = marsTrans.getRange(marsTopOffset,1, marsTrans.getLastRow()-marsTopOffset, 14);
      var marsFilter = filterRange.createFilter();
    }
    marsFilter.sort(9,true);
    //Row 0 of marsData is row 2 of the spreadsheet
    var marsData = marsTrans.getDataRange().getValues();
    //Loop through rows of MARS data
    for (let i = marsTopOffset; i < marsData.length; i++) {
      //Only loop through those after the start date
      if (marsData[i][8].valueOf() < startDate.valueOf()) {
        continue;
      }
      
      //Only take items where net amount is nonzero
      if (marsData[i][12] != 0) {
        //Extract the lastname from PCARD transactions
        if (marsData[i][2] == "PCARD") {
          var lastName = String(marsData[i][3]).split("/")[1].split(" ")[1];
        }
        else {
          lastName = "";
        }
        //Logger.log(lastName)
        //found[0] = row, found[1] = MatchString
        var found = findCorrespondingOBLOG(oblogTrans, marsData[i][12], marsData[i][7], marsData[i][5], marsData[i][2], lastName);
        if (found[0] > 0) {
          Logger.log("Found Matching for row %s: %s, non-matched: %s", i, found[0], found[1]);
          // If full match, mark it off
          let markerString = found[1];
          if (found[1] == "") {
            oblogTrans.getRange(found[0], 14).setValue("Yes");
            markerString = "x";
          }
          //find other PCARD for same transaction (probably should just sort after the fact)
          for (let matchRow = 0; matchRow < marsData.length; matchRow++) {
            if (marsData[matchRow][3] == marsData[i][3]) {
              marsTrans.getRange(matchRow + marsDataRangeOffset,14).setValue(markerString);
            }
          }
        }
      }
    }
  } 
  else {
    Logger.log("Transaction Report not found in MARS sheet");
  }
}

function findAllMatchingMARS(description, maxRow) {
  for (let row = 0; row < maxRow; row++) {
    
  }
}

function findCorrespondingOBLOG(oblogSheet, cost, occ, projectCode, type, lastName) {
  Logger.log("Reconciling - Cost: " + cost + " OCC: " + occ + " Proj: " + projectCode + " type:" + type + " Name:" + lastName);
  
  let fullMatch = false;
  let startRow = oblogDataStart;
  let matchList = [];
  
  while (!fullMatch) {
    var matchRow = rowOfCost(oblogSheet, cost, startRow);
    var matchString = "";
    if (matchRow > 0) {
      var matchDetails = oblogSheet.getRange(matchRow, 6, 1, 9).getValues()[0];
      //Already reconciled
      if (matchDetails[8] == "Yes") {
        Logger.log("Found already reconciled");
        //String is shorter to prefer over one that doesn't match
        matchList.push([matchRow, "[Rec]"]);
        startRow = matchRow + 1;
        continue;
      }
      if (type != matchDetails[0]) {
        matchString += "[Type]";
      }
      if (lastName != String(matchDetails[1]).toUpperCase()){
        matchString += "[Name]";
      }
      if (!String(matchDetails[2]).startsWith(projectCode)) {
        matchString += "[Proj]";
      }
      if (!String(matchDetails[4]).startsWith(occ)) {
        matchString += "[OCC ]";
      }
      if (matchString == "") {
        fullMatch = true;
        return [matchRow, matchString];
      } else {
        //add the match quality to the list
        matchList.push([matchRow, matchString]);
      }
    } else {
      if (matchList.length == 0) {
        Logger.log("Corresponding OBLOG Entry Not Found by cost");
      }
      break;
    }
    startRow = matchRow + 1;
  }
  //If more than one partial match found, return the best
  if (matchList.length > 0) {
    let bestMatch = 0;
    let minDiff = matchList[0][1].length;
    for (var i = 0; i < matchList.length; i++) {
      //if this is a better match (less issues noted)
      if (matchList[i][1].length <= minDiff) {
        bestMatch = i;
      }
    }
    return matchList[bestMatch];
  }
  return [-1, ""];
}

function rowOfCost(oblogSheet, cost, startRow = oblogDataStart){
  //Set the oblog costs range if this is the first run
  if (oblogCosts == null) {
    var costData = oblogSheet.getRange("M:M").getValues();
    let ar=costData.map(x => x[0]); //turns 2D array to 1D array, so we can use indexOf
    const lastRow=ar.indexOf('');
    oblogCosts = oblogSheet.getRange(oblogDataStart, 13, lastRow-startRow+1).getValues();
    
    /*for (var i = startRow; i < oblogData.length; i++) {
      if (oblogData[i][12] == ""){
        oblogCosts = oblogSheet.getRange(startRow, 13, i-startRow+1).getValues();
        break;
      }
    }*/
  }

  for(var i = startRow - oblogDataStart; i < oblogCosts.length;i++){
    if(oblogCosts[i] == cost){
      Logger.log("Found on OBLOG row: " + (i+oblogDataStart));
      return i+oblogDataStart;
    }
  }
  return -1;
}