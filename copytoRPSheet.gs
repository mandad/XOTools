function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // Create custom menu in Google Sheet
  ui.createMenu('Augmenter Tools')
      .addItem('Clear Sheet', 'ClearSheet')
      .addItem('Copy Values to RP Sheet', 'copyValuesToRP')
      .addToUi();
}

function ClearSheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D9:J44').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('L9:R44').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('C48:C61').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('N48:T61').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('D62:J62').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('O62:T62').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('D7').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('E7').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('F7').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('G7').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('H7').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('I7').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('J7').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('L7').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('M7').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('N7').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('O7').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('P7').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('Q7').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('R7').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRange('R8').activate();
};

//Fun note - this script was almost entirely generated by ChatGPT
function copyValuesToRP() {
  // Prompt the user for the source and destination spreadsheet IDs
  const ui = SpreadsheetApp.getUi();
  //const sourceSpreadsheetId = ui.prompt("Enter the ID of the source spreadsheet:").getResponseText();
  const destinationSpreadsheetId = ui.prompt("Enter the ID of the destination spreadsheet:").getResponseText();
  
  // Get the source and destination spreadsheets
  //const sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const destinationSpreadsheet = SpreadsheetApp.openById(destinationSpreadsheetId);

  var shoreCopy = ui.alert('Copy Shore Leave Record', 'Do you want to copy the shore leave records?', ui.ButtonSet.YES_NO);
  var rangesToCopy = [];
  
  // Define the ranges to copy
  if (shoreCopy == ui.Button.YES) {
    rangesToCopy = [    { sourceRange: "D7:J7", destinationRange: "D7:J7" }, { sourceRange: "L7:R7", destinationRange: "L7:R7" }, { sourceRange: "D9:J44", destinationRange: "D9:J44" },    { sourceRange: "L9:R44", destinationRange: "L9:R44" },    { sourceRange: "C48:C61", destinationRange: "C48:C61" },    { sourceRange: "N48:N61", destinationRange: "N48:N61" }  ];
  } else {
    rangesToCopy = [    { sourceRange: "D9:J44", destinationRange: "D9:J44" },    { sourceRange: "L9:R44", destinationRange: "L9:R44" },    { sourceRange: "C48:C61", destinationRange: "C48:C61" },    { sourceRange: "N48:N61", destinationRange: "N48:N61" }  ];
  }
  // Loop through all sheets in the source spreadsheet
  // Sheets 0 and 1 are info and template
  const sourceSheets = sourceSpreadsheet.getSheets();
  for (let i = 2; i < sourceSheets.length; i++) {
    let sourceSheet = sourceSheets[i];
    const sourceSheetName = sourceSheet.getName();
    
    // Get the destination sheet with the same name as the source sheet
    const destinationSheet = destinationSpreadsheet.getSheetByName(sourceSheetName);
    
    if (destinationSheet) {
      // Loop through the ranges to copy
      for (let j = 0; j < rangesToCopy.length; j++) {
        const sourceRange = sourceSheet.getRange(rangesToCopy[j].sourceRange);
        //const sourceValues = sourceRange.getValues();
        
        let destinationRange = destinationSheet.getRange(rangesToCopy[j].destinationRange);
        //destinationRange.setValues(sourceValues);
        
        //Note this will skip the shore leave credit days already entered but this is probably ok for now
        copyIfBlank(sourceRange, destinationRange, sourceSheetName);
      }
    }
  }
}

function copyIfBlank(srcRange, destRange, copySheetName) {
  // Validate the input ranges
  if (!srcRange || !destRange) {
    throw new Error("Both source and destination ranges are required.");
  }
  
  if (srcRange.getNumRows() !== destRange.getNumRows() || srcRange.getNumColumns() !== destRange.getNumColumns()) {
    throw new Error("Source and destination ranges must be of the same size.");
  }
  
  // Get the values from both ranges
  var srcValues = srcRange.getValues();
  var destValues = destRange.getValues();
  
  // Store conflicts
  var conflicts = [];
  
  // Iterate over the values and copy from source to destination if destination cell is blank or zero
  for (var i = 0; i < srcValues.length; i++) {
    for (var j = 0; j < srcValues[i].length; j++) {
      if (!destValues[i][j] || destValues[i][j] === " ") {
        destValues[i][j] = srcValues[i][j];
      } else if (destValues[i][j] !== srcValues[i][j]) {
        // Capture conflicts
        var cell = destRange.offset(i, j).getCell(1, 1).getA1Notation();
        conflicts.push("Cell " + cell + ": Source (" + srcValues[i][j] + ") vs. Destination (" + destValues[i][j] + ")");
      }
    }
  }
  
  // Set the values back to the destination range
  destRange.setValues(destValues);
  
  // If there are conflicts, show them to the user
  if (conflicts.length > 0) {
    var message = "Conflicts found for " + copySheetName +":\n\n" + conflicts.join("\n");
    SpreadsheetApp.getUi().alert(message);
  }
}
