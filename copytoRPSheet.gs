//Fun note - this script was almost entirely generated by ChatGPT
function copyValuestoRP() {
  // Prompt the user for the source and destination spreadsheet IDs
  const ui = SpreadsheetApp.getUi();
  const sourceSpreadsheetId = ui.prompt("Enter the ID of the source spreadsheet:").getResponseText();
  const destinationSpreadsheetId = ui.prompt("Enter the ID of the destination spreadsheet:").getResponseText();
  
  // Get the source and destination spreadsheets
  const sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  const destinationSpreadsheet = SpreadsheetApp.openById(destinationSpreadsheetId);
  
  // Define the ranges to copy
  const rangesToCopy = [    { sourceRange: "D7:J7", destinationRange: "D7:J7" }, { sourceRange: "L7:R7", destinationRange: "L7:R7" }, { sourceRange: "D9:J44", destinationRange: "D9:J44" },    { sourceRange: "L9:R44", destinationRange: "L9:R44" },    { sourceRange: "C48:C61", destinationRange: "C48:C61" },    { sourceRange: "N48:N61", destinationRange: "N48:N61" }  ];
  
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
        const sourceValues = sourceRange.getValues();
        
        let destinationRange = destinationSheet.getRange(rangesToCopy[j].destinationRange);
        destinationRange.setValues(sourceValues);
      }
    }
  }
}
