function createOpenerFunnels() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();
  var openersMap = {}; 

  // Step 1: Map out all the data by Opener and by Sheet Name
  allSheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    var data = sheet.getDataRange().getValues();
    
    // Skip sheets with no data
    if (data.length < 2) return; 
    
    var headers = data[0];
    var openerColIndex = headers.indexOf("Opener");
    
    // Only process tabs that actually have an "Opener" column
    if (openerColIndex === -1) return; 
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var opener = row[openerColIndex];
      
      if (opener) {
        opener = opener.toString().trim();
        
        // Skip empty cells
        if (opener === "") continue; 
        
        if (!openersMap[opener]) {
          openersMap[opener] = {};
        }
        
        // Initialize the sheet in the opener's map with the header row
        if (!openersMap[opener][sheetName]) {
          openersMap[opener][sheetName] = [headers];
        }
        
        openersMap[opener][sheetName].push(row);
      }
    }
  });

  // Step 2: Create a new Spreadsheet for each Opener
  for (var opener in openersMap) {
    // Creates a new file like "Ben BD Meetings"
    var newSs = SpreadsheetApp.create(opener + " BD Meetings");
    var sheetDataMap = openersMap[opener];
    var isFirstSheet = true;
    
    for (var sheetName in sheetDataMap) {
      var targetSheet;
      if (isFirstSheet) {
        // Use the default "Sheet1" for the first tab and rename it
        targetSheet = newSs.getActiveSheet();
        targetSheet.setName(sheetName);
        isFirstSheet = false;
      } else {
        // Create new tabs for subsequent funnel stages
        targetSheet = newSs.insertSheet(sheetName);
      }
      
      var exportData = sheetDataMap[sheetName];
      // Paste the filtered data
      targetSheet.getRange(1, 1, exportData.length, exportData[0].length).setValues(exportData);
    }
  }
  
  SpreadsheetApp.getUi().alert("Funnel extraction complete! Check your Google Drive root folder for the new spreadsheets.");
}