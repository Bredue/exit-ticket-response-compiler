function MoonPhasesController() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Google Sheet Name in Google Sheet that App Script is running in");
    const forms = importMoonPhasesResponses();
    const responses = compileFormResponses(forms);
    fillGoogleSheet(sheet, forms, responses);
  }
  
  function importMoonPhasesResponses() {  
    return [
      {order: 0, url: "Google Form Edit Link", startColumn: 1},  // Columns A-D
      {order: 1, url: "Google Form Edit Link", startColumn: 7},  // Columns G-J
      {order: 2, url: "Google Form Edit Link", startColumn: 13}, // Columns M-P
      {order: 3, url: "Google Form Edit Link", startColumn: 19}, // Columns S-V
    ];
  }