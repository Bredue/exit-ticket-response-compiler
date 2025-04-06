function ExitTicketController() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const setupSheet = ss.getSheetByName("Setup");
  
    if (!setupSheet) {
      Logger.log("Setup sheet not found!");
      return;
    };
  
    const data = setupSheet.getDataRange().getValues(); // Read all data
  
    for (let i = 1; i < data.length; i++) {
      const name = data[i][0]; // Column A (Name)
      if (!name) continue;
  
      let sheet = ss.getSheetByName(name);
      if (!sheet) {
        sheet = ss.insertSheet(name, ss.getNumSheets()); // Create a new sheet at the end
      };
  
      const forms = [];
      let startColumn = 1;
      let order = 0;
  
      for (let j = 1; j < data[i].length; j++) { // Start from Column B
        const url = data[i][j];
  
        if (url 
          && typeof url === "string" 
          && url.startsWith("http")
        ) {
          forms.push({
            order: order,
            url: url.trim(),
            startColumn: startColumn
          });
  
          order++;
          startColumn += 6;
        }
      }
  
      if (forms.length > 0) {
        const allFormData = compileFormResponses(forms);
        fillGoogleSheet(sheet, forms, allFormData);
      };
    };
  };  