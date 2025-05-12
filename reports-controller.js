function ReportsController() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setupSheet = ss.getSheetByName("Setup");

  if (!setupSheet) {
    Logger.log("Setup sheet not found!");
    return;
  };

  const data = setupSheet.getDataRange().getValues(); // Read all data
  const forms = [];

  let order = 0;

  for (let i = 1; i < data.length; i++) {
    const strand = data[i][0]; // Column A (Name)
    if (!strand) continue;

    for (let j = 1; j < data[i].length; j++) { // Start from Column B
      const url = data[i][j];

      if (url 
        && typeof url === "string" 
        && url.startsWith("http")
      ) {
        forms.push({
          order: order,
          url: url.trim(),
          strand: strand,
        });

        order++;
      };
    };
  };

  if (forms.length > 0) {
    const report = buildReports(forms);
    fillReportsSheet(report);
  };
};