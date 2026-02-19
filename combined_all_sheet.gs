function combineSheetsInBatches() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  const batchSize = 5;   // Change if needed
  const startIndex = 0;   // Change to 20, 40, 60 for next batches

  const newSS = SpreadsheetApp.create("Combined_Batch_Output");
  const combinedSheet = newSS.getActiveSheet();
  
  let writeRow = 1;

  for (let i = startIndex; i < startIndex + batchSize && i < sheets.length; i++) {
    const sheet = sheets[i];
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow > 0 && lastCol > 0) {
      const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      
      combinedSheet
        .getRange(writeRow, 1, data.length, data[0].length)
        .setValues(data);

      writeRow += data.length;
      SpreadsheetApp.flush();
    }
  }

  Logger.log("Batch file URL: " + newSS.getUrl());
}
