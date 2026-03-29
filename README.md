# Google Apps Script: Combine Sheets in Batches

## 📌 Overview
The function `combineSheetsInBatches()` is used to merge data from multiple sheets inside a Google Spreadsheet into a new spreadsheet.

Instead of combining all sheets at once (which may hit limits), it processes sheets in **small batches**.

---

## ⚙️ Function Code

```javascript
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
```

---

## 🔍 How It Works

### 1. Get Active Spreadsheet
```javascript
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheets = ss.getSheets();
```
- Fetches the current spreadsheet
- Gets all sheets inside it

---

### 2. Configure Batch Settings
```javascript
const batchSize = 5;
const startIndex = 0;
```
- `batchSize`: Number of sheets to process at a time  
- `startIndex`: Starting sheet index  
- Example:
  - `0` → first 5 sheets
  - `5` → next 5 sheets
  - `10` → next batch

---

### 3. Create New Spreadsheet
```javascript
const newSS = SpreadsheetApp.create("Combined_Batch_Output");
const combinedSheet = newSS.getActiveSheet();
```
- Creates a new spreadsheet
- Uses its first sheet for combined data

---

### 4. Loop Through Sheets
```javascript
for (let i = startIndex; i < startIndex + batchSize && i < sheets.length; i++)
```
- Loops only through a limited number of sheets
- Prevents exceeding Google Apps Script limits

---

### 5. Read Data from Each Sheet
```javascript
const lastRow = sheet.getLastRow();
const lastCol = sheet.getLastColumn();
```
- Determines data boundaries

```javascript
const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
```
- Reads all data from the sheet

---

### 6. Write Data to Combined Sheet
```javascript
combinedSheet
  .getRange(writeRow, 1, data.length, data[0].length)
  .setValues(data);
```
- Writes data into the new spreadsheet
- Starts writing from `writeRow`

```javascript
writeRow += data.length;
```
- Moves pointer down for next dataset

---

### 7. Force Immediate Write
```javascript
SpreadsheetApp.flush();
```
- Ensures data is written instantly
- Helps avoid execution issues

---

### 8. Output Result
```javascript
Logger.log("Batch file URL: " + newSS.getUrl());
```
- Logs the URL of the newly created spreadsheet

---

## 🚀 How to Use

1. Open Google Apps Script
2. Paste the function
3. Run `combineSheetsInBatches()`
4. Check **Logs** (`View → Logs`) for output file link

---

## 🔁 Running Multiple Batches

To process all sheets:

| Batch | startIndex | batchSize |
|------|------------|----------|
| 1    | 0          | 5        |
| 2    | 5          | 5        |
| 3    | 10         | 5        |

Update:
```javascript
const startIndex = 5;
```

---

## ⚠️ Notes

- Works best when all sheets have similar structure
- Large datasets may hit execution limits → batching solves this
- Empty sheets are skipped automatically

---

## 💡 Use Cases

- Merging monthly reports
- Combining department data
- Exporting multiple sheets into one file

---

## ✅ Summary

This script:
- Reads multiple sheets
- Combines them into one spreadsheet
- Processes data safely using batching

---
