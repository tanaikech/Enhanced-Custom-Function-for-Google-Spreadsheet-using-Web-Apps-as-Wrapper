const key = "samplekey"; // This is a key for using Web Apps. You can freely modify this.

// Sample function 1.
// This sample script returns the filenames in the folder by giving the folder name.
function getFileNamesFromFolderName(folderName) {
  const files = DriveApp.getFoldersByName(folderName).next().getFiles();
  let ar = [];
  while (files.hasNext()) ar.push(files.next().getName());
  return ar;
}

// Sample function 2.
// This sample script put "values" to "Sheet1" in the active Spreadsheet.
// Sheets API is used.
function putValues(values) {
  if (!Array.isArray(values)) values = [values];
  const id = SpreadsheetApp.getActiveSpreadsheet().getId();
  Sheets.Spreadsheets.Values.append({ values: [values] }, id, "Sheet1", {
    valueInputOption: "USER_ENTERED",
  });
  return "Done";
}

// Sample function 3.
// This sample script set the background colors with the gradation colors.
// Sheets API is used.
function setColors(values) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const step1 = 1 / values[2];
  const step2 = 1 / values[3];
  let rows = [];
  for (let i = 0; i <= 1; i += step1) {
    let cols = [];
    for (let j = 0; j <= 1; j += step2) {
      cols.push({
        userEnteredFormat: { backgroundColor: { red: 1, green: i, blue: j } },
      });
    }
    rows.push({ values: cols });
  }
  const resource = {
    requests: [
      {
        updateCells: {
          range: {
            sheetId: 0,
            startRowIndex: values[0] - 1,
            endRowIndex: values[0] + values[2],
            startColumnIndex: values[1] - 1,
            endColumnIndex: values[1] + values[3],
          },
          rows: rows,
          fields: "userEnteredFormat.backgroundColor",
        },
      },
    ],
  };
  Sheets.Spreadsheets.batchUpdate(resource, ss.getId());
  return "Done";
}

//
// The following script is the script for the enhanced custom function for Google Spreadsheet using Web Apps as the wrapper.
//
// Web Apps using as the wrapper for authorizing.
function doGet(e) {
  let res = "";
  if (e.parameter.key === key) {
    try {
      res = this[e.parameter.name](
        e.parameter.args.includes(",")
          ? e.parameter.args.split(",")
          : e.parameter.args
      );
    } catch (err) {
      res = `Error: ${err.message}`;
    }
  } else {
    res = "Key error.";
  }
  return ContentService.createTextOutput(JSON.stringify({ value: res }));
}

/**
 * Run GAS function.
 * @param {"functionName"} functionName Function name you want to run in this container-bound script.
 * @param {"arg1", "arg2",,,} ...args Arguments for the function.
 * @return Returned values from the function.
 * @customfunction
 */
function RUN(functionName, ...args) {
  const webAppsUrl = "https://script.google.com/macros/s/###/exec"; // Please set the URL of Web Apps after you set the Web Apps.

  if (!functionName) throw new Error("No function name.");
  const url = `${webAppsUrl}?name=${functionName}&args=${args}&key=${key}`;
  const res = UrlFetchApp.fetch(url);
  if (res.getResponseCode() != 200) throw new Error(res.getContentText());
  return JSON.parse(res.getContentText()).value;
}
