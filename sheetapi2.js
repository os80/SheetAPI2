const Error_sheet_not_found = "Error: sheet not found"
const Error_need_sheet_id = "Error: enter sheet id or name"
const Message_downlod_sheet = "Downloading the sheet..."

/**
 * Opens the spreadsheet with the given ID. 
 * A spreadsheet ID can be extracted from its URL. 
 * For example, the spreadsheet ID in the URL https://docs.google.com/spreadsheets/d/abc1234567/edit#gid=0 is "abc1234567". 
 * If ID is null - opens current spreadsheet.
 * @class
 * @classdesc Opens the spreadsheet with the given ID.
 * @param {string} ss_id - The unique identifier for the spreadsheet.
 * @return {Table}
 */
function GetTable(ss_id){
  return new Table(ss_id)
}

/** 
 * @lends GetTable.prototype 
 */
class Table {

  constructor(ss_id) {

    if (ss_id) {
      this.ss_id = ss_id;
      this.file = SpreadsheetApp.openById(ss_id);
    } else {
      const SS = SpreadsheetApp.getActiveSpreadsheet();
      this.ss_id = SS.getId();
      this.file = SS;
    }

  }

  /**
   * Returns a sheet with the given name or ID. 
   * For example, the sheet ID in the URL https://docs.google.com/spreadsheets/d/abc1234567/edit#gid=894321 is "894321". 
   * @param {number} sheet_id - The unique identifier for the sheet.
   */
  GetSheet(sheet_id) {
    return new Sheet(this.file, this.ss_id, sheet_id);
  }

  /**
   * Create the sheet with specified ID. Returns a sheet.
   * @param {number} sheet_id - The unique identifier for the sheet.
   * @param {string} sheet_name - The name for the sheet.
   * @param {number} sheet_index - The index of the new sheet. If is null - create the last sheet. If is "0" - create the first sheet.
   */
  CreateSheet(sheet_id, sheet_name, sheet_index) {

    if (!sheet_index && sheet_index !== 0) {
      sheet_index = this.file.getSheets().length
    }

    const requests = [{
      'addSheet': {
        'properties': {
          'title': String(sheet_name),
          "index": +sheet_index,
          "sheetId": +sheet_id,
          'gridProperties': {
            'rowCount': 1000,
            'columnCount': 12
          },
          // 'tabColor': {
          //   'red': 1.0,
          //   'green': 0.3,
          //   'blue': 0.4
          // }
        }
      }
    }];

    try {
      Sheets.Spreadsheets.batchUpdate({ 'requests': requests }, this.ss_id);
    } catch (err) {
      return { error: err.message, function_name: "CreateSheet()" };
    }

    SpreadsheetApp.flush();

    return this.file.getSheetByName(sheet_name);
  }

  /**
   * Deletes the sheet
   * @param {number} sheet_id - The unique identifier for the sheet.
   */
  DeleteSheet(sheet_id) {

    const requests = [
      {
        "deleteSheet": {
          "sheetId": +sheet_id
        }
      }
    ];

    try {
      Sheets.Spreadsheets.batchUpdate({ 'requests': requests }, this.ss_id);
    } catch (err) {
      return { error: err.message, function_name: "DeleteSheet()" };
    }

    return true;

  }

  /**
   * Adds or removes the specified list of users from all protected ranges.
   * @param {String[]} gmails - The array with gmails.
   * @param {Boolean}  adding - true - adds. false - removes.
   */
  ChangeEditorsInProtectedRanges(gmails, adding) {

    let requests = [];
    const all_sheets_data = Sheets.Spreadsheets.get(this.ss_id);


    for (let i = 0; i < all_sheets_data.sheets.length; i++) {

      const all_protected_ranges_from_sheet = all_sheets_data?.sheets?.[i]?.protectedRanges;
      if (!all_protected_ranges_from_sheet) continue;

      for (let n = 0; n < all_protected_ranges_from_sheet.length; n++) {

        const sheet_protected_range = all_protected_ranges_from_sheet[n];
        let list_of_editors = sheet_protected_range?.editors?.users || [];

        gmails.forEach(gmail => {
          if (adding && !list_of_editors.includes(gmail)) list_of_editors.push(gmail);
          if (!adding && list_of_editors.includes(gmail)) list_of_editors.splice(list_of_editors.indexOf(gmail), 1);
        })

        requests.push(
          {
            updateProtectedRange: {
              protectedRange: {
                protectedRangeId: sheet_protected_range.protectedRangeId,
                editors: {
                  users: list_of_editors
                }
              },
              fields: "editors"
            }
          }
        )

      }
    }

    Sheets.Spreadsheets.batchUpdate({ "requests": requests }, this.ss_id);


  }

}

class Sheet {

  constructor(SS, ss_id, sheet_id) {

    this.sheet_id = sheet_id;
    this.ss_id = ss_id;

    if (typeof sheet_id == "number") {

      this.sheet = SS.getSheets().filter(t => t.getSheetId() == sheet_id)[0];
      if (!this.sheet) console.log(Error_sheet_not_found)
      this.sheet_name = this.sheet.getName();

    }

    if (typeof sheet_id == "string") {

      this.sheet = SS.getSheetByName(sheet_id)
      if (!this.sheet) console.log(Error_sheet_not_found)
      this.sheet_name = sheet_id;

    }

  }

  get sheet_id() {
    return this._sheet_id;
  }

  set sheet_id(value) {
    if (!value && value !== 0) {
      console.log(Error_need_sheet_id);
      return;
    }
    this._sheet_id = value;
  }

  /**
  * Returns a two-dimensional array of values, indexed by row, then by column
  * @param {number} firstRow - first row. Default = 1
  * @param {number} firstCol - first column. Default = 1
  * @param {number} rows - Row count or "false" for all
  * @param {number} columns - Columns count or "false" for all
  */
  GetValues(firstRow = 1, firstCol = 1, rows = "", columns = 99) {
    return new Values(this.ss_id, this.sheet_id, this.sheet_name, firstRow, firstCol, rows, columns)
  }

  /**
  * Sets a rectangular grid of values.
  * @param {number} output_arr - 2D array
  * @param {number} firstRow - first row. Default = 1
  * @param {number} firstCol - first column. Default = 1
  */
  SetValues(output_arr, firstRow = 1, firstCol = 1) {

    const range = this.sheet_name + "!" + this.columnToLetter_(firstCol) + firstRow;

    Sheets.Spreadsheets.Values.update(
      {
        majorDimension: "ROWS", // ROWS, COLUMNS
        values: output_arr
      },
      this.ss_id,
      range,
      { valueInputOption: "USER_ENTERED" } // USER_ENTERED, RAW
    )


  }

  /**
  * Clears the range of contents and formats.
  * @param {number} firstRow - first row. Default = 1
  * @param {number} firstCol - first column. Default = 1
  * @param {number} rows - Row count or "false" for all
  * @param {number} columns - Columns count or "false" for all 
  */
  Clear(firstRow = 1, firstCol = 1, rows = "", columns = 999) {

    const range =
      this.sheet_name + "!" +
      this.columnToLetter_(firstCol) +
      firstRow + ":" +
      this.columnToLetter_(firstCol + columns - 1) +
      (rows ? firstRow + rows - 1 : "");

    Sheets.Spreadsheets.Values.clear('', this.ss_id, range);

  }

  /**
  * Clears formatting for this range.
  * This clears text formatting for the cell or cells in the range, but does not reset any number formatting rules.
  * @param {number} firstRow - first row. Default = 1
  * @param {number} firstCol - first column. Default = 1
  * @param {number} rows - Row count or "false" for all
  * @param {number} columns - Columns count or "false" for all 
  */
  ClearContent(firstRow = 1, firstCol = 1, rows = 0, columns = 99) {

    if (!rows)
      rows = this.sheet.getLastRow();

    const range = {
      sheetId: this.sheet_id,
      startRowIndex: firstRow - 1,
      endRowIndex: firstRow + rows - 1,
      startColumnIndex: firstCol - 1,
      endColumnIndex: firstCol + columns - 1
    }

    const requests = {
      updateCells: {
        rows: [],
        fields: "userEnteredValue",
        range: range,
      }
    }

    Sheets.Spreadsheets.batchUpdate({ "requests": requests }, this.ss_id);

  }

  /**
  * Delete duplicates for this range.
  * @param {number} firstRow - first row. Default = 1
  * @param {number} firstCol - first column. Default = 1
  * @param {number} rows - Row count or "false" for all
  * @param {number} columns - Columns count or "false" for all 
  */
  DeleteDuplicates(firstRow = 1, firstCol = 1, rows, columns) {

    if (!rows)
      rows = this.sheet.getMaxRows() - firstRow;

    if (!columns)
      columns = this.sheet.getMaxColumns() - firstCol;

    let sheet_id = this.sheet_id;

    if (typeof sheet_id == "string")
      sheet_id = this.sheet.getSheetId();


    const requests = {
      deleteDuplicates: {
        comparisonColumns: {
          sheetId: sheet_id,
          dimension: "COLUMNS",
          startIndex: firstCol - 1,
          endIndex: columns + firstCol
        },
        range: {
          sheetId: sheet_id,
          startRowIndex: firstRow - 1,
          endRowIndex: rows,
          startColumnIndex: firstCol - 1,
          endColumnIndex: columns + firstCol
        }
      }
    }

    return Sheets.Spreadsheets.batchUpdate({ "requests": requests }, this.ss_id);

  }

  /**
   * Deletes rows.
   * @param {array} rows_arr - array of indexes
   */
  DeleteRows(rows_arr) {

    let requests = [];

    for (let i = rows_arr.length - 1; i >= 0; i--)
      requests.push({
        "deleteDimension": {
          "range": {
            "sheetId": this.sheet_id,
            "dimension": "ROWS",
            "startIndex": rows_arr[i] - 1,
            "endIndex": rows_arr[i]
          }
        }
      });

    return Sheets.Spreadsheets.batchUpdate({ requests: requests }, this.ss_id)

  }

  /**
   * Creates a new file in the folder in the selected format from the specified sheet. Returns the created file
   * Formats: pdf, xlsx, ods, zip, csv
   * @param {string} name - File name
   * @param {object} target_folder - Target folder
   * @param {string} format - File format. Default = "pdf"
   */
  ConvertSheet(name, target_folder, format = "pdf") {

    const fr = 0, fc = 0, lc = 14, lr = 54;
    const url = "https://docs.google.com/spreadsheets/d/" + this.ss_id + "/export" +
      "?format=" + format + "&" +
      "size=7&" +
      "fzr=true&" +
      "portrait=true&" +
      "fitw=true&" +
      "gridlines=false&" +
      "printtitle=false&" +
      "top_margin=0.5&" +
      "bottom_margin=0.25&" +
      "left_margin=0.5&" +
      "right_margin=0.5&" +
      "sheetnames=false&" +
      "pagenum=UNDEFINED&" +
      "attachment=true&" +
      "gid=" + this.sheet_id + '&' +
      "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

    const params = {
      method: "GET",
      headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true,
    };

    const blob = UrlFetchApp.fetch(url, params).getBlob().setName(`${name}.${format}`);
    return target_folder.createFile(blob);

  }

  /**
   * Download the specified sheet in the selected format.
   * Formats: pdf, xlsx, ods, zip, csv
   * @param {string} format - File format. Default = "pdf"
   */
  ConvertSheetAndDownload(format = "pdf") {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ss_id = ss.getId();
    var sheet_id = ss.getActiveSheet().getSheetId();
    var urlToOpen = "https://docs.google.com/spreadsheets/d/" + ss_id + "/export?format=" + format + "&gid=" + sheet_id;

    var html = "<script>window.open('" + urlToOpen + "');google.script.host.close();</script>";
    var userInterface = HtmlService.createHtmlOutput(html).setHeight(100);
    SpreadsheetApp.getUi().showModalDialog(userInterface, Message_downlod_sheet);

  }

  columnToLetter_(column) {
    let temp, letter = '';
    while (column > 0) {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }
    return letter;
  }

}

class Values {

  constructor(ss_id, sheet_id, sheet_name, firstRow = 1, firstCol = 1, rows = "", columns = 99) {

    this.sheet_id = sheet_id;
    this.ss_id = ss_id;
    this.sheet_name = sheet_name;

    this._firstRow = firstRow;

    const range =
      this.sheet_name + "!" +
      this.columnToLetter_(firstCol) +
      firstRow + ":" +
      this.columnToLetter_((firstCol + columns - 1)) +
      (rows ? firstRow + rows - 1 : "");

    const options = {
      valueRenderOption: 'UNFORMATTED_VALUE',  // FORMATTED_VALUE, UNFORMATTED_VALUE, FORMULA
      dateTimeRenderOption: 'FORMATTED_STRING',   // SERIAL_NUMBER, FORMATTED_STRING
      majorDimension: 'ROWS',               // COLUMNS, ROWS
    }

    this.values = Sheets.Spreadsheets.Values.get(this.ss_id, range, options).values


  }

  /**
   * Returns the data that a basic filter displays in a sheet.
   */
  GetValuesFromBasicFilter() {

    const fields = 'sheets(data(rowMetadata(hiddenByFilter,hiddenByUser)),properties(sheetId,title))';
    const sheets = Sheets.Spreadsheets.get(this.ss_id, { fields }).sheets


    let sheet;
    if (typeof this.sheet_id == "number")
      sheet = sheets.filter(({ properties }) => {
        return String(properties.sheetId) === String(this.sheet_id);
      })[0];

    if (typeof this.sheet_id == "string")
      sheet = sheets.filter(({ properties }) => {
        return String(properties.title) === this.sheet_id;
      })[0];

    const rowMetadata = sheet.data[0].rowMetadata.filter((_, index) => index >= this._firstRow - 1)

    return this.values.filter((_, index) => !rowMetadata[index].hiddenByFilter)

  }

  columnToLetter_(column) {
    let temp, letter = '';
    while (column > 0) {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }
    return letter;
  }

}
