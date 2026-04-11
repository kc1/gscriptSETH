function test3() {

  const id = "17E6Q7nnglX0m8u1Z9xFCWVheBgO-Vy44yWzgYKKNYIrYEL76_co9_T4E";

  // var id = '1iNkVRNexbuxqXJHSbz5kW3DXb4Igv13IdOG48Qzm-uQ'

  var sheet = SpreadsheetApp.openById(id).getSheetByName(' ');

  //  var selection = SpreadsheetApp.getActiveSpreadsheet().getSelection();
  //  var activeRange = selection.getActiveRange().getValues();

  //  Logger.log(activeRange);

  var rows = sheet2Json(sheet)
  console.log(rows)

}

function getRowCellRangeByColumnName(rowJSON, header) {
  Logger.log(rowJSON);
  Logger.log(header);
  // const rowObj = JSON.parse(rowJSON);
  const keys = Object.keys(rowJSON);
  Logger.log(keys);
  const i = keys.indexOf(header);
  Logger.log(i);
  const columnLetter = columnNumberToLetter(i + 1);
  // I'M just using E for now - you should be able to calc column with columnNumberToLetter
  Logger.log(columnLetter);

  const cellRange = columnLetter + rowJSON.absoluteRow;
  return cellRange;

}

function columnNumberToLetter(columnNumber) {
  var columnLetter = '';
  while (columnNumber > 0) {
    var remainder = (columnNumber - 1) % 26;
    columnLetter = String.fromCharCode(65 + remainder) + columnLetter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return columnLetter;
}


function getSelectedRowObject() {

  // NOTE - ONLY WORKS FROM SHEET, NOT THIS EDITOR !!

  let activeSheet = SpreadsheetApp.getActiveSheet();
  Logger.log(activeSheet.getName());

  var selection = activeSheet.getSelection();
  var activerow = selection.getCurrentCell().getRow();
  Logger.log(activerow)
  var sheet = SpreadsheetApp.getActiveSheet();

  const json = sheet2Json(sheet);
  return json[activerow - 2];

}

function getSelectedRowsRange() {

  // I think you have to select use SHIFT to select

  // Get the active spreadsheet and sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Get the range of selected cells
  const selectedRange = sheet.getActiveRange();

  if (!selectedRange) {
    Logger.log('No range selected');
    return;
  }

  // Get the start and end rows of the selection
  const startRow = selectedRange.getRow();
  const numRows = selectedRange.getNumRows();
  const endRow = startRow + numRows - 1;

  Logger.log(`Selected range starts at row ${startRow} and ends at row ${endRow}`);

  // Return the range object
  // return sheet.getRange(startRow, 1, numRows, sheet.getMaxColumns());
  return { "startRow": startRow, "endRow": endRow }
}


function getMultipleSelectedRowObjects() {

  const rangeObj = getSelectedRowsRange()
  var sheet = SpreadsheetApp.getActiveSheet();
  const json = sheet2Json(sheet);
  console.log(json[0])
  const start = rangeObj.startRow - 2;
  const end = rangeObj.endRow - 1;
  let selectedRowObjectArray = []
  for (var i = start; i <= end-1; i++) {
    const obj = json[i]
    selectedRowObjectArray.push(obj)
  }

  // let selectedRowObjectArray = json.slice(rangeObj.startRow - 2, rangeObj.endRow - 1); // includes start and end row
  console.log(selectedRowObjectArray);

  return selectedRowObjectArray;

}

function getMultipleSelectedRowObjectsDiscontinuous() {
  const selectedRowNumbers = getSelectedRowsDiscontinuous();
  if (selectedRowNumbers.length === 0) return [];

  const sheet = SpreadsheetApp.getActiveSheet();
  const allRowsAsObjects = sheet2Json(sheet);

  const selectedObjects = selectedRowNumbers
    .map(rowNum => {
      // Map each row number to its corresponding object
      const jsonIndex = rowNum - 2;
      return (jsonIndex >= 0 && jsonIndex < allRowsAsObjects.length) 
        ? allRowsAsObjects[jsonIndex] 
        : null; // Return null for invalid rows
    })
    .filter(obj => obj !== null); // Filter out the nulls

  console.log(selectedObjects);
  return selectedObjects;
}


/**
 * Gets all unique row numbers from one or more discontinuous selections in the active sheet.
 * To use, select multiple rows/cells using Ctrl-Click (Windows/Linux) or Cmd-Click (Mac).
 *
 * @returns {number[]} An array of unique, sorted row numbers that were selected. Returns an empty array if nothing is selected.
 */
function getSelectedRowsDiscontinuous() {
  // Get the active spreadsheet and sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Use getActiveRangeList() to handle multiple, non-contiguous selections
  const rangeList = sheet.getActiveRangeList();

  if (!rangeList) {
    Logger.log('No range was selected.');
    // Optional: Show an alert to the user
    SpreadsheetApp.getUi().alert('Please select one or more cell ranges first.');
    return [];
  }

  // Get an array of all the individual Range objects
  const ranges = rangeList.getRanges();
  
  if (ranges.length === 0) {
    Logger.log('Selection is empty.');
    return [];
  }

  // Use a Set to automatically handle duplicate row numbers
  // This is useful if the user selects, for example, A5 and C5. The row '5' will only be added once.
  const selectedRowsSet = new Set();

  // Loop through each individual selected range
  for (const range of ranges) {
    const startRow = range.getRow();
    const numRows = range.getNumRows();
    const endRow = startRow + numRows - 1;

    // Add every row number in the current range to our set
    for (let i = startRow; i <= endRow; i++) {
      selectedRowsSet.add(i);
    }
  }

  // Convert the Set to an Array and sort it numerically
  const uniqueSortedRows = Array.from(selectedRowsSet).sort((a, b) => a - b);

  if (uniqueSortedRows.length > 0) {
    Logger.log(`Selected rows are: ${uniqueSortedRows.join(', ')}`);
  } else {
    Logger.log('No rows were identified in the selection.');
  }

  return uniqueSortedRows;
}


/**
 * A simple test function to demonstrate how to use getSelectedRowsDiscontinuous().
 * 1. Open your spreadsheet.
 * 2. Hold Ctrl (or Cmd) and click on a few different rows (e.g., row 2, row 5, and drag to select rows 9 and 10).
 * 3. Run this test function from the script editor.
 * 4. Check the logs (View > Logs).
 */
function testSelection() {
  const selectedRows = getSelectedRowsDiscontinuous();
  
  if (selectedRows.length > 0) {
    SpreadsheetApp.getUi().alert(`You selected the following rows: ${selectedRows.join(', ')}`);
  }
}



function sheet2Json(sheet) { // USE THIS RATHER THAN OTHERS - BETTER JSON PRODUCED


  var HeaderRow = sheet.getFrozenRows(); // THIS SHOULD GET HEADER ROW 

  //  var colStartIndex = 1;
  //  var rowNum = 1;
  //  var firstRange = sheet.getRange(HeaderRow, 1, 1, sheet.getLastColumn()); 
  //  var firstRowValues = firstRange.getValues();
  //  var titleColumns = firstRowValues[0];

  var titleColumns = sheet.getDataRange().offset(HeaderRow - 1, 0, 1).getValues()[0];
  //        Logger.log(headings);
  //      Logger.log(titleColumns);


  //      x=y;

  //  return;

  // after the second line(data)
  var lastRow = sheet.getLastRow();
  var rowValues = [];
  //  for(var rowIndex=1; rowIndex<=lastRow; rowIndex++) {
  //      for(var rowIndex=1; rowIndex<=50; rowIndex++) {

  var colStartIndex = 1;
  var rowNum = 1;
  //    var range = sheet.getRange(rowIndex, colStartIndex, rowNum, sheet.getLastColumn());
  //  var range = sheet.getRange(HeaderRow, 1, sheet.getLastRow(), sheet.getLastColumn());
  var range = sheet.getRange(HeaderRow + 1, 1, sheet.getLastRow(), sheet.getLastColumn());
  var rowValues = range.getValues(); // this is a 0 0 based array
  //    rowValues.push(values[0]);
  //        Logger.log(values);

  //  Logger.log('jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj');

  //    Logger.log(rowValues);

  //    var backgrounds = range.getBackgrounds();
  //        Logger.log(backgrounds);

  //  return;



  // create json
  var jsonArray = [];
  for (var i = 0; i < rowValues.length; i++) { //THE FIRST ROW AFTER HEADER ROW - eg main first row data = row 4 
    var line = rowValues[i];
    var json = new Object();
    for (var j = 0; j < titleColumns.length; j++) {
      json[titleColumns[j]] = line[j];
    }
    json['relativeRow'] = i + 1; //DATA AFTER HEADER ROW
    json['absoluteRow'] = i + HeaderRow + 1; // add one to convert to 1 based array for cellvalue

    //        json['colour'] = backgrounds[i +HeaderRow+1]

    jsonArray.push(json);
  }

  //  Logger.log(jsonArray);
  return jsonArray;

}



function updateCell(sheet, singleRowJson, columnToUpdate, value) { // THIS HAS TO BE FIXED USING ID AND gmail_list NOT MAIN 

  var headerRow = sheet.getFrozenRows(); // THIS SHOULD GET HEADER ROW 
  var headings = sheet.getDataRange().offset(headerRow - 1, 0, 1).getValues()[0];
  var columnNumber = headings.indexOf(columnToUpdate);
  //  var absoluteRow = Number(singleRowJson['relativeRow']) + Number(headerRow) +1; // THIS CONTAINS relativeRow IN THE json from convertSheet2JsonByNameAndHeaderRow
  var absoluteRow = Number(singleRowJson['absoluteRow']); // THIS CONTAINS relativeRow IN THE json from convertSheet2JsonByNameAndHeaderRow


  var cell = sheet.getRange(absoluteRow, columnNumber + 1);

  cell.setValue(value);

  return singleRowJson;

}

function cellTest() {

  var id = '1L_U7MhmVop5NWpdVYpeOmRn6lVk-YeQmVjigkqlDoUg'

  var sheet = SpreadsheetApp.openById(id).getSheetByName('Main');


  var cell = sheet.getRange(5, 1); //affects row 5 column 1

  cell.setValue('hello');

}


function jsonRowsFromSheet(Name) {

  var id = '1UOBRLtYvDnYrRAfDzUvgxvYX4lBubAxHcKgPXcx82Y0';

  var sheet = SpreadsheetApp.openById(id).getSheetByName(Name);
  var headerRow = sheet.getFrozenRows();
  var headings = sheet.getDataRange().offset(headerRow - 1, 0, 1).getValues()[0];
  Logger.log('jsonRowsFromSheet');


  return sheet2Json(sheet);
}

function addRowFromObject(myObj, sheet) {

  Logger.log('in addRowFromObject');
  Logger.log(myObj);

  // var headerRow = sheet.getFrozenRows(); // THIS SHOULD GET HEADER ROW 
  var headerRow = sheet.getFrozenRows() + sheet.getRange("A1").getRow() - 1;

  var headings = sheet.getDataRange().offset(headerRow - 1, 0, 1).getValues()[0];
  var newRow = [];

  headings.forEach(function (heading) {

    if (myObj.hasOwnProperty(heading)) {
      newRow.push(myObj[heading])
    } else {
      newRow.push('');
    }


  });


  Logger.log(newRow);

  sheet.appendRow(newRow);





  return true


}


function updateRowWithObject(Name) {


}


function getRowFromSheetByColumnValue(sheetName, column, value) {

  var rows = jsonRowsFromSheet(sheetName);
  return rows.filter(row => row[column] == value)[0]

}