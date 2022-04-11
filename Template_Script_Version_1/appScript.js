// @ts-nocheck
/*

Goals met:
 Checks sheet cells for whiteSpace, removes it from main sheet, highlights report sheet cell light green, inserts Comment 'WhiteSpace Removed' within report sheet cell
 Check for duplicate emails, highlights main sheet and report sheet cells, inserts comment 'Duplicate Email' within report sheet cell, and sorts both main and report sheets by email
 When inserting comments to cells, insertCommentToSheetCell function checks to see if comment exisits, and if so, concatenates new comment instead of overriding the old comment(s)


Problem:
  Automate the checking of the User Import Template

  Explicit Requirements:
    Flag invalid emails (Add new sheet reporting what error were flagged and in which cells as well as what information was changed in which cells)
    Turn these cells red:
        invalid emails
        missing emails
        missing first names
        missing last names
        phone numbers with an invalid number of digits
        zip codes with an invalid number of digits
    Change full state names to theri two-letter codes
    Change full County names to their two-letter code
    Change dates to the YYYY-MM-DD format (within "User Date Added" and "Birthday" columns)
    Change first and last names to first letter capitilized, rest of letters undercase
    Remove Whitespace from the beginning and end of cell entires
    Flag links that do no match their text (for exmaple "Click here")
    Flag special characters as orange for review
    Make sure to create errors for throwing when the script has an issue

     
 
 */

/*

Email Validation:
const validateEmail = (email) => {
  return String(email)
    .toLowerCase()
    .match(
      /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/
    );
};
 */

// function trimEachCell(sheet) {
//   let range = sheet.getDataRange();

// }
const REPORT_SHEET_NAME = 'Report';
const WHITE_SPACE_REMOVED_COMMENT = 'Whitespace Removed';
const DUPLICATE_EMAIL_COMMENT = 'Duplicate email';
const LIGHT_GREEN_HEX_CODE = '#b6d7a8';
const LIGHT_RED_HEX_CODE = '#f4cccc';
let sheet = SpreadsheetApp.getActiveSheet();
let ss = SpreadsheetApp.getActiveSpreadsheet();
let data = sheet.getDataRange();
let values = data.getValues();
let columnHeaders = values[0];
//Consider values[0].map(header => (header[0] + header.substr(1)).trim()) to help avoid errors here

function getSheet(sheetName) {
  return ss.getSheetByName(sheetName);
}

function setFrozenRows(sheetBinding, numOfRowsToFreeze) {
  sheetBinding.setFrozenRows(numOfRowsToFreeze);
}

function createSheetCopy(spreadSheet, newSheetName, valuesToCopyOver) {
  let newSheet = spreadSheet.insertSheet();
  newSheet.setName(newSheetName);
  valuesToCopyOver.forEach(row => newSheet.appendRow(row));

  return newSheet;
}

  function getSheetCell(sheetBinding, row, column) {
  return sheetBinding.getRange(row, column);
}

function setSheetCellBackground(sheetCellBinding, color) {
  sheetCellBinding.setBackground(color);
}

function checkCellForComment(sheetCell) {
  return sheetCell.getComment() ? true: false;
}

//Consider searching for if the comment already exists and not adding a new one in that case
function insertCommentToSheetCell(sheetCell, comment) {
  if (checkCellForComment(sheetCell)) {
    let previousComment = sheetCell.getComment();
    let newComment = `${previousComment}, ${comment}`;
    sheetCell.setComment(newComment);
  } else {
    sheetCell.setComment(comment);
  }

}

function deleteSheetIfExists(sheetBinding) {
  if (sheetBinding !== null) {
   ss.deleteSheet(sheetBinding);
  }
}

function getColumnRange(columnName, sheetName) {
  let column = columnHeaders.indexOf(columnName);
  if (column !== -1) {
    return sheetName.getRange(2,column+1,sheetName.getMaxRows());
  }

  return null;
}

function getValues(range) {
  return range.getValues();
}

function removeWhiteSpaceFromCells(mainSheetBinding, valuesToCheck, reportSheetBinding) {
  let rowCount = valuesToCheck.length;
  let columnCount = valuesToCheck[0].length;
  let searchRange = (1, 1, rowCount, columnCount);

  let cellsTrimmed = [];

    for (let row = 1; row < rowCount; row += 1) {
    for (let column = 0; column < columnCount; column += 1) {
      let currentVal = valuesToCheck[row][column];
      let firstChar = currentVal.toString()[0];
      let lastChar = currentVal[currentVal.toString().length - 1];
      let currentCell = getSheetCell(mainSheetBinding, row + 1, column + 1);
      let reportSheetCell = getSheetCell(reportSheetBinding, row + 1, column + 1);

      if (firstChar === " " || lastChar === " ") {
      currentCell.setValue(`${currentVal.trim()}`);
      setSheetCellBackground(reportSheetCell, LIGHT_GREEN_HEX_CODE);
      insertCommentToSheetCell(reportSheetCell, WHITE_SPACE_REMOVED_COMMENT);

      //Remove later or check row and column information to be in A1 Notations

      cellsTrimmed.push(`Value: ${currentVal} Row: ${currentCell.getRow()} Column: ${currentCell.getColumn()}`);
      }
    }
  }
  //Remove later or use for logging events
  console.log(cellsTrimmed);
}


function checkForDuplicateEmails(sheetBinding, reportSheetBinding) {
  let emailColumnRange = getColumnRange('Email', sheetBinding);
  let emailColumnValues = getValues(emailColumnRange).map(email => email[0].toLowerCase().trim());
  let emailColumnPosition = emailColumnRange.getColumn();

  let duplicates = [];

  emailColumnValues.forEach((email, index) => {
    let currentEmail = String(email);
    let row = index + 2;

      if (duplicates.indexOf(currentEmail) === -1 && currentEmail.length > 0) {
       if (emailColumnValues.filter(val => String(val) === currentEmail).length > 1) {
        let currentCell = getSheetCell(sheetBinding, row, emailColumnPosition);
        let reportSheetCell = getSheetCell(reportSheetBinding, row, emailColumnPosition);

        //Line below for testing

        SpreadsheetApp.getUi().alert(`Duplicate found! ${currentEmail} ${reportSheetCell.getA1Notation()}`);
        //testing line ends here
        setSheetCellBackground(reportSheetCell, LIGHT_RED_HEX_CODE);
        setSheetCellBackground(currentCell, LIGHT_RED_HEX_CODE);
        insertCommentToSheetCell(reportSheetCell, DUPLICATE_EMAIL_COMMENT);

        duplicates.push(currentEmail);
        }
      }
      }
    );

  reportSheetBinding.sort(emailColumnPosition);
  sheetBinding.sort(emailColumnPosition)
}

/*

https://developers.google.com/apps-script/reference/spreadsheet/conditional-format-rule-builder?hl=en
let sheet = SpreadsheetApp.getActiveSheet();
let range = sheet.getRange("A1:B3");
let rule = sheet.newConditionalFormatRule()
  .whenFormulaSatisfied("=COUNTIF(2:C323,C2)>1")
  .setBackground("#FF0000")
  .setRanges([range])
  .build();
let rules = sheet.getConditionalFormatRules();
rules.push(rule);
sheet.setConditionalFormatRules(rules);
*/

// Function that executes when user clicks "User Template Check" Button

function checkUserImportTemplate() {
let reportSheet = getSheet(REPORT_SHEET_NAME);

deleteSheetIfExists(reportSheet);


reportSheet = createSheetCopy(ss, REPORT_SHEET_NAME, values);

setFrozenRows(sheet, 1);
setFrozenRows(reportSheet, 1);

removeWhiteSpaceFromCells(sheet, values, reportSheet);
checkForDuplicateEmails(sheet, reportSheet);


}

  
// Creates "User Template Check" navigation button within Spreadsheet UI

function onOpen() {
let ui = SpreadsheetApp.getUi()
ui.createMenu('User Template Check').addItem('Check User Import Template', 'checkUserImportTemplate').addToUi();
}
