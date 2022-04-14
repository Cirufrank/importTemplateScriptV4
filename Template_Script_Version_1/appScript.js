// @ts-nocheck
/*

Goals met:
 Checks for whiteSpace, remove it, highlights cell light green, inserts Comment 'WhiteSpace Removed'
 Check for duplicate, highlights cell, inserts comment 'Duplicate Email'
 Check to see if comment exisits, and if so, concatenates new comment instead of overriding the old comment(s)
 Check for missing first names, last names, or emails, highlights cells red on both report and main sheet, and    
 insterts comment based on item (i.e. `Missing ${firstName/lastName/email`)
 Does not highlight completely empty first name, last name, and email cells
 Automatically formats "User Date Added" and "Birthday(YYYY-MM-DD)" columns to yyyy-mm-dd format, inserts comment on report sheet with green background lettings users know that this was done
 Converts states that are fully written out, to their two letting codes, highlights cells on reports sheet, and insterts a comment on report sheet letting users know this was done


 Things to be aware of:
  Relies on a first name, last name, and email column being present to function correctly
  Needs each column header to be perfect
  The trim whitespace function is relied on heavily by other functions



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
const FIRST_NAME_MISSING_COMMENT = 'First name missing';
const LAST_NAME_MISSING_COMMENT = 'Last name missing';
const EMAIL_MISSING_COMMENT = 'Email missing';
const DATE_FORMATTED_YYYY_MM_DD_COMMENT = 'Date formatted to YYYY-MM-DD';
const STATE_TWO_LETTER_CODE_COMMENT = 'State converted to two letter code';
const LIGHT_GREEN_HEX_CODE = '#b6d7a8';
const LIGHT_RED_HEX_CODE = '#f4cccc';
const US_STATE_TO_ABBREVIATION = {
  "Alabama": "AL",
  "Alaska": "AK",
  "Arizona": "AZ",
  "Arkansas": "AR",
  "California": "CA",
  "Colorado": "CO",
  "Connecticut": "CT",
  "Delaware": "DE",
  "Florida": "FL",
  "Georgia": "GA",
  "Hawaii": "HI",
  "Idaho": "ID",
  "Illinois": "IL",
  "Indiana": "IN",
  "Iowa": "IA",
  "Kansas": "KS",
  "Kentucky": "KY",
  "Louisiana": "LA",
  "Maine": "ME",
  "Maryland": "MD",
  "Massachusetts": "MA",
  "Michigan": "MI",
  "Minnesota": "MN",
  "Mississippi": "MS",
  "Missouri": "MO",
  "Montana": "MT",
  "Nebraska": "NE",
  "Nevada": "NV",
  "New Hampshire": "NH",
  "New Jersey": "NJ",
  "New Mexico": "NM",
  "New York": "NY",
  "North Carolina": "NC",
  "North Dakota": "ND",
  "Ohio": "OH",
  "Oklahoma": "OK",
  "Oregon": "OR",
  "Pennsylvania": "PA",
  "Rhode Island": "RI",
  "South Carolina": "SC",
  "South Dakota": "SD",
  "Tennessee": "TN",
  "Texas": "TX",
  "Utah": "UT",
  "Vermont": "VT",
  "Virginia": "VA",
  "Washington": "WA",
  "West Virginia": "WV",
  "Wisconsin": "WI",
  "Wyoming": "WY",
  "District of Columbia": "DC",
  "American Samoa": "AS",
  "Guam": "GU",
  "Northern Mariana Islands": "MP",
  "Puerto Rico": "PR",
  "United States Minor Outlying Islands": "UM",
  "U.S. Virgin Islands": "VI",
}
const usFullStateNames = Object.keys(US_STATE_TO_ABBREVIATION);
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

function createArrayOfNamesAndEmail(firstNameRangeBinding, lastNameRangeBinding, emailColmnRangeBinding) {
  let firstNameRangeValues = getValues(firstNameRangeBinding);
  let lastNameRangeValues = getValues(lastNameRangeBinding);
  let emailColumnRangeValues = getValues(emailColmnRangeBinding);
  let firstNameLastNameEmailValueArrayBinding = [];
  for (let i = 0; i < emailColumnRangeValues.length; i += 1) {
  firstNameLastNameEmailValueArrayBinding.push([firstNameRangeValues[i], lastNameRangeValues[i], emailColumnRangeValues[i]]);
  }

   return firstNameLastNameEmailValueArrayBinding;
}

function checkForMissingNamesOrEmails(sheetBinding, reportSheetBinding) {
  let firstNameRange = getColumnRange('First Name', sheet);
  let firstNameRangePosition = firstNameRange.getColumn();
  let lastNameRange = getColumnRange('Last Name', sheet);
  let lastNameRangePosition = lastNameRange.getColumn();
  let emailColumnRange = getColumnRange('Email', sheet);
  let emailColumnPosition = emailColumnRange.getColumn();
  let firstNameLastNameEmailValueArray = createArrayOfNamesAndEmail(firstNameRange, lastNameRange, emailColumnRange);

  for (let i = 0; i < firstNameLastNameEmailValueArray.length; i +=1) {

    let row = i + 2;
    let firstName = String(firstNameLastNameEmailValueArray[i][0]);
    let firstNameCurrentCell = getSheetCell(sheetBinding, row, firstNameRangePosition);
    let reportSheetCurrentFNCell = getSheetCell(reportSheetBinding, row, firstNameRangePosition);
    let lastName = String(firstNameLastNameEmailValueArray[i][1]);
    let lastNameCurrentCell = getSheetCell(sheetBinding, row, lastNameRangePosition);
    let reportSheetLNCurrentCell = getSheetCell(reportSheetBinding, row, lastNameRangePosition);
    let email = String(firstNameLastNameEmailValueArray[i][2]);
    let emailCurrentCell = getSheetCell(sheetBinding, row, emailColumnPosition);
    let reportSheetEmailCurrentCell = getSheetCell(reportSheetBinding, row, emailColumnPosition);

    if (firstName.length !== 0 || lastName.length !== 0 || email.length !== 0) {
      firstNameLastNameEmailValueArray[i].forEach((val, index) => {
        let currentCellValue = String(val);
        if (currentCellValue === "") {
          switch(index) {
            case 0: 
              setSheetCellBackground(reportSheetCurrentFNCell, LIGHT_RED_HEX_CODE);
              setSheetCellBackground(firstNameCurrentCell, LIGHT_RED_HEX_CODE);
              insertCommentToSheetCell(reportSheetCurrentFNCell, FIRST_NAME_MISSING_COMMENT);
              break;
            case 1:
              setSheetCellBackground(reportSheetLNCurrentCell, LIGHT_RED_HEX_CODE);
              setSheetCellBackground(lastNameCurrentCell, LIGHT_RED_HEX_CODE);
              insertCommentToSheetCell(reportSheetLNCurrentCell, LAST_NAME_MISSING_COMMENT);
              break;
            case 2:
              setSheetCellBackground(reportSheetEmailCurrentCell, LIGHT_RED_HEX_CODE);
              setSheetCellBackground(emailCurrentCell, LIGHT_RED_HEX_CODE);
              insertCommentToSheetCell(reportSheetEmailCurrentCell, EMAIL_MISSING_COMMENT);
              break;
          }
        }
      })
    }
  }

  
}

function setColumnToYYYYMMDDFormat(columnRangeBinding) {
columnRangeBinding.setNumberFormat('yyyy-mm-dd');
}

function formatUserDateAddedAndBirthdayColumns(sheetBinding, reportSheetBinding) {
let row = 1;
let birthdayColumn = getColumnRange('Birthday (YYYY-MM-DD)', sheetBinding);
let userDateAddedColumn = getColumnRange('User Date Added', sheetBinding);

if (birthdayColumn && userDateAddedColumn) {
  let birthdayColumnPosition = birthdayColumn.getColumn();
  let reportSheetBirthdayColumnHeaderCell = getSheetCell(reportSheetBinding, row, birthdayColumnPosition);
  let userDateAddedColumnPosition = userDateAddedColumn.getColumn();
  let reportSheetUserDateAddedColumnHeaderCell = getSheetCell(reportSheetBinding, row, userDateAddedColumnPosition);

  setColumnToYYYYMMDDFormat(birthdayColumn);
  setColumnToYYYYMMDDFormat(userDateAddedColumn);
  setSheetCellBackground(reportSheetBirthdayColumnHeaderCell, LIGHT_GREEN_HEX_CODE);
  insertCommentToSheetCell(reportSheetBirthdayColumnHeaderCell, DATE_FORMATTED_YYYY_MM_DD_COMMENT);
  setSheetCellBackground(reportSheetUserDateAddedColumnHeaderCell, LIGHT_GREEN_HEX_CODE);
  insertCommentToSheetCell(reportSheetUserDateAddedColumnHeaderCell, DATE_FORMATTED_YYYY_MM_DD_COMMENT);
} else if (birthdayColumn) {
  let birthdayColumnNumberFormat = birthdayColumn.getNumberFormat();
  let birthdayColumnPosition = birthdayColumn.getColumn();
  let reportSheetBirthdayColumnHeaderCell = getSheetCell(reportSheetBinding, row, birthdayColumnPosition);
  setColumnToYYYYMMDDFormat(birthdayColumn);

  setSheetCellBackground(reportSheetBirthdayColumnHeaderCell, LIGHT_GREEN_HEX_CODE);
  insertCommentToSheetCell(reportSheetBirthdayColumnHeaderCell, DATE_FORMATTED_YYYY_MM_DD_COMMENT); 
} else if (userDateAddedColumn) {
  let userDateAddedColumnNumberFormat = userDateAddedColumn.getNumberFormat();
  let userDateAddedColumnPosition = userDateAddedColumn.getColumn();
  let reportSheetUserDateAddedColumnHeaderCell = getSheetCell(reportSheetBinding, row, userDateAddedColumnPosition);

  setColumnToYYYYMMDDFormat(userDateAddedColumn);
  setSheetCellBackground(reportSheetUserDateAddedColumnHeaderCell, LIGHT_GREEN_HEX_CODE);
  insertCommentToSheetCell(reportSheetUserDateAddedColumnHeaderCell, DATE_FORMATTED_YYYY_MM_DD_COMMENT);
}
}

function convertStatesToTwoLetterCode(sheetBinding, reportSheetBinding) {
  let stateColumnRange = getColumnRange('State (Ex: NH)', sheetBinding);

  if (stateColumnRange) {
    let stateColumnRangeValues = getValues(stateColumnRange).map(val => {
    let currentState = String(val);
    if (currentState.length > 2) {
      return (currentState[0].toUpperCase() + currentState.substr(1).toLowerCase());
    } else {
      return currentState;
    }
  });

  // SpreadsheetApp.getUi().alert(`${stateColumnRangeValues}`); Testing

    let stateColumnRangePosition = stateColumnRange.getColumn();
    
    stateColumnRangeValues.forEach((val, index) => {
      let currentState = String(val);
      // SpreadsheetApp.getUi().alert(`${currentState}`); Testing
      // SpreadsheetApp.getUi().alert(`${currentState.length}`); Testing
      let row = index + 2;
      let currentCell = getSheetCell(sheetBinding, row, stateColumnRangePosition);
      let currentReportCell = getSheetCell(reportSheetBinding, row, stateColumnRangePosition);
      if (currentState.length > 2 && usFullStateNames.includes(currentState)) {
        currentCell.setValue(US_STATE_TO_ABBREVIATION[currentState]);
        setSheetCellBackground(currentReportCell, LIGHT_RED_HEX_CODE);
        insertCommentToSheetCell(currentReportCell, STATE_TWO_LETTER_CODE_COMMENT);

      }
  });


  }
  

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
checkForMissingNamesOrEmails(sheet, reportSheet);
formatUserDateAddedAndBirthdayColumns(sheet, reportSheet);
convertStatesToTwoLetterCode(sheet, reportSheet);


}

  
// Creates "User Template Check" navigation button within Spreadsheet UI

function onOpen() {
let ui = SpreadsheetApp.getUi()
ui.createMenu('User Template Check').addItem('Check User Import Template', 'checkUserImportTemplate').addToUi();
}
