// @ts-nocheck
/*
ON 04/21/2022 VERSION OF TEMPLATE THAT SHOW EVERYTHING HAPPENING. USE FOR DEMOS

Run time with 8808 cels to check = about 3.5 minutes

Goals met:

 Checks for whiteSpace, remove it, highlights cell light green, inserts Comment 'WhiteSpace Removed'
 Check for duplicate, highlights cell, inserts comment 'Duplicate Email'
 Check to see if comment exisits, and if so, concatenates new comment instead of overriding the old comment(s)
 Check for missing first names, last names, or emails, highlights cells red on both report and main sheet, and    
 insterts comment based on item (i.e. `Missing ${firstName/lastName/email`)
 Does not highlight completely empty first name, last name, and email cells
 Automatically formats "User Date Added" and "Birthday(YYYY-MM-DD)" columns to yyyy-mm-dd format, inserts comment on report sheet with green background lettings users know that this was done
 Converts states that are fully written out, to their two letting codes, highlights cells on reports sheet, and insterts a comment on report sheet letting users know this was done
 Performs email validation, highlights main sheet and report sheet cells red, and insets a comment to reprot sheet cell saying "invalid email"
 Performs postal code validation, highlights main sheet and report sheet cells red, and insets a comment to reprot sheet cell saying "invalid postal code"
 Performs both home and mobile phone number validation
 Sets the last header cell of the report sheet to "REPORT SUMMARY" and provides a comment of all check that were ran
 Handles exceptions
 When errors are found, if not value is found in the assocaited row, the message "ERROR FOUND" is place within the cell so users can sort for error records (both reprot sheet and main sheet)
 Additionally, when errors are found the respective header column cell is hightlights list red and given a comment letting users know the errors that it contains
 Refactored the code to not have global variables declared with let so the code can be manageablly expanded to additional sheets
 Each function (such as check user templte) will have its own reprot sheet title so report sheets do not get deleted when multiple imports are ran

 ToDos: 
 
 Create documentation of script
 Add date served formatting, email check and validation, and hours check for individual hours template
 Publish the script privately for use by those frome the galaxydigital.com domain



 Things to be aware of:
  Relies on a first name, last name, and email column being present to function correctly
  Needs each column header to be perfect
  The trim whitespace function is relied on heavily by other functions
  Assumes that there is a zip, phone, and mobile column



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
const WHITE_SPACE_REMOVED_COMMENT = 'Whitespace Removed';
const DUPLICATE_EMAIL_COMMENT = 'Duplicate email';
const FIRST_NAME_MISSING_COMMENT = 'First name missing';
const LAST_NAME_MISSING_COMMENT = 'Last name missing';
const EMAIL_MISSING_COMMENT = 'Email missing';
const DATE_FORMATTED_YYYY_MM_DD_COMMENT = 'Date formatted to YYYY-MM-DD';
const STATE_TWO_LETTER_CODE_COMMENT = 'State converted to two letter code';
const CAPITALIZATION_FUNCTION_RAN_ON_FN_COMMENT = 'Capitalization function ran on first name column';
const CAPITALIZATION_FUNCTION_RAN_ON_LN_COMMENT = 'Capitalization function ran on last name column';
const INVALID_EMAIL_COMMENT = 'Invalid email format';
const INVALID_POSTAL_CODE_COMMENT = 'Invalid postal code';
const INVALID_PHONE_NUMBER_COMMENT = 'Invalid phone number';
const LIGHT_GREEN_HEX_CODE = '#b6d7a8';
const LIGHT_RED_HEX_CODE = '#f4cccc';
const ERROR_FOUND_MESSAGE_FOR_ROW = 'ERROR FOUND';
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
const ss = SpreadsheetApp.getActiveSpreadsheet();

function getSheet(sheetName) {
  return ss.getSheetByName(sheetName);
}

function setFrozenRows(sheetBinding, numOfRowsToFreeze) {
  sheetBinding.setFrozenRows(numOfRowsToFreeze);
}

function createSheetCopy(spreadSheet, newSheetName) {
  let newSheet = spreadSheet.insertSheet();
  // let newSheet = spreadSheet.duplicateActiveSheet();
  newSheet.setName(newSheetName);
  // valuesToCopyOver.forEach(row => newSheet.appendRow(row));

  return newSheet;
}

function copyValuesToSheet(targetSheet, valuesToCopyOver) {
  let rowCount =  valuesToCopyOver.length;
  let columnCount = valuesToCopyOver[0].length;
  let dataRange = targetSheet.getRange(1, 1, rowCount, columnCount);
  dataRange.setValues(valuesToCopyOver);
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

function getColumnRange(columnName, sheetName, columnsHeadersBinding) {
  let column =  columnsHeadersBinding.indexOf(columnName);
  if (column !== -1) {
    return sheetName.getRange(2,column+1,sheetName.getMaxRows());
  }

  return null;
}

function getValues(range) {
  return range.getValues();
}

function insertHeaderComment(headerCell, commentToInsert) {
  if (!headerCell.getComment().match(commentToInsert)) {
    if (!headerCell.getComment().match(",")) {
      insertCommentToSheetCell(headerCell, commentToInsert);
    } else {
      insertCommentToSheetCell(headerCell, `, ${commentToInsert}`);
    }
  }
}


function setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, rowBinding) {
  let reportSummaryCell = getSheetCell(reportSheetBinding, rowBinding, reportSummaryColumnPositionBinding);
  let reportSummaryCellValue = reportSummaryCell.getValue();
  let mainSheetErrorMessageCell = getSheetCell(sheetBinding, rowBinding, reportSummaryColumnPositionBinding);

   if (!reportSummaryCellValue) {
           reportSummaryCell.setValue(ERROR_FOUND_MESSAGE_FOR_ROW);
           mainSheetErrorMessageCell.setValue(ERROR_FOUND_MESSAGE_FOR_ROW);
          }
  
}

function removeWhiteSpaceFromCells(mainSheetBinding, valuesToCheck, reportSummaryCommentsBinding) {
let rowCount = valuesToCheck.length;
let columnCount = valuesToCheck[0].length;
let searchRange = mainSheetBinding.getRange(1, 1, rowCount, columnCount);
searchRange.trimWhitespace();

// let cellsTrimmed = [];

//   for (let row = 1; row < rowCount; row += 1) {
//   for (let column = 0; column < columnCount; column += 1) {
//     let currentVal = valuesToCheck[row][column];
//     let firstChar = currentVal.toString()[0];
//     let lastChar = currentVal[currentVal.toString().length - 1];
//     let currentCell = getSheetCell(mainSheetBinding, row + 1, column + 1);
//     let reportSheetCell = getSheetCell(reportSheetBinding, row + 1, column + 1);

//     if (firstChar === " " || lastChar === " ") {
//     currentCell.setValue(`${currentVal.trim()}`);
//     setSheetCellBackground(reportSheetCell, LIGHT_GREEN_HEX_CODE);
//     insertCommentToSheetCell(reportSheetCell, WHITE_SPACE_REMOVED_COMMENT);

//     //Remove later or check row and column information to be in A1 Notations

//     cellsTrimmed.push(`Value: ${currentVal} Row: ${currentCell.getRow()} Column: ${currentCell.getColumn()}`);
//     }
//   }
// }
//Remove later or use for logging events
// console.log(cellsTrimmed);
reportSummaryCommentsBinding.push("Success: removed white space from cells");
}


function checkForDuplicateEmails(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding, reportSummaryColumnPositionBinding) {
  let headerRow = 1;
  let emailColumnRange = getColumnRange('Email', sheetBinding, columnsHeadersBinding);
  let reportSheetEmailColumnRange = getColumnRange('Email', reportSheetBinding, columnsHeadersBinding);
  // let emailColumnValues = getValues(emailColumnRange).map(email => email[0].toLowerCase().trim());
  let emailColumnPosition = emailColumnRange.getColumn();
  let mainSheetHeaderCell = getSheetCell(sheetBinding, headerRow, emailColumnPosition);
  let beginningOfEmailColumnRange = emailColumnRange.getA1Notation().slice(0,2); //0,2 represents beginning cell for column range i.e.C1
  let endOfEmailColumnRange = emailColumnRange.getA1Notation().slice(3); // slice 3 represents the end of the column range excluding the comma
  // SpreadsheetApp.getUi().alert(`${beginningOfEmailColumnRange}: ${endOfEmailColumnRange}`); for tsting
  const duplicateEmailsRuleMainSheet = SpreadsheetApp.newConditionalFormatRule()
  .whenFormulaSatisfied(`=COUNTIF(${beginningOfEmailColumnRange}:${endOfEmailColumnRange}, ${beginningOfEmailColumnRange})>1`)
  .setBackground(LIGHT_RED_HEX_CODE)
  // .insertHeaderComment(mainSheetHeaderCell, "Duplicate Emails Found")
  .setRanges([emailColumnRange])
  .build();
  let rules = sheetBinding.getConditionalFormatRules();
  rules.push(duplicateEmailsRuleMainSheet);
  sheetBinding.setConditionalFormatRules(rules);

  const duplicateEmailsRuleReportSheet = SpreadsheetApp.newConditionalFormatRule()
  .whenFormulaSatisfied(`=COUNTIF(${beginningOfEmailColumnRange}:${endOfEmailColumnRange}, ${beginningOfEmailColumnRange})>1`)
  .setBackground(LIGHT_RED_HEX_CODE)
  // .setComment(DUPLICATE_EMAIL_COMMENT)
  .setRanges([reportSheetEmailColumnRange])
  .build();

  let rules2 = reportSheetBinding.getConditionalFormatRules();
  rules2.push(duplicateEmailsRuleReportSheet);
  reportSheetBinding.setConditionalFormatRules(rules2);

  // let duplicates = [];

  // emailColumnValues.forEach((email, index) => {
  //   let currentEmail = String(email);
  //   let row = index + 2;

  //     if (duplicates.indexOf(currentEmail) === -1 && currentEmail.length > 0) {
  //      if (emailColumnValues.filter(val => String(val) === currentEmail).length > 1) {
  //       let currentCell = getSheetCell(sheetBinding, row, emailColumnPosition);
  //       let reportSheetCell = getSheetCell(reportSheetBinding, row, emailColumnPosition);
  //       let mainSheetHeaderCell = getSheetCell(sheetBinding, headerRow, emailColumnPosition);


  //       //Line below for testing

  //       // SpreadsheetApp.getUi().alert(`Duplicate found! ${currentEmail} ${reportSheetCell.getA1Notation()}`);
  //       //testing line ends here
  //       setSheetCellBackground(reportSheetCell, LIGHT_RED_HEX_CODE);
  //       setSheetCellBackground(currentCell, LIGHT_RED_HEX_CODE);
  //       setSheetCellBackground(mainSheetHeaderCell, LIGHT_RED_HEX_CODE);
  //       insertCommentToSheetCell(reportSheetCell, DUPLICATE_EMAIL_COMMENT);
  //       insertHeaderComment(mainSheetHeaderCell, "Duplicate Emails Found");
  //       setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
  //       duplicates.push(currentEmail);
  //       }
  //     }
  //     }
    // );

  reportSheetBinding.sort(emailColumnPosition);
  sheetBinding.sort(emailColumnPosition);
  reportSummaryCommentsBinding.push("Success: checked emails column for duplicates");

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


function checkForMissingNamesOrEmails(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding, reportSummaryColumnPositionBinding) {
  let headerRow = 1;
  let firstNameRange = getColumnRange('First Name', sheetBinding, columnsHeadersBinding);
  let firstNameRangePosition = firstNameRange.getColumn();
  let lastNameRange = getColumnRange('Last Name', sheetBinding, columnsHeadersBinding);
  let lastNameRangePosition = lastNameRange.getColumn();
  let emailColumnRange = getColumnRange('Email', sheetBinding, columnsHeadersBinding);
  let emailColumnPosition = emailColumnRange.getColumn();
  let firstNameLastNameEmailValueArray = createArrayOfNamesAndEmail(firstNameRange, lastNameRange, emailColumnRange);

  for (let i = 0; i < firstNameLastNameEmailValueArray.length; i +=1) {

    let row = i + 2;
    let firstName = String(firstNameLastNameEmailValueArray[i][0]);
    let firstNameCurrentCell = getSheetCell(sheetBinding, row, firstNameRangePosition);
    let mainSheetFNHeaderCell = getSheetCell(sheetBinding, headerRow, firstNameRangePosition);
    let reportSheetCurrentFNCell = getSheetCell(reportSheetBinding, row, firstNameRangePosition);
    let lastName = String(firstNameLastNameEmailValueArray[i][1]);
    let lastNameCurrentCell = getSheetCell(sheetBinding, row, lastNameRangePosition);
    let mainSheetLNHeaderCell = getSheetCell(sheetBinding, headerRow, lastNameRangePosition);
    let reportSheetLNCurrentCell = getSheetCell(reportSheetBinding, row, lastNameRangePosition);
    let email = String(firstNameLastNameEmailValueArray[i][2]);
    let emailCurrentCell = getSheetCell(sheetBinding, row, emailColumnPosition);
    let mainSheetEmailHeaderCell = getSheetCell(sheetBinding, headerRow, emailColumnPosition);
    let reportSheetEmailCurrentCell = getSheetCell(reportSheetBinding, row, emailColumnPosition);

    if (firstName.length !== 0 || lastName.length !== 0 || email.length !== 0) {
      firstNameLastNameEmailValueArray[i].forEach((val, index) => {
      let currentCellValue = String(val);
      if (currentCellValue === "") {
        switch(index) {
          case 0: 
            setSheetCellBackground(reportSheetCurrentFNCell, LIGHT_RED_HEX_CODE);
            setSheetCellBackground(firstNameCurrentCell, LIGHT_RED_HEX_CODE);
            setSheetCellBackground(mainSheetFNHeaderCell, LIGHT_RED_HEX_CODE);
            insertCommentToSheetCell(reportSheetCurrentFNCell, FIRST_NAME_MISSING_COMMENT);
            setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
            insertHeaderComment(mainSheetFNHeaderCell, "First Name/Names Missing");
            break;
          case 1:
            setSheetCellBackground(reportSheetLNCurrentCell, LIGHT_RED_HEX_CODE);
            setSheetCellBackground(lastNameCurrentCell, LIGHT_RED_HEX_CODE);
            setSheetCellBackground(mainSheetLNHeaderCell, LIGHT_RED_HEX_CODE);
            insertCommentToSheetCell(reportSheetLNCurrentCell, LAST_NAME_MISSING_COMMENT);
            setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
            insertHeaderComment(mainSheetLNHeaderCell, "Last Name/Names Missing");
            break;
          case 2:
            setSheetCellBackground(reportSheetEmailCurrentCell, LIGHT_RED_HEX_CODE);
            setSheetCellBackground(emailCurrentCell, LIGHT_RED_HEX_CODE);
            setSheetCellBackground(mainSheetEmailHeaderCell, LIGHT_RED_HEX_CODE);
            insertCommentToSheetCell(reportSheetEmailCurrentCell, EMAIL_MISSING_COMMENT);
            setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
            insertHeaderComment(mainSheetEmailHeaderCell, "Email/Emails Missing");
            break;
        }
      }
    })
  }
}
reportSummaryCommentsBinding.push("Success: checked for missing names and emails");
}

function setColumnToYYYYMMDDFormat(columnRangeBinding) {
  columnRangeBinding.setNumberFormat('yyyy-mm-dd');
}

function formatUserDateAddedAndBirthdayColumns(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding) {
let row = 1;
let birthdayColumn = getColumnRange('Birthday (YYYY-MM-DD)', sheetBinding, columnsHeadersBinding);
let userDateAddedColumn = getColumnRange('User Date Added', sheetBinding, columnsHeadersBinding);

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
  reportSummaryCommentsBinding.push("Success: formatted birthday column and user date added column");
} else if (birthdayColumn) {
  let birthdayColumnNumberFormat = birthdayColumn.getNumberFormat();
  let birthdayColumnPosition = birthdayColumn.getColumn();
  let reportSheetBirthdayColumnHeaderCell = getSheetCell(reportSheetBinding, row, birthdayColumnPosition);
  setColumnToYYYYMMDDFormat(birthdayColumn);

  setSheetCellBackground(reportSheetBirthdayColumnHeaderCell, LIGHT_GREEN_HEX_CODE);
  insertCommentToSheetCell(reportSheetBirthdayColumnHeaderCell, DATE_FORMATTED_YYYY_MM_DD_COMMENT); 
  reportSummaryCommentsBinding.push("Success: formatted birthday column and did not find user date added column");
} else if (userDateAddedColumn) {
  let userDateAddedColumnNumberFormat = userDateAddedColumn.getNumberFormat();
  let userDateAddedColumnPosition = userDateAddedColumn.getColumn();
  let reportSheetUserDateAddedColumnHeaderCell = getSheetCell(reportSheetBinding, row, userDateAddedColumnPosition);

  setColumnToYYYYMMDDFormat(userDateAddedColumn);
  setSheetCellBackground(reportSheetUserDateAddedColumnHeaderCell, LIGHT_GREEN_HEX_CODE);
  insertCommentToSheetCell(reportSheetUserDateAddedColumnHeaderCell, DATE_FORMATTED_YYYY_MM_DD_COMMENT);
  reportSummaryCommentsBinding.push("Success: formatted user date added column and did not find birthday column")
}

}

function convertStatesToTwoLetterCode(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding) {
  let stateColumnRange = getColumnRange('State (Ex: NH)', sheetBinding, columnsHeadersBinding);

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

  reportSummaryCommentsBinding.push("Success: ran check to convert states to two-digit format");


  } else {
    reportSummaryCommentsBinding.push("Success: ran check to convert states to two-digit format but did not find column");

  }
  

}



function capitalizeFirstLetterOfAName(name) {
  return name.split(" ").filter(name => name.length > 0 && name).map(name => name.trim()[0].toUpperCase() + name.trim().substr(1).toLowerCase()).join(" ").split("-").map(name => name.trim()).filter(name => name !== "-" && name.length > 0).map(name => (name.split(" ").length > 1) ? name : name.trim()[0].toUpperCase() + name.trim().substr(1).toLowerCase()).join("-");
}

//Not using this function for now, but saving for later or personal use
// function capitalizeFirstLetterOfWords(sheetBinding, reportSheetBinding, columnsHeadersBinding) {
//   let firstNameRange = getColumnRange('First Name', sheetBinding, columnsHeadersBinding);
//   let firstNameColumnPosition = firstNameRange.getColumn();
//   let lastNameRange = getColumnRange('Last Name', sheetBinding, columnsHeadersBinding);
//   let lastNameRangeColumnPosition = lastNameRange.getColumn();
//   let firstNameRangeValues = getValues(firstNameRange);
//   let lastNameRangeValues = getValues(lastNameRange);
//   let reportSheetFirstNameHeaderCell = getSheetCell(reportSheetBinding, 1, firstNameColumnPosition);
//   let reportSheetLastNameHeaderCell = getSheetCell(reportSheetBinding, 1, lastNameRangeColumnPosition);

//   firstNameRangeValues.forEach((name, index) => {
//     let currentName = String(name);
//     if (currentName.length > 0) {
//       let row = index + 2;
//       let currentCell = getSheetCell(sheetBinding, row, firstNameColumnPosition);
//       currentCell.setValue(capitalizeFirstLetterOfAName(currentName));
    
//     }
  
//   });

//   lastNameRangeValues.forEach((name, index) => {
//     let currentName = String(name);

//     if (currentName.length > 0) {
//       let row = index + 2;
//       let currentCell = getSheetCell(sheetBinding, row, lastNameRangeColumnPosition);
//       currentCell.setValue(capitalizeFirstLetterOfAName(currentName));
    
//     }
  
//   });

//   setSheetCellBackground(reportSheetFirstNameHeaderCell, LIGHT_GREEN_HEX_CODE);
//   insertCommentToSheetCell(reportSheetFirstNameHeaderCell, CAPITALIZATION_FUNCTION_RAN_ON_FN_COMMENT);
//   setSheetCellBackground(reportSheetLastNameHeaderCell, LIGHT_GREEN_HEX_CODE);
//   insertCommentToSheetCell(reportSheetLastNameHeaderCell, CAPITALIZATION_FUNCTION_RAN_ON_LN_COMMENT);

// } 

const validateEmail = (email) => {
return String(email)
  .toLowerCase()
  .match(
    /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/
  );
};

const validatePhoneNumbers = (number) => {
return String(number)
  .toLowerCase()
  .match(
/^[\+]?[(]?[0-9]{3}[)]?[-\s\.]?[0-9]{3}[-\s\.]?[0-9]{4,6}$/im
  );
};

const validatePostalCode = (postalCode) => {
return /(^\d{5}$)|(^\d{5}-\d{4}$)/.test(postalCode);
}
// Consider this match: /^w+([.-]?w+)*@w+([.-]?w+)*(.w{2,3})+$/;

function checkForInvalidEmails(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding, reportSummaryColumnPositionBinding) {
  let headerRow = 1;
  let emailColumnRange = getColumnRange('Email', sheetBinding, columnsHeadersBinding);
  let emailColumnValues = getValues(emailColumnRange);
  let emailColumnPosition = emailColumnRange.getColumn();

  emailColumnValues.forEach((email, index) => {
    let currentEmail = String(email);
    let row = index + 2;

      if (currentEmail !== "" && !validateEmail(currentEmail)) {
        let currentCell = getSheetCell(sheetBinding, row, emailColumnPosition);
        let reportSheetCell = getSheetCell(reportSheetBinding, row, emailColumnPosition);
        let mainSheetHeaderCell = getSheetCell(sheetBinding, headerRow, emailColumnPosition);

        //Line below for testing

        // SpreadsheetApp.getUi().alert(`Invalid Email! ${currentEmail} ${reportSheetCell.getA1Notation()}`);
        //testing line ends here
        setSheetCellBackground(reportSheetCell, LIGHT_RED_HEX_CODE);
        setSheetCellBackground(currentCell, LIGHT_RED_HEX_CODE);
        setSheetCellBackground(mainSheetHeaderCell, LIGHT_RED_HEX_CODE);
        insertCommentToSheetCell(reportSheetCell, INVALID_EMAIL_COMMENT);
        insertHeaderComment(mainSheetHeaderCell, "Invalid Email/Emails found");
        setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
        }
      }
    );

    reportSummaryCommentsBinding.push("Success: ran check for invalid emails");
}

function checkForInvalidNumbers(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding, reportSummaryColumnPositionBinding) {
  let headerRow = 1;
  let homePhoneNumberRange = getColumnRange('Phone', sheetBinding, columnsHeadersBinding);
  let cellPhoneNumberRange = getColumnRange('Mobile', sheetBinding, columnsHeadersBinding);

  if (homePhoneNumberRange) {
    let homePhoneNumberRangeValues = getValues(homePhoneNumberRange);
    let homePhoneNumberRangePosition = homePhoneNumberRange.getColumn();
    let mainSheetHomePhoneHeaderCell = getSheetCell(sheetBinding, headerRow, homePhoneNumberRangePosition);

    homePhoneNumberRangeValues.forEach((number, index) => {

      let currentNumber= String(number);
      let row = index + 2;

      if (currentNumber !== "" && !validatePhoneNumbers(currentNumber)) {
        let currentCell = getSheetCell(sheetBinding, row, homePhoneNumberRangePosition);
        let reportSheetCell = getSheetCell(reportSheetBinding, row, homePhoneNumberRangePosition);

        //Line below for testing

        // SpreadsheetApp.getUi().alert(`Invalid Email! ${currentEmail} ${reportSheetCell.getA1Notation()}`);
        //testing line ends here
        setSheetCellBackground(reportSheetCell, LIGHT_RED_HEX_CODE);
        setSheetCellBackground(currentCell, LIGHT_RED_HEX_CODE);
        setSheetCellBackground(mainSheetHomePhoneHeaderCell, LIGHT_RED_HEX_CODE);
        insertCommentToSheetCell(reportSheetCell, INVALID_PHONE_NUMBER_COMMENT);
        insertHeaderComment(mainSheetHomePhoneHeaderCell, "Invalid Home Phone Number/Numbers Found");
        setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
        }
      }
      
    );

    reportSummaryCommentsBinding.push("Success: ran check for invalid home phone numbers");

  } else {
    reportSummaryCommentsBinding.push("Success: ran check for invalid home phone numbers, but did not find column");
  }

  if (cellPhoneNumberRange) {
    let cellPhoneNumberRangeValues = getValues(cellPhoneNumberRange);
    let cellPhoneNumberRangePosition = cellPhoneNumberRange.getColumn();
    let mainSheetCellPhoneHeaderCell = getSheetCell(sheetBinding, headerRow, cellPhoneNumberRangePosition);

    cellPhoneNumberRangeValues.forEach((number, index) => {

    let currentNumber= String(number);
    let row = index + 2;

      if (currentNumber !== "" && !validatePhoneNumbers(currentNumber)) {
        let currentCell = getSheetCell(sheetBinding, row, cellPhoneNumberRangePosition);
        let reportSheetCell = getSheetCell(reportSheetBinding, row, cellPhoneNumberRangePosition);

        //Line below for testing

        // SpreadsheetApp.getUi().alert(`Invalid Email! ${currentEmail} ${reportSheetCell.getA1Notation()}`);
        //testing line ends here
        setSheetCellBackground(reportSheetCell, LIGHT_RED_HEX_CODE);
        setSheetCellBackground(currentCell, LIGHT_RED_HEX_CODE);
        setSheetCellBackground(mainSheetCellPhoneHeaderCell, LIGHT_RED_HEX_CODE);
        insertCommentToSheetCell(reportSheetCell, INVALID_PHONE_NUMBER_COMMENT);
        insertHeaderComment(mainSheetCellPhoneHeaderCell, "Invalid Cell Phone Number/Numbers Found");
        setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
        }
      }
    );

    reportSummaryCommentsBinding.push("Success: ran check for invalid mobile phone numbers");

  } else {
    reportSummaryCommentsBinding.push("Success: ran check for invalid mobile phone numbers, but did not find column");
  }

}

function checkForInvalidPostalCodes(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding, reportSummaryColumnPositionBinding) {
  let headerRow = 1;
  let postalCodeColumnRange = getColumnRange('Zip', sheetBinding, columnsHeadersBinding);

  if (postalCodeColumnRange) {
    let postalCodeColumnRangeValues = getValues(postalCodeColumnRange);
    let postalCodeColumnRangePosition = postalCodeColumnRange.getColumn();
    let mainSheetPostalHeaderCell = getSheetCell(sheetBinding, headerRow, postalCodeColumnRangePosition);

    postalCodeColumnRangeValues.forEach((code, index) => {
      let currentCode = String(code);
      let row = index + 2;

      if (currentCode !== "" && !validatePostalCode(currentCode)) {
        let currentCell = getSheetCell(sheetBinding, row, postalCodeColumnRangePosition);
        let reportSheetCell = getSheetCell(reportSheetBinding, row, postalCodeColumnRangePosition);

        //Line below for testing

        // SpreadsheetApp.getUi().alert(`Invalid Email! ${currentEmail} ${reportSheetCell.getA1Notation()}`);
        //testing line ends here
        setSheetCellBackground(reportSheetCell, LIGHT_RED_HEX_CODE);
        setSheetCellBackground(currentCell, LIGHT_RED_HEX_CODE);
        setSheetCellBackground(mainSheetPostalHeaderCell, LIGHT_RED_HEX_CODE);
        insertCommentToSheetCell(reportSheetCell, INVALID_POSTAL_CODE_COMMENT);
        insertHeaderComment(mainSheetPostalHeaderCell, "Invalid Postal Code/Codes Found");
        setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
        }
      }
    );

    reportSummaryCommentsBinding.push("Success: ran check for invalid postal codes");
  } else {
    reportSummaryCommentsBinding.push("Success: ran check for invalid postal codes, but did not find column");
  }
  
}

function clearSheetSummaryColumn(sheetBinding, reportSummaryColumnPositionBinding) {
  let row = 1;
  const clearRange = () => {sheetBinding.getRange(1,reportSummaryColumnPositionBinding, sheetBinding.getMaxRows()).clear();
 }
 clearRange();
}

function setErrorColumnHeaderInMainSheet(sheetBinding, reportSummaryColumnPositionBinding) {
 let row = 1;

let mainSheetErrorMessageCell = getSheetCell(sheetBinding, row, reportSummaryColumnPositionBinding);
let mainSheetErrorMessageCellValue = mainSheetErrorMessageCell.getValue();

   if (!mainSheetErrorMessageCellValue) {
           mainSheetErrorMessageCell.setValue("Error Alert Column");
          }

}

function setCommentsOnReportCell(reportSheetBinding, reportSummaryCommentsBinding, reportSummaryColumnPositionBinding) {
  let reportSummaryCommentsString = reportSummaryCommentsBinding.join(", ");
 let row = 1;
 
 let reportCell = getSheetCell(reportSheetBinding, row, reportSummaryColumnPositionBinding);

 reportCell.setValue(`REPORT OVERVIEW`);
 insertCommentToSheetCell(reportCell, reportSummaryCommentsString);
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
function createReportSheet(ssBinding, valuesBinding, reportName) {
let reportSheetBinding = getSheet(reportName);
deleteSheetIfExists(reportSheetBinding);
reportSheetBinding = createSheetCopy(ssBinding, reportName);
copyValuesToSheet(reportSheetBinding, valuesBinding);

return reportSheetBinding;
}

try {
  function checkUserImportTemplate() {
  const REPORT_SHEET_NAME = 'User Report';
  const USER_IMPORT_TEMPLATE_NAME = 'User Import Template';
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange();
  const values = data.getValues();
  const columnHeaders = values[0].map(header => header.trim());
  //Consider values[0].map(header => (header[0] + header.substr(1)).trim()) to help avoid errors here
  const reportSummaryColumnPosition = columnHeaders.length + 1;
  const reportSummaryComments = [];
  let reportSheet = createReportSheet(ss, values, REPORT_SHEET_NAME);

  sheet.setName(USER_IMPORT_TEMPLATE_NAME);
  
  ss.setActiveSheet(getSheet(USER_IMPORT_TEMPLATE_NAME));
 

  setFrozenRows(sheet, 1);
  setFrozenRows(reportSheet, 1);

  
  try {
    removeWhiteSpaceFromCells(sheet, values, reportSummaryComments);
  } catch (err) {
    Logger.log(err);
    reportSummaryComments.push("Failed: remove white space from cells");
    throw new Error(`White space not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }

  clearSheetSummaryColumn(sheet, reportSummaryColumnPosition);
  clearSheetSummaryColumn(reportSheet, reportSummaryColumnPosition);
  setErrorColumnHeaderInMainSheet(sheet, reportSummaryColumnPosition);
  
  try {
    checkForDuplicateEmails(sheet, reportSheet, columnHeaders, reportSummaryComments, reportSummaryColumnPosition);
  } catch(err) {
    Logger.log(err);
    reportSummaryComments.push("Failed: check emails column for duplicates");
    throw new Error(`Emails not checked for duplicates. Reason: ${err.name}: ${err.message}. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
  }

  try{
    checkForMissingNamesOrEmails(sheet, reportSheet, columnHeaders, reportSummaryComments, reportSummaryColumnPosition);
  } catch(err) {
    Logger.log(err);
    reportSummaryComments.push("Failed: check for missing names and emails");
    throw new Error(`Check not ran for missing names and emails. Reason: ${err.name}: ${err.message}.  Please revert sheet to previous version, ensure the first name, last name, and email columns are titled "Email", "First Name", and "Last Name" within their header columns, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
  }

  try {
    formatUserDateAddedAndBirthdayColumns(sheet, reportSheet, columnHeaders, reportSummaryComments);
  } catch(err) {
    Logger.log(err);
    reportSummaryComments.push("Failed: did not format user date added and birthday columns");
    throw new Error(`Check not ran for formatting of user date added and birthday columns. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  
  try {
    convertStatesToTwoLetterCode(sheet, reportSheet, columnHeaders, reportSummaryComments);
  } catch (err) {
    Logger.log(err);
    reportSummaryComments.push("Failed: check not ran for converting states to two-letter code");
    throw new Error(`Check not ran for converting states to two-letter code: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  
  // capitalizeFirstLetterOfWords(sheet, reportSheet);

  try {
    checkForInvalidEmails(sheet, reportSheet, columnHeaders, reportSummaryComments, reportSummaryColumnPosition);
  } catch(err) {
      Logger.log(err);
      reportSummaryComments.push("Failed: check not ran for invalid emails");
      throw new Error(`Check not ran for invalid emails. Reason: ${err.name}: ${err.message} at line. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
    } 
  
  try {
    checkForInvalidPostalCodes(sheet, reportSheet, columnHeaders, reportSummaryComments, reportSummaryColumnPosition);
  } catch(err) {
    Logger.log(err);
    reportSummaryComments.push("Failed: check not ran for invalid postal codes");
    throw new Error(`Check not ran for invalid postal codes. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  
  try {
    checkForInvalidNumbers(sheet, reportSheet,columnHeaders, reportSummaryComments, reportSummaryColumnPosition);
  } catch(err) {
    Logger.log(err);
    reportSummaryComments.push("Failed: check not ran for invalid home or mobile phone numbers");
    throw new Error(`Check not ran for invalid home or mobile phone numbers. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  
  try {
    setCommentsOnReportCell(reportSheet, reportSummaryComments, reportSummaryColumnPosition);
  } catch(err) {
    Logger.log(err);
    throw new Error(`Report sheet cell comment not added for summary of checks. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }

   SpreadsheetApp.getUi().alert("User Import Check Complete");
  
}

} catch(err) {
Logger.log(err);
throw new Error(`An error occured the the user import template check did not successfully run. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);


}


  
// Creates "User Template Check" navigation button within Spreadsheet UI
// try {
  function onOpen() {
  let ui = SpreadsheetApp.getUi()
  ui.createMenu('User Template Check').addItem('Check User Import Template', 'checkUserImportTemplate').addToUi();
}

// } catch (err) {
//   SpreadsheetApp.getUi().alert(`An error occured and the user import template check button did not load. Please refresh the page: ${err.name}: ${err.message}`);
//   Logger.log(err);
// }


