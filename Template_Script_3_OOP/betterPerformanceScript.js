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
 Now does relies on column position for checking missing names and emails and such (checks for missing values in first three columns)
 Clears formatting of values at beggining of import
 Checks for invalid states and test for Canada states and postal codes

 ToDos: 
 
 Create documentation of script
 Add date served formatting, email check and validation, and hours check for individual hours template DONE YASSS
 Publish the script privately for use by those frome the galaxydigital.com domain
 insert head columns automatically (data will be position based)
 insert header columns for loggically for users at the end (i.e. "Email => "Contact Email")



 Things to be aware of:
  If two check are ran on the same template an error will be thrown (may want to kepp this to remind user to revert doc)
  Now, users can create multiple report sheets
  Users ad Programs/Agencies import template checks have been tested, need to test others
  Relies on a first name, last name, and email column being present to function correctly
  Needs each column header to be perfect
  The trim whitespace function is relied on heavily by other functions
  Assumes that there is a zip, phone, and mobile column
  Will have the opportunity to consolidate many functions once the templates are position-based



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

/*

Organiing the code with OOP:

class TemplateCheck

 */


class Template {
    constructor() {
      this.ss = SpreadsheetApp.getActiveSpreadsheet();
      this._sheet = SpreadsheetApp.getActiveSheet();
      this._templateName = this._sheet.getName();
      this._errorAlertColumnHeader = "Error Alert Column";
      this._reportOverviewHeaderMessage = `REPORT OVERVIEW`;
      this._whiteSpaceRemovedComment = 'Whitespace Removed';
      this._whiteSpaceRemovedSuccessMessage = "Success: removed white space from cells";
      this._removedFormattingSuccessMessage = "Success: removed formatting from cells";
      this._ranEmailValidationCheckSuccessMessage = "Success: ran check for invalid emails";
      this._valuesMissingComment = 'Values Missing';
      this._invalidEmailComment = 'Invalid email format';
      this._invalidEmailCellMessage = "Invalid Email/Emails found";
      this._failedRemovedWhiteSpaceFromCellsComment = "Failed: remove white space from cells";
      this._failedRemovedFormattingFromCellsComment = "Failed: remove formatting from cells";
      this._failedInvalidEmailCheckMessage = "Failed: check not ran for invalid emails";
      this._dateFormat = 'yyyy-mm-dd';
      this._lightGreenHexCode = '#b6d7a8';
      this._lightRedHexCode = '#f4cccc';
      this._errorFoundMessage = 'ERROR FOUND';
      this._dateFormattedComment = 'Date formatted to YYYY-MM-DD';
      
      this._data = this._sheet.getDataRange();
      this._values = this._data.getValues();
      this._columnHeaders = this._values[0].map(header => header.trim());
      this._reportSummaryColumnPosition = this._columnHeaders.length + 1;
      this._reportSummaryComments = [];
      
    }
    get whiteSpaceRemovedComment() {
      return this._whiteSpaceRemovedComment;
    }
    get whiteSpaceRemovedSuccessMessage() {
      return this._whiteSpaceRemovedSuccessMessage;
    }
    get removedFormattingSuccessMessage() {
      return this._removedFormattingSuccessMessage;
    }
    get valuesMissingComment() {
      return this._valuesMissingComment;
    }
    get invalidEmailComment() {
      return this._invalidEmailComment;
    }
    get dateFormat() {
      return this._dateFormat;
    }
    get lightGreenHexCode() {
      return this._lightGreenHexCode;
    }
    get lightRedHexCode() {
      return this._lightRedHexCode;
    }
    get errorFoundMessage() {
      return this._errorFoundMessage;
    }
    get dateFormattedComment() {
      return this._dateFormattedComment;
    }
    get templateName() {
      return this._templateName;
    }
    get reportSheetName() {
      return this._reportSheetName;
    }
    get sheet() {
      return this._sheet;
    }
    get data() {
      return this._data;
    }
    get values() {
      return this._values;
    }
    get columnHeaders() {
      return this._columnHeaders;
    }
    get reportSummaryColumnPosition() {
      return this._reportSummaryColumnPosition;
    }
    get reportSummaryComments() {
      return this._reportSummaryComments;
    }
    get failedRemovedWhiteSpaceFromCellsComment() {
      return this._failedRemovedWhiteSpaceFromCellsComment;
    }
    get failedRemovedFormattingFromCellsComment() {
      return this._failedRemovedFormattingFromCellsComment;
    }
    get failedInvalidEmailCheckMessage() {
      return this._failedInvalidEmailCheckMessage;
    }
    set reportSummaryComments(comments) {
      this._reportSummaryComments.push(comments);
    }
    getSheet(sheetName) {
      return this.ss.getSheetByName(sheetName);
    }
    setFrozenRows(sheetBinding, numOfRowsToFreeze) { 
    sheetBinding.setFrozenRows(numOfRowsToFreeze);
    }
    createSheetCopy(newSheetName) { 
      let newSheet = this.ss.insertSheet();
      newSheet.setName(newSheetName);
  
      return newSheet;
    }
    copyValuesToSheet(targetSheet, valuesToCopyOver) {
      let rowCount =  valuesToCopyOver.length;
      let columnCount = valuesToCopyOver[0].length;
      let dataRange = targetSheet.getRange(1, 1, rowCount, columnCount);
      dataRange.setValues(valuesToCopyOver);
    }
    getSheetCell(sheetBinding, row, column) {
      return sheetBinding.getRange(row, column);
    }
    setSheetCellBackground(sheetCellBinding, color) {
      sheetCellBinding.setBackground(color);
    }
    checkCellForComment(sheetCell) {
      return sheetCell.getComment() ? true: false;
    }
    insertCommentToSheetCell(sheetCell, comment) {
      if (this.checkCellForComment(sheetCell)) {
        let previousComment = sheetCell.getComment();
        let newComment = `${previousComment}, ${comment}`;
        sheetCell.setComment(newComment);
      } else {
        sheetCell.setComment(comment);
      }
    }
    deleteSheetIfExists(sheetBinding) {
      if (sheetBinding !== null) {
      this.ss.deleteSheet(sheetBinding);
      }
    }
   createReportSheet() { 
      // let reportSheetBinding = this.getSheet(this._reportSheetName); 
      // this.deleteSheetIfExists(reportSheetBinding);
      let reportSheetName = `${this.templateName} Report`;
      reportSheetBinding = this.createSheetCopy(reportSheetName);
      this.copyValuesToSheet(reportSheetBinding, this._values);
      return reportSheetBinding;
    }
    getColumnRange(columnName, sheetBinding) {
      let column =  this._columnHeaders.indexOf(columnName);
      if (column !== -1) {
        return sheetBinding.getRange(2,column+1,sheetBinding.getMaxRows());
      }
      return null;
    }
    getValues(range) {
      return range.getValues();
    }
    insertHeaderComment(headerCell, commentToInsert) {
      if (!headerCell.getComment().match(commentToInsert)) {
        if (!headerCell.getComment().match(",")) {
          this.insertCommentToSheetCell(headerCell, commentToInsert);
        } else {
          this.insertCommentToSheetCell(headerCell, `, ${commentToInsert}`);
        }
      }
    }
    clearSheetSummaryColumn(sheetBinding) {
      let row = 1;
      const clearRange = () => {sheetBinding.getRange(row,this._reportSummaryColumnPosition, sheetBinding.getMaxRows()).clear();
    }
    clearRange();
  }
    setErrorColumns(reportSheetBinding, rowBinding) {
      let reportSummaryCell = this.getSheetCell(reportSheetBinding, rowBinding, this._reportSummaryColumnPosition);
      let reportSummaryCellValue = reportSummaryCell.getValue();
      let mainSheetErrorMessageCell = this.getSheetCell(this._sheet, rowBinding, this._reportSummaryColumnPosition);
  
      if (!reportSummaryCellValue) {
              reportSummaryCell.setValue(this._errorFoundMessage);
              mainSheetErrorMessageCell.setValue(this._errorFoundMessage);
              }
        
      }
      
    setErrorColumnHeaderInMainSheet() {
      let row = 1;
  
      let mainSheetErrorMessageCell = this.getSheetCell(this._sheet, row, this._reportSummaryColumnPosition);
      let mainSheetErrorMessageCellValue = mainSheetErrorMessageCell.getValue();
  
      if (!mainSheetErrorMessageCellValue) {
              mainSheetErrorMessageCell.setValue(this._errorAlertColumnHeader);
              }
  
    }
    setCommentsOnReportCell(reportSheetBinding) {
      let reportSummaryCommentsString = this._reportSummaryComments.join(", ");
      let row = 1;
      
      let reportCell = this.getSheetCell(reportSheetBinding, row, this._reportSummaryColumnPosition);
  
      reportCell.setValue(this._reportOverviewHeaderMessage);
      this.insertCommentToSheetCell(reportCell, reportSummaryCommentsString);
    }
  
    // Function that executes when user clicks "User Template Check" Button
    createReportSheet() {
      // let reportSheetBinding = this.getSheet(reportName); 
      // this.deleteSheetIfExists(reportSheetBinding);
      let reportSheetName = `${this.sheet.getName()} Report`;
     let  reportSheetBinding = this.createSheetCopy(this.ss, reportSheetName);
      reportSheetBinding.setName(reportSheetName);
      this.copyValuesToSheet(reportSheetBinding, this._values);
  
      return reportSheetBinding;
    }
  
    removeWhiteSpaceFromCells() {
      let rowCount = this._values.length;
      let columnCount = this._values[0].length;
      let searchRange = this._sheet.getRange(1, 1, rowCount, columnCount);
      searchRange.trimWhitespace();
  
  
      this._reportSummaryComments.push(this._whiteSpaceRemovedSuccessMessage);
      }
    removeFormattingFromSheetCells() {
      let rowCount = this._values.length;
      let columnCount = this._values[0].length;
      let searchRange = this._sheet.getRange(2, 1, rowCount, columnCount);
      searchRange.clearFormat();
  
      this._reportSummaryComments.push(this._removedFormattingSuccessMessage);
    }
    setColumnToYYYYMMDDFormat(columnRangeBinding) { //ALL BUT PROGRAMS AND AGENCIES
      columnRangeBinding.setNumberFormat(this._dateFormat);
    }
    validateEmail(email){
      return String(email)
        .toLowerCase()
        .match(
          /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/);
    }
  
   checkForInvalidEmails(reportSheetBinding) {
      let headerRow = 1;
      let emailColumnRange = this.getColumnRange('Email', this._sheet, this._columnHeaders);
      let emailColumnValues = this.getValues(emailColumnRange);
      let emailColumnPosition = emailColumnRange.getColumn();
  
      emailColumnValues.forEach((email, index) => {
        let currentEmail = String(email);
        let row = index + 2;
  
          if (currentEmail !== "" && !this.validateEmail(currentEmail)) {
            let currentCell = this.getSheetCell(this._sheet, row, emailColumnPosition);
            let reportSheetCell = this.getSheetCell(reportSheetBinding, row, emailColumnPosition);
            let mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, emailColumnPosition);
            this.setSheetCellBackground(reportSheetCell, this._lightRedHexCode);
            this.setSheetCellBackground(currentCell, this._lightRedHexCode);
            this.setSheetCellBackground(mainSheetHeaderCell, this._lightRedHexCode);
            this.insertCommentToSheetCell(reportSheetCell, this._invalidEmailComment);
            this.insertHeaderComment(mainSheetHeaderCell, this._invalidEmailCellMessage);
            this.setErrorColumns(reportSheetBinding, row);
            }
          }
        );
  
        this._reportSummaryComments.push(this._ranEmailValidationCheckSuccessMessage);
    }
  
  }
  
  class UsersNeedsAndAgenciesTemplate extends Template {
    constructor() {
      super();
      this._stateTwoLetterCodeComment = 'State converted to two letter code'; 
      this._invalidPostalCodeComment = 'Invalid postal code'; 
      this._invalidPhoneNumberComment = 'Invalid phone number'; 
      this._invalidStateComment = 'Invalid State'; 
      this._invalidStateFoundHeaderComment = "Invalid State/States found";
      this._invalidHomePhoneNumbersFoundHeaderComment =  "Invalid Home Phone Number/Numbers Found";
      this._invalidCellPhoneNumbersFoundHeaderComment = "Invalid Cell Phone Number/Numbers Found";
      this._invalidPostalCodeFoundHeaderComment = "Invalid Postal Code/Codes Found";
      this._stateColumnFoundAndConversionFunctionRanComment = "Success: ran check to convert states to two-digit format";
      this._stateColumnNotFoundAndConversionRanComment = "Success: ran check to convert states to two-digit format, but did not find column";
      this._stateColumnFoundAndValidationCheckRanComment = "Success: ran check for invalid states";
      this._stateColumnNotFoundAndValidationCheckRanComment = "Success: ran check for invalid states, but did not find column";
      this._homePhoneNumberColumnFoundAndValidationCheckRanComment = "Success: ran check for invalid home phone numbers";
      this._homePhoneNumberColumnNotFoundAndValidationCheckRanComment = "Success: ran check for invalid home phone numbers, but did not find column";
      this._mobilePhoneNumberColumnFoundAndValidationCheckRanComment = "Success: ran check for invalid mobile phone numbers";
      this._mobilePhoneNumberColumnNotFoundAndValidationCheckRanComment = "Success: ran check for invalid mobile phone numbers, but did not find column";
      this._postalCodeColumnFoundAndValidationCheckRan = "Success: ran check for invalid postal codes";
      this._postalCodeColumnNotFoundAndValidationCheckRan = "Success: ran check for invalid postal codes, but did not find column";
      this._failedCheckNotRanForConvertingStatesToTwoLetterCodes = "Failed: check not ran for converting states to two-letter code";
      this._failedInvalidStatesCheckMessage = "Failed: check not ran for invalid states";
      this._failedInvalidPostalCodeCheck = "Failed: check not ran for invalid postal codes";
      this._failedInvalidHomeAndMobilePhoneNumbersCheck = "Failed: check not ran for invalid home or mobile phone numbers";
      this._usStateToAbbreviation = {
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
      this._usFullStateNames = Object.keys(this._usStateToAbbreviation); 
      this._usStateAbbreviations = Object.values(this._usStateToAbbreviation);
    }
    get failedCheckNotRanForConvertingStatesToTwoLetterCodes() {
      return this._failedCheckNotRanForConvertingStatesToTwoLetterCodes;
    }
    get failedInvalidStatesCheckMessage() {
      return this._failedInvalidStatesCheckMessage;
    }
    get failedInvalidPostalCodeCheck() {
      return this._failedInvalidPostalCodeCheck;
    }
    get failedInvalidHomeAndMobilePhoneNumbersCheck() {
      return this._failedInvalidHomeAndMobilePhoneNumbersCheck;
    }
    convertStatesToTwoLetterCode(reportSheetBinding) {
      let stateColumnRange = this.getColumnRange('State', this._sheet);
  
      if (stateColumnRange) {
        let stateColumnRangeValues = this.getValues(stateColumnRange).map(val => {
        let currentState = String(val);
        if (currentState.length > 2) {
          return (currentState[0].toUpperCase() + currentState.substr(1).toLowerCase());
        } else {
          return currentState;
        }
      });
  
       let stateColumnRangePosition = stateColumnRange.getColumn();
      
        stateColumnRangeValues.forEach((val, index) => {
          let currentState = String(val);
          let row = index + 2;
          let currentCell = this.getSheetCell(this._sheet, row, stateColumnRangePosition);
          let currentReportCell = this.getSheetCell(reportSheetBinding, row, stateColumnRangePosition);
          if (currentState.length > 2 && this._usFullStateNames.includes(currentState)) {
            currentCell.setValue(this._usStateToAbbreviation[currentState]);
            this.setSheetCellBackground(currentReportCell, this._lightRedHexCode);
            this.insertCommentToSheetCell(currentReportCell, this._stateTwoLetterCodeComment);
  
          }
      });
  
      this._reportSummaryComments.push(this._stateColumnFoundAndConversionFunctionRanComment);
  
  
      } else {
        this._reportSummaryComments.push(this._stateColumnNotFoundAndConversionRanComment);
  
      }
      
  
    }
    validateStates(reportSheetBinding) {
      let stateColumnRange = this.getColumnRange('State', this._sheet);
  
      if (stateColumnRange) {
        let stateColumnRangeValues = this.getValues(stateColumnRange);
  
        let stateColumnRangePosition = stateColumnRange.getColumn();
        
        stateColumnRangeValues.forEach((val, index) => {
          let currentState = String(val);
          
          if (!this._usStateAbbreviations.includes(currentState) && currentState.length !== 0) {
            let headerRow = 1;
            let row = index + 2;
            let currentCell = this.getSheetCell(this._sheet, row, stateColumnRangePosition);
            let currentReportCell = this.getSheetCell(reportSheetBinding, row, stateColumnRangePosition);
            let mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, stateColumnRangePosition);
            this.setSheetCellBackground(currentCell, this._lightRedHexCode);
            this.setSheetCellBackground(currentReportCell, this._lightRedHexCode);
            this.insertCommentToSheetCell(currentReportCell, this._invalidStateComment);
            this.setSheetCellBackground(mainSheetHeaderCell, this._lightRedHexCode);
            this.insertHeaderComment(mainSheetHeaderCell, this._invalidStateFoundHeaderComment);
            this.setErrorColumns(reportSheetBinding,row);
          }
      });
  
      this._reportSummaryComments.push(this._stateColumnFoundAndValidationCheckRanComment);
  
  
      } else {
        this._reportSummaryComments.push(this._stateColumnNotFoundAndValidationCheckRanComment);
  
      }
      
  
    }
    validatePhoneNumbers(number) { 
      return String(number)
        .toLowerCase()
        .match(
      /^[\+]?[(]?[0-9]{3}[)]?[-\s\.]?[0-9]{3}[-\s\.]?[0-9]{4,6}$/im
        );
    }
  
    validatePostalCode(postalCode) { 
      return /(^\d{5}$)|(^\d{5}-\d{4}$)/.test(postalCode);
    }
    checkForInvalidNumbers(reportSheetBinding) {
      let headerRow = 1;
      let homePhoneNumberRange = this.getColumnRange('Phone', this._sheet);
      let cellPhoneNumberRange = this.getColumnRange('Mobile', this._sheet);
  
      if (homePhoneNumberRange) {
        let homePhoneNumberRangeValues = this.getValues(homePhoneNumberRange);
        let homePhoneNumberRangePosition = homePhoneNumberRange.getColumn();
        let mainSheetHomePhoneHeaderCell = this.getSheetCell(this._sheet, headerRow, homePhoneNumberRangePosition);
  
        homePhoneNumberRangeValues.forEach((number, index) => {
  
          let currentNumber= String(number);
          let row = index + 2;
  
          if (currentNumber !== "" && !this.validatePhoneNumbers(currentNumber)) {
            let currentCell = this.getSheetCell(this._sheet, row, homePhoneNumberRangePosition);
            let reportSheetCell = this.getSheetCell(reportSheetBinding, row, homePhoneNumberRangePosition);
  
            this.setSheetCellBackground(reportSheetCell, this._lightRedHexCode);
            this.setSheetCellBackground(currentCell, this._lightRedHexCode);
            this.setSheetCellBackground(mainSheetHomePhoneHeaderCell, this._lightRedHexCode);
            this.insertCommentToSheetCell(reportSheetCell, this._invalidPhoneNumberComment);
            this.insertHeaderComment(mainSheetHomePhoneHeaderCell,this._invalidHomePhoneNumbersFoundHeaderComment);
            this.setErrorColumns(reportSheetBinding, row);
            }
          }
          
        );
  
        this._reportSummaryComments.push(this._homePhoneNumberColumnFoundAndValidationCheckRanComment);
  
      } else {
        this._reportSummaryComments.push(this._homePhoneNumberColumnNotFoundAndValidationCheckRanComment);
      }
  
      if (cellPhoneNumberRange) {
        let cellPhoneNumberRangeValues = this.getValues(cellPhoneNumberRange);
        let cellPhoneNumberRangePosition = cellPhoneNumberRange.getColumn();
        let mainSheetCellPhoneHeaderCell = this.getSheetCell(this._sheet, headerRow, cellPhoneNumberRangePosition);
  
        cellPhoneNumberRangeValues.forEach((number, index) => {
  
        let currentNumber= String(number);
        let row = index + 2;
  
          if (currentNumber !== "" && !this.validatePhoneNumbers(currentNumber)) {
            let currentCell = this.getSheetCell(this._sheet, row, cellPhoneNumberRangePosition);
            let reportSheetCell = this.getSheetCell(reportSheetBinding, row, cellPhoneNumberRangePosition);
  
            this.setSheetCellBackground(reportSheetCell, this._lightRedHexCode);
            this.setSheetCellBackground(currentCell, this._lightRedHexCode);
            this.setSheetCellBackground(mainSheetCellPhoneHeaderCell, this._lightRedHexCode);
            this.insertCommentToSheetCell(reportSheetCell, this._invalidPhoneNumberComment);
            this.insertHeaderComment(mainSheetCellPhoneHeaderCell, this._invalidCellPhoneNumbersFoundHeaderComment);
            this.setErrorColumns(reportSheetBinding, row);
            }
          }
        );
  
        this._reportSummaryComments.push(this._mobilePhoneNumberColumnFoundAndValidationCheckRanComment);
  
      } else {
        this._reportSummaryComments.push(this._mobilePhoneNumberColumnNotFoundAndValidationCheckRanComment);
      }
  
    }
    checkForInvalidPostalCodes(reportSheetBinding) {
      let headerRow = 1;
      let postalCodeColumnRange = this.getColumnRange('Zip', this._sheet);
  
      if (postalCodeColumnRange) {
        let postalCodeColumnRangeValues = this.getValues(postalCodeColumnRange);
        let postalCodeColumnRangePosition = postalCodeColumnRange.getColumn();
        let mainSheetPostalHeaderCell = this.getSheetCell(this._sheet, headerRow, postalCodeColumnRangePosition);
  
        postalCodeColumnRangeValues.forEach((code, index) => {
          let currentCode = String(code);
          let row = index + 2;
  
          if (currentCode !== "" && !this.validatePostalCode(currentCode)) {
            let currentCell = this.getSheetCell(this._sheet, row, postalCodeColumnRangePosition);
            let reportSheetCell = this.getSheetCell(reportSheetBinding, row, postalCodeColumnRangePosition);
            this.setSheetCellBackground(reportSheetCell, this._lightRedHexCode);
            this.setSheetCellBackground(currentCell, this._lightRedHexCode);
            this.setSheetCellBackground(mainSheetPostalHeaderCell, this._lightRedHexCode);
            this.insertCommentToSheetCell(reportSheetCell, this._invalidPostalCodeComment);
            this.insertHeaderComment(mainSheetPostalHeaderCell, this._invalidPostalCodeFoundHeaderComment);
            this.setErrorColumns(reportSheetBinding, row);
            }
          }
        );
  
        this._reportSummaryComments.push(this._postalCodeColumnFoundAndValidationCheckRan);
      } else {
        this._reportSummaryComments.push(this._postalCodeColumnNotFoundAndValidationCheckRan);
      }
      
    }
  
  }
  
  class UserTemplate extends UsersNeedsAndAgenciesTemplate {
    constructor(templateName, reportSheetName) {
      super(templateName, reportSheetName);
      this._duplicateEmailComment = 'Duplicate email';
      this._genderExpandedComment = 'Gender option expanded';
      this._invalidGenderOptionComment = 'Invalid Gender Option';
      this._genderOptionsAbbreviationToFullObject = {
        "F":"Female",
        "M":"Male",
        "N/A":"",
      }
      this._genderOptionsAbbreviations = Object.keys(this._genderOptionsAbbreviationToFullObject);
      this._fullGenderOptions = ['female', 'male','prefer not to say', 'other'];
      this._foundEmailColumnAndCheckedForDuplicatesComment = "Success: checked emails column for duplicates";
      this._foundFirstThreeColumnsAndCheckedForMissingValuesComment = "Success: checked for missing values in first three columns";
      this._foundUserDateAddedAndBirthdayColumnsAndFormattedThemComment = "Success: formatted birthday column and user date added column";
      this._didNotFindUserDateAddedColumnButDidFindBirthdayColumnAndFormattedItComment = "Success: formatted birthday column and did not find user date added column";
      this._didNotFindBirthdayColumnButDidFindUserDateAddedColumnAndFormattedItComment = "Success: formatted user date added column and did not find birthday column";
      this._invalidGenderOptionsFoundHeaderComment = "Invalid Gender Option/Options found";
      this._foundGenderOptionsColumnAndRanCheckSuccessfullyComment = "Success: ran check for invalid gender options";
      this._didNotFindGenderOptionsColumnAndRanCheckSuccessfullyComment = "Success: ran check for invalid gender options, but did not find column";
      this._failedEmailDuplicatesCheckMessage = "Failed: check emails column for duplicates";
      this._failedFirstThreeColumnsBlankCheck = "Failed: check first three columns for missing values";
      this._failedFormatUserDateAddedAndBirthdayColumnsCheckMessage = "Failed: did not format user date added and birthday columns";
      this._failedInvalidGenderOptionsCheckMessage = "Failed: check not ran for invalid gender options";
  
    }
    get failedEmailDuplicatesCheckMessage() {
      return this._failedEmailDuplicatesCheckMessage;
    }
    get failedFirstThreeColumnsBlankCheck() {
      return this._failedFirstThreeColumnsBlankCheck;
    }
    get failedFormatUserDateAddedAndBirthdayColumnsCheckMessage() {
      return this._failedFormatUserDateAddedAndBirthdayColumnsCheckMessage;
    }
    get failedInvalidGenderOptionsCheckMessage() {
      return this._failedInvalidGenderOptionsCheckMessage;
    }
    checkForDuplicateEmails(reportSheetBinding) { 
      let headerRow = 1;
      let emailColumnRange = this.getColumnRange('Email', this._sheet);
      let reportSheetEmailColumnRange = this.getColumnRange('Email', reportSheetBinding);
      let emailColumnPosition = emailColumnRange.getColumn();
      let beginningOfEmailColumnRange = emailColumnRange.getA1Notation().slice(0,2); //0,2 represents beginning cell for column range i.e.C1
      let endOfEmailColumnRange = emailColumnRange.getA1Notation().slice(3); // slice 3 represents the end of the column range excluding the comma
      const duplicateEmailsRuleMainSheet = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=COUNTIF(${beginningOfEmailColumnRange}:${endOfEmailColumnRange}, ${beginningOfEmailColumnRange})>1`)
      .setBackground(this._lightRedHexCode)
      .setRanges([emailColumnRange])
      .build();
      let rules = this._sheet.getConditionalFormatRules();
      rules.push(duplicateEmailsRuleMainSheet);
      this._sheet.setConditionalFormatRules(rules);
  
      const duplicateEmailsRuleReportSheet = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=COUNTIF(${beginningOfEmailColumnRange}:${endOfEmailColumnRange}, ${beginningOfEmailColumnRange})>1`)
      .setBackground(this._lightRedHexCode)
      .setRanges([reportSheetEmailColumnRange])
      .build();
  
      let rules2 = reportSheetBinding.getConditionalFormatRules();
      rules2.push(duplicateEmailsRuleReportSheet);
      reportSheetBinding.setConditionalFormatRules(rules2);
      reportSheetBinding.sort(emailColumnPosition);
      this._sheet.sort(emailColumnPosition);
      this._reportSummaryComments.push(this._foundEmailColumnAndCheckedForDuplicatesComment);
  
    }
    checkFirstThreeColumnsForBlanks(reportSheetBinding) {
      let rowStartPosition = 2;
      let columnStartPosition = 1;
      let middleColumnPosition = 2;
      let thirdColumnPositon = 3;
      let maxRows = this._sheet.getMaxRows();
      let totalColumsToCheck = 3;
      let range = this._sheet.getRange(rowStartPosition, columnStartPosition, maxRows, totalColumsToCheck);
      let values = range.getValues();
  
      for (let row = 0; row < values.length; row += 1) {
        let headerRowPosition = 1;
        let cellRow = row + 2;
        let currentRow = values[row];
        let item1 = currentRow[0];
        let item1CurrentCell = this.getSheetCell(this._sheet, cellRow, columnStartPosition);
        let item1ReportCurrentCell = this.getSheetCell(reportSheetBinding, cellRow, columnStartPosition);
        let item1HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, columnStartPosition);
        let item2 = currentRow[1];
        let item2CurrentCell = this.getSheetCell(this._sheet, cellRow, middleColumnPosition);
        let item2ReportCurrentCell = this.getSheetCell(reportSheetBinding, cellRow, middleColumnPosition);
        let item2HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, middleColumnPosition);
        let item3 = currentRow[2];
        let item3CurrentCell = this.getSheetCell(this._sheet, cellRow, thirdColumnPositon);
        let item3ReportCell = this.getSheetCell(reportSheetBinding, cellRow, thirdColumnPositon);
        let item3HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, thirdColumnPositon);
  
        if (item1.length !== 0 || item2.length !== 0  || item3.length !== 0) {
          currentRow.forEach((val, index) => {
            if (val === "") {
              switch (index) {
                case 0:
                  this.setSheetCellBackground(item1CurrentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item1ReportCurrentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item1HeaderCell, this._lightRedHexCode);
                  this.insertCommentToSheetCell(item1ReportCurrentCell, this._valuesMissingComment);
                  this.setErrorColumns(reportSheetBinding, cellRow);
                  this.insertHeaderComment(item1HeaderCell, this._valuesMissingComment);
                  break;
                case 1:
                  this.setSheetCellBackground(item2CurrentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item2ReportCurrentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item2HeaderCell, this._lightRedHexCode);
                  this.insertCommentToSheetCell(item2ReportCurrentCell, this._valuesMissingComment);
                  this.setErrorColumns(reportSheetBinding, cellRow);
                  this.insertHeaderComment(item2HeaderCell, this._valuesMissingComment);
                  break;
                case 2:
                  this.setSheetCellBackground(item3CurrentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item3ReportCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item3HeaderCell, this._lightRedHexCode);
                  this.insertCommentToSheetCell(item3ReportCell, this._valuesMissingComment);
                  this.setErrorColumns(reportSheetBinding, cellRow);
                  this.insertHeaderComment(item3HeaderCell, this._valuesMissingComment);
                  break;
              }
            }
          });
        }
      }
  
    this._reportSummaryComments.push(this._foundFirstThreeColumnsAndCheckedForMissingValuesComment);
  
    }
    formatUserDateAddedAndBirthdayColumns(reportSheetBinding) {
    let row = 1;
    let birthdayColumn = this.getColumnRange('Birthday (YYYY-MM-DD)', this._sheet);
    let userDateAddedColumn = this.getColumnRange('User Date Added', this._sheet);
  
    if (birthdayColumn && userDateAddedColumn) {
        let birthdayColumnPosition = birthdayColumn.getColumn();
        let reportSheetBirthdayColumnHeaderCell = this.getSheetCell(reportSheetBinding, row, birthdayColumnPosition);
        let userDateAddedColumnPosition = userDateAddedColumn.getColumn();
        let reportSheetUserDateAddedColumnHeaderCell = this.getSheetCell(reportSheetBinding, row, userDateAddedColumnPosition);
  
        this.setColumnToYYYYMMDDFormat(birthdayColumn);
        this.setColumnToYYYYMMDDFormat(userDateAddedColumn);
        this.setSheetCellBackground(reportSheetBirthdayColumnHeaderCell, this._lightRedHexCode);
        this.insertCommentToSheetCell(reportSheetBirthdayColumnHeaderCell, this._dateFormattedComment);
        this.setSheetCellBackground(reportSheetUserDateAddedColumnHeaderCell, this._lightRedHexCode);
        this.insertCommentToSheetCell(reportSheetUserDateAddedColumnHeaderCell, this._dateFormattedComment);
        this._reportSummaryComments.push(this._foundFirstThreeColumnsAndCheckedForMissingValuesComment);
  
    } else if (birthdayColumn) {
        let birthdayColumnPosition = birthdayColumn.getColumn();
        let reportSheetBirthdayColumnHeaderCell = this.getSheetCell(reportSheetBinding, row, birthdayColumnPosition);
        this.setColumnToYYYYMMDDFormat(birthdayColumn);
  
        this.setSheetCellBackground(reportSheetBirthdayColumnHeaderCell, this._lightGreenHexCode);
        this.insertCommentToSheetCell(reportSheetBirthdayColumnHeaderCell, this._dateFormattedComment); 
        this._reportSummaryComments.push(this._didNotFindUserDateAddedColumnButDidFindBirthdayColumnAndFormattedItComment);
    } else if (userDateAddedColumn) {
        let userDateAddedColumnPosition = userDateAddedColumn.getColumn();
        let reportSheetUserDateAddedColumnHeaderCell = this.getSheetCell(reportSheetBinding, row, userDateAddedColumnPosition);
  
        this.setColumnToYYYYMMDDFormat(userDateAddedColumn);
        this.setSheetCellBackground(reportSheetUserDateAddedColumnHeaderCell, this._lightGreenHexCode);
        this.insertCommentToSheetCell(reportSheetUserDateAddedColumnHeaderCell, this._dateFormattedComment);
        this._reportSummaryComments.push(this._didNotFindBirthdayColumnButDidFindUserDateAddedColumnAndFormattedItComment);
      }
  
    }
    validateGenderOptions(reportSheetBinding) {
      let genderOptionColumnRange = this.getColumnRange('Gender', this._sheet);
  
      if (genderOptionColumnRange) {
        let genderOptionColumnRangeValues = this.getValues(genderOptionColumnRange);
  
        let genderOptionColumnRangePoition = genderOptionColumnRange.getColumn();
        
        genderOptionColumnRangeValues.forEach((val, index) => {
          let currentGenderOption = String(val).toLowerCase();
          
          if (!this._fullGenderOptions.includes(currentGenderOption) && currentGenderOption.length !== 0) {
            let headerRow = 1;
            let row = index + 2;
            let currentCell = this.getSheetCell(this._sheet, row, genderOptionColumnRangePoition);
            let currentReportCell = this.getSheetCell(reportSheetBinding, row, genderOptionColumnRangePoition);
            let mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, genderOptionColumnRangePoition);
            this.setSheetCellBackground(currentCell, this._lightRedHexCode);
            this.setSheetCellBackground(currentReportCell, this._lightRedHexCode);
            this.insertCommentToSheetCell(currentReportCell, this._invalidGenderOptionComment);
            this.setSheetCellBackground(mainSheetHeaderCell, this._lightRedHexCode);
            this.insertHeaderComment(mainSheetHeaderCell, this._invalidGenderOptionsFoundHeaderComment);
            this.setErrorColumns(reportSheetBinding, row);
  
          }
      });
  
      this._reportSummaryComments.push(this._foundGenderOptionsColumnAndRanCheckSuccessfullyComment);
  
  
      } else {
        this._reportSummaryComments.push(this._didNotFindGenderOptionsColumnAndRanCheckSuccessfullyComment);
  
      }
      
  
    }
  
  }
  
  // const WHITE_SPACE_REMOVED_COMMENT = 'Whitespace Removed'; //All templates TEMPLATE CLASS added
  // const VALUES_MISSING_COMMENT = 'Values Missing'; // all templates TEMPLATE CLASS added
  // const INVALID_EMAIL_COMMENT = 'Invalid email format'; // ALL TEMPLATES TEAMPLTE CLASS added
  // const LIGHT_GREEN_HEX_CODE = '#b6d7a8'; //ALL TEMPLATES TEMPLATE CLASS  added
  // const LIGHT_RED_HEX_CODE = '#f4cccc'; //ALL TEMPLATE TEMPLATE CLASS added
  // const ERROR_FOUND_MESSAGE_FOR_ROW = 'ERROR FOUND'; //ALL TEMPLATES, TEMPLATE CLASS added
  // const ss = SpreadsheetApp.getActiveSpreadsheet(); //ALL TEMPLATE CLASS added 
  // const DATE_FORMATTED_YYYY_MM_DD_COMMENT = 'Date formatted to YYYY-MM-DD'; // ALL BUT PROGRAMS/AGENCIES TEMPLATE
  
   
    // const REPORT_SHEET_NAME = 'User Report'; ADDED
    // const USER_IMPORT_TEMPLATE_NAME = 'User Import Template'; ADDED
    // const sheet = SpreadsheetApp.getActiveSheet(); ADDED
    // const data = sheet.getDataRange(); ADDED
    // const values = data.getValues(); ADDED
    // const columnHeaders = values[0].map(header => header.trim()); ADDED
    //Consider values[0].map(header => (header[0] + header.substr(1)).trim()) to help avoid errors here
    // const reportSummaryColumnPosition = columnHeaders.length + 1; ADDED
    // const reportSummaryComments = []; ADDED
  
  
    /*
    
     TO DOOOOOOO
  
  
    let reportSheet = createReportSheet(ss, values, REPORT_SHEET_NAME);
  
    sheet.setName(USER_IMPORT_TEMPLATE_NAME);
    
  
    ss.setActiveSheet(getSheet(USER_IMPORT_TEMPLATE_NAME));
   
  
    setFrozenRows(sheet, 1);
    setFrozenRows(reportSheet, 1);
      clearSheetSummaryColumn(sheet, reportSummaryColumnPosition);
    clearSheetSummaryColumn(reportSheet, reportSummaryColumnPosition);
    setErrorColumnHeaderInMainSheet(sheet, reportSummaryColumnPosition);
  
  
  TOOO DOOOOO
  
  */
  
  
  // function getSheet(sheetName) { //ALL TEMPLATE CLASS ADDED
  //   return ss.getSheetByName(sheetName);
  // }
  
  // function setFrozenRows(sheetBinding, numOfRowsToFreeze) { //ALL TEMPLATE CLASS ADDED
  //   sheetBinding.setFrozenRows(numOfRowsToFreeze);
  // }
  
  // function createSheetCopy(spreadSheet, newSheetName) { //ALL TEMPLATE CLASS ADDED
  //   let newSheet = spreadSheet.insertSheet();
  //   // let newSheet = spreadSheet.duplicateActiveSheet();
  //   newSheet.setName(newSheetName);
  //   // valuesToCopyOver.forEach(row => newSheet.appendRow(row));
  
  //   return newSheet;
  // }
  
  // function copyValuesToSheet(targetSheet, valuesToCopyOver) { //ALL TEMPLATE CLASS ADDED
  //   let rowCount =  valuesToCopyOver.length;
  //   let columnCount = valuesToCopyOver[0].length;
  //   let dataRange = targetSheet.getRange(1, 1, rowCount, columnCount);
  //   dataRange.setValues(valuesToCopyOver);
  // }
  
  //   function getSheetCell(sheetBinding, row, column) { //ALL TEMPLATE CLASS ADDED
  //   return sheetBinding.getRange(row, column);
  // }
  
  // function setSheetCellBackground(sheetCellBinding, color) { //ALL TEMPLATE CLASS ADDED
  //   sheetCellBinding.setBackground(color);
  // }
  
  // function checkCellForComment(sheetCell) { //ALL TEMPLATE CLASS ADDED
  //   return sheetCell.getComment() ? true: false;
  // }
  
  // //Consider searching for if the comment already exists and not adding a new one in that case
  // function insertCommentToSheetCell(sheetCell, comment) { //ALL TEMPLATE CLASS ADDED
  //   if (checkCellForComment(sheetCell)) {
  //     let previousComment = sheetCell.getComment();
  //     let newComment = `${previousComment}, ${comment}`;
  //     sheetCell.setComment(newComment);
  //   } else {
  //     sheetCell.setComment(comment);
  //   }
  
  // }
  
  // function deleteSheetIfExists(sheetBinding) { //ALL TEMPLATE CLASS ADDED
  //   if (sheetBinding !== null) {
  //    ss.deleteSheet(sheetBinding);
  //   }
  // }
  
  // function getColumnRange(columnName, sheetName, columnsHeadersBinding) { //ALL TEMPLATE CLASS ADDED
  //   let column =  columnsHeadersBinding.indexOf(columnName);
  //   if (column !== -1) {
  //     return sheetName.getRange(2,column+1,sheetName.getMaxRows());
  //   }
  
  //   return null;
  // }
  
  // function getValues(range) { //ALL TEMPLATE CLASS ADDED
  //   return range.getValues();
  // }
  
  // function insertHeaderComment(headerCell, commentToInsert) { //ALL TEMPLATE CLASS ADDED
  //   if (!headerCell.getComment().match(commentToInsert)) {
  //     if (!headerCell.getComment().match(",")) {
  //       insertCommentToSheetCell(headerCell, commentToInsert);
  //     } else {
  //       insertCommentToSheetCell(headerCell, `, ${commentToInsert}`);
  //     }
  //   }
  // }
  
  
  // function setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, rowBinding) { //ALL TEMPLATE CLASS ADDED
  //   let reportSummaryCell = getSheetCell(reportSheetBinding, rowBinding, reportSummaryColumnPositionBinding);
  //   let reportSummaryCellValue = reportSummaryCell.getValue();
  //   let mainSheetErrorMessageCell = getSheetCell(sheetBinding, rowBinding, reportSummaryColumnPositionBinding);
  
  //    if (!reportSummaryCellValue) {
  //            reportSummaryCell.setValue(ERROR_FOUND_MESSAGE_FOR_ROW);
  //            mainSheetErrorMessageCell.setValue(ERROR_FOUND_MESSAGE_FOR_ROW);
  //           }
    
  // }
  
  // function removeWhiteSpaceFromCells(mainSheetBinding, valuesToCheck, reportSummaryCommentsBinding) { //ALL TEMPLATE CLASS ADDED
  // let rowCount = valuesToCheck.length;
  // let columnCount = valuesToCheck[0].length;
  // let searchRange = mainSheetBinding.getRange(1, 1, rowCount, columnCount);
  // searchRange.trimWhitespace();
  
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
  // reportSummaryCommentsBinding.push("Success: removed white space from cells");
  // }
  
  // function removeFormattingFromSheetCells(sheetBinding, valuesToCheck, reportSummaryCommentsBinding) { //ALL TEMPLATE CLASS ADDED
  // let rowCount = valuesToCheck.length;
  // let columnCount = valuesToCheck[0].length;
  // let searchRange = sheetBinding.getRange(2, 1, rowCount, columnCount);
  // searchRange.clearFormat();
  
  // reportSummaryCommentsBinding.push("Success: removed formatting from cells");
  // }
  
  // function setColumnToYYYYMMDDFormat(columnRangeBinding) { //ALL BUT PROGRAMS AND AGENCIES
  //   columnRangeBinding.setNumberFormat('yyyy-mm-dd');
  // }
  
  // const validateEmail = (email) => { //ALL TEMPLATES ADDED
  // return String(email)
  //   .toLowerCase()
  //   .match(
  //     /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/
  //   );
  // };
  
  // function checkForInvalidEmails(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding, reportSummaryColumnPositionBinding) { //ALL TEMPLATES ADDED
  //   let headerRow = 1;
  //   let emailColumnRange = getColumnRange('Email', sheetBinding, columnsHeadersBinding);
  //   let emailColumnValues = getValues(emailColumnRange);
  //   let emailColumnPosition = emailColumnRange.getColumn();
  
  //   emailColumnValues.forEach((email, index) => {
  //     let currentEmail = String(email);
  //     let row = index + 2;
  
  //       if (currentEmail !== "" && !validateEmail(currentEmail)) {
  //         let currentCell = getSheetCell(sheetBinding, row, emailColumnPosition);
  //         let reportSheetCell = getSheetCell(reportSheetBinding, row, emailColumnPosition);
  //         let mainSheetHeaderCell = getSheetCell(sheetBinding, headerRow, emailColumnPosition);
  
  //         //Line below for testing
  
  //         // SpreadsheetApp.getUi().alert(`Invalid Email! ${currentEmail} ${reportSheetCell.getA1Notation()}`);
  //         //testing line ends here
  //         setSheetCellBackground(reportSheetCell, LIGHT_RED_HEX_CODE);
  //         setSheetCellBackground(currentCell, LIGHT_RED_HEX_CODE);
  //         setSheetCellBackground(mainSheetHeaderCell, LIGHT_RED_HEX_CODE);
  //         insertCommentToSheetCell(reportSheetCell, INVALID_EMAIL_COMMENT);
  //         insertHeaderComment(mainSheetHeaderCell, "Invalid Email/Emails found");
  //         setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
  //         }
  //       }
  //     );
  
  //     reportSummaryCommentsBinding.push("Success: ran check for invalid emails");
  // }
  
  // function clearSheetSummaryColumn(sheetBinding, reportSummaryColumnPositionBinding) { //ALL TEMPLATES ADDED
  //   let row = 1;
  //   const clearRange = () => {sheetBinding.getRange(1,reportSummaryColumnPositionBinding, sheetBinding.getMaxRows()).clear();
  //  }
  //  clearRange();
  // }
  
  // function setErrorColumnHeaderInMainSheet(sheetBinding, reportSummaryColumnPositionBinding) { //ALL TEMPLATES ADDED
  // let row = 1;
  
  // let mainSheetErrorMessageCell = getSheetCell(sheetBinding, row, reportSummaryColumnPositionBinding);
  // let mainSheetErrorMessageCellValue = mainSheetErrorMessageCell.getValue();
  
  //  if (!mainSheetErrorMessageCellValue) {
  //          mainSheetErrorMessageCell.setValue("Error Alert Column");
  //         }
  
  // }
  
  // function setCommentsOnReportCell(reportSheetBinding, reportSummaryCommentsBinding, reportSummaryColumnPositionBinding) { //ALL SHEETS ADDED
  //   let reportSummaryCommentsString = reportSummaryCommentsBinding.join(", ");
  //  let row = 1;
   
  //  let reportCell = getSheetCell(reportSheetBinding, row, reportSummaryColumnPositionBinding);
  
  //  reportCell.setValue(`REPORT OVERVIEW`);
  //  insertCommentToSheetCell(reportCell, reportSummaryCommentsString);
  // }
  
  // // Function that executes when user clicks "User Template Check" Button
  // function createReportSheet(ssBinding, valuesBinding, reportName) { //ALL SHEETS ADDED
  // let reportSheetBinding = getSheet(reportName); 
  // deleteSheetIfExists(reportSheetBinding);
  // reportSheetBinding = createSheetCopy(ssBinding, reportName);
  // copyValuesToSheet(reportSheetBinding, valuesBinding);
  
  // return reportSheetBinding;
  // }
  
  
  /*
  
  END OF ALL TEMPLATE CLASS 
  
  
  
  */
  
  /*
  
  
  BEGINNING OF AGENCY, NEED, and USER CLASS
  
  
  
   */
  
  // const STATE_TWO_LETTER_CODE_COMMENT = 'State converted to two letter code'; // USER, PRORGAMS/AGENCIES, and NEED TEMPLATES
  // const INVALID_POSTAL_CODE_COMMENT = 'Invalid postal code'; // USERS, NEEDS, and AGENCIES TEMPLATE ONLY
  // const INVALID_PHONE_NUMBER_COMMENT = 'Invalid phone number'; //USERS and AGENCIES only
  // const INVALID_STATE_COMMENT = 'Invalid State'; // USERS, NEEDS, and AGENCIES ONLY
  // const US_STATE_TO_ABBREVIATION = { //USERS< NEEDS, and AGENCIES
  //   "Alabama": "AL",
  //   "Alaska": "AK",
  //   "Arizona": "AZ",
  //   "Arkansas": "AR",
  //   "California": "CA",
  //   "Colorado": "CO",
  //   "Connecticut": "CT",
  //   "Delaware": "DE",
  //   "Florida": "FL",
  //   "Georgia": "GA",
  //   "Hawaii": "HI",
  //   "Idaho": "ID",
  //   "Illinois": "IL",
  //   "Indiana": "IN",
  //   "Iowa": "IA",
  //   "Kansas": "KS",
  //   "Kentucky": "KY",
  //   "Louisiana": "LA",
  //   "Maine": "ME",
  //   "Maryland": "MD",
  //   "Massachusetts": "MA",
  //   "Michigan": "MI",
  //   "Minnesota": "MN",
  //   "Mississippi": "MS",
  //   "Missouri": "MO",
  //   "Montana": "MT",
  //   "Nebraska": "NE",
  //   "Nevada": "NV",
  //   "New Hampshire": "NH",
  //   "New Jersey": "NJ",
  //   "New Mexico": "NM",
  //   "New York": "NY",
  //   "North Carolina": "NC",
  //   "North Dakota": "ND",
  //   "Ohio": "OH",
  //   "Oklahoma": "OK",
  //   "Oregon": "OR",
  //   "Pennsylvania": "PA",
  //   "Rhode Island": "RI",
  //   "South Carolina": "SC",
  //   "South Dakota": "SD",
  //   "Tennessee": "TN",
  //   "Texas": "TX",
  //   "Utah": "UT",
  //   "Vermont": "VT",
  //   "Virginia": "VA",
  //   "Washington": "WA",
  //   "West Virginia": "WV",
  //   "Wisconsin": "WI",
  //   "Wyoming": "WY",
  //   "District of Columbia": "DC",
  //   "American Samoa": "AS",
  //   "Guam": "GU",
  //   "Northern Mariana Islands": "MP",
  //   "Puerto Rico": "PR",
  //   "United States Minor Outlying Islands": "UM",
  //   "U.S. Virgin Islands": "VI",
  // }
  // const usFullStateNames = Object.keys(US_STATE_TO_ABBREVIATION); // USERS, NEEDS, AGENCIES
  // const usStateAbbreviations = Object.values(US_STATE_TO_ABBREVIATION); // USERS, NEEDS, AGENCIES
  // const genderOptionsAbbreviations = Object.keys(GENDER_OPTIONS_ABBREVIATION_TO_FULL); //USERS, NEEDS, AGENCIES
  
  // function convertStatesToTwoLetterCode(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding) { //USERS, NEEDS, PROGRAMS/AGENCIES ADDED
  //   let stateColumnRange = getColumnRange('State', sheetBinding, columnsHeadersBinding);
  
  //   if (stateColumnRange) {
  //     let stateColumnRangeValues = getValues(stateColumnRange).map(val => {
  //     let currentState = String(val);
  //     if (currentState.length > 2) {
  //       return (currentState[0].toUpperCase() + currentState.substr(1).toLowerCase());
  //     } else {
  //       return currentState;
  //     }
  //   });
  
  //   // SpreadsheetApp.getUi().alert(`${stateColumnRangeValues}`); Testing
  
  //     let stateColumnRangePosition = stateColumnRange.getColumn();
      
  //     stateColumnRangeValues.forEach((val, index) => {
  //       let currentState = String(val);
  //       // SpreadsheetApp.getUi().alert(`${currentState}`); Testing
  //       // SpreadsheetApp.getUi().alert(`${currentState.length}`); Testing
  //       let row = index + 2;
  //       let currentCell = getSheetCell(sheetBinding, row, stateColumnRangePosition);
  //       let currentReportCell = getSheetCell(reportSheetBinding, row, stateColumnRangePosition);
  //       if (currentState.length > 2 && usFullStateNames.includes(currentState)) {
  //         currentCell.setValue(US_STATE_TO_ABBREVIATION[currentState]);
  //         setSheetCellBackground(currentReportCell, LIGHT_RED_HEX_CODE);
  //         insertCommentToSheetCell(currentReportCell, STATE_TWO_LETTER_CODE_COMMENT);
  
  //       }
  //   });
  
  //   reportSummaryCommentsBinding.push("Success: ran check to convert states to two-digit format");
  
  
  //   } else {
  //     reportSummaryCommentsBinding.push("Success: ran check to convert states to two-digit format, but did not find column");
  
  //   }
    
  
  // }
  
  // function validateStates(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding, reportSummaryColumnPositionBinding) { //USERS, NEEDS, PROGRAMS/AGENCIES ADDED
  //   let stateColumnRange = getColumnRange('State', sheetBinding, columnsHeadersBinding);
  
  //   if (stateColumnRange) {
  //     let stateColumnRangeValues = getValues(stateColumnRange);
  
  //     let stateColumnRangePosition = stateColumnRange.getColumn();
      
  //     stateColumnRangeValues.forEach((val, index) => {
  //       let currentState = String(val);
  //       // SpreadsheetApp.getUi().alert(`${currentState}`); Testing
  //       // SpreadsheetApp.getUi().alert(`${currentState.length}`); Testing
        
  //       if (!usStateAbbreviations.includes(currentState) && currentState.length !== 0) {
  //         let headerRow = 1;
  //         let row = index + 2;
  //         let currentCell = getSheetCell(sheetBinding, row, stateColumnRangePosition);
  //         let currentReportCell = getSheetCell(reportSheetBinding, row, stateColumnRangePosition);
  //         let mainSheetHeaderCell = getSheetCell(sheetBinding, headerRow, stateColumnRangePosition);
  //         setSheetCellBackground(currentCell, LIGHT_RED_HEX_CODE);
  //         setSheetCellBackground(currentReportCell, LIGHT_RED_HEX_CODE);
  //         insertCommentToSheetCell(currentReportCell, INVALID_STATE_COMMENT);
  //         setSheetCellBackground(mainSheetHeaderCell, LIGHT_RED_HEX_CODE);
  //         insertHeaderComment(mainSheetHeaderCell, "Invalid State/States found");
  //         setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
  
  //       }
  //   });
  
  //   reportSummaryCommentsBinding.push("Success: ran check for invalid states");
  
  
  //   } else {
  //     reportSummaryCommentsBinding.push("Success: ran check for invalid states, but did not find column");
  
  //   }
    
  
  // }
  
  
  // const validatePhoneNumbers = (number) => { //USERS AND AGENCIES/PROGRAMS ADDED
  // return String(number)
  //   .toLowerCase()
  //   .match(
  // /^[\+]?[(]?[0-9]{3}[)]?[-\s\.]?[0-9]{3}[-\s\.]?[0-9]{4,6}$/im
  //   );
  // };
  
  // const validatePostalCode = (postalCode) => { //USERS, NEEDS, AGENCIES/PROGRAMS ADDED
  // return /(^\d{5}$)|(^\d{5}-\d{4}$)/.test(postalCode);
  // }
  // // Consider this match: /^w+([.-]?w+)*@w+([.-]?w+)*(.w{2,3})+$/;
  
  // function checkForInvalidNumbers(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding, reportSummaryColumnPositionBinding) { //USERS AND AGENCIES
  //   let headerRow = 1;
  //   let homePhoneNumberRange = getColumnRange('Phone', sheetBinding, columnsHeadersBinding);
  //   let cellPhoneNumberRange = getColumnRange('Mobile', sheetBinding, columnsHeadersBinding);
  
  //   if (homePhoneNumberRange) {
  //     let homePhoneNumberRangeValues = getValues(homePhoneNumberRange);
  //     let homePhoneNumberRangePosition = homePhoneNumberRange.getColumn();
  //     let mainSheetHomePhoneHeaderCell = getSheetCell(sheetBinding, headerRow, homePhoneNumberRangePosition);
  
  //     homePhoneNumberRangeValues.forEach((number, index) => {
  
  //       let currentNumber= String(number);
  //       let row = index + 2;
  
  //       if (currentNumber !== "" && !validatePhoneNumbers(currentNumber)) {
  //         let currentCell = getSheetCell(sheetBinding, row, homePhoneNumberRangePosition);
  //         let reportSheetCell = getSheetCell(reportSheetBinding, row, homePhoneNumberRangePosition);
  
  //         //Line below for testing
  
  //         // SpreadsheetApp.getUi().alert(`Invalid Email! ${currentEmail} ${reportSheetCell.getA1Notation()}`);
  //         //testing line ends here
  //         setSheetCellBackground(reportSheetCell, LIGHT_RED_HEX_CODE);
  //         setSheetCellBackground(currentCell, LIGHT_RED_HEX_CODE);
  //         setSheetCellBackground(mainSheetHomePhoneHeaderCell, LIGHT_RED_HEX_CODE);
  //         insertCommentToSheetCell(reportSheetCell, INVALID_PHONE_NUMBER_COMMENT);
  //         insertHeaderComment(mainSheetHomePhoneHeaderCell, "Invalid Home Phone Number/Numbers Found");
  //         setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
  //         }
  //       }
        
  //     );
  
  //     reportSummaryCommentsBinding.push("Success: ran check for invalid home phone numbers");
  
  //   } else {
  //     reportSummaryCommentsBinding.push("Success: ran check for invalid home phone numbers, but did not find column");
  //   }
  
  //   if (cellPhoneNumberRange) {
  //     let cellPhoneNumberRangeValues = getValues(cellPhoneNumberRange);
  //     let cellPhoneNumberRangePosition = cellPhoneNumberRange.getColumn();
  //     let mainSheetCellPhoneHeaderCell = getSheetCell(sheetBinding, headerRow, cellPhoneNumberRangePosition);
  
  //     cellPhoneNumberRangeValues.forEach((number, index) => {
  
  //     let currentNumber= String(number);
  //     let row = index + 2;
  
  //       if (currentNumber !== "" && !validatePhoneNumbers(currentNumber)) {
  //         let currentCell = getSheetCell(sheetBinding, row, cellPhoneNumberRangePosition);
  //         let reportSheetCell = getSheetCell(reportSheetBinding, row, cellPhoneNumberRangePosition);
  
  //         //Line below for testing
  
  //         // SpreadsheetApp.getUi().alert(`Invalid Email! ${currentEmail} ${reportSheetCell.getA1Notation()}`);
  //         //testing line ends here
  //         setSheetCellBackground(reportSheetCell, LIGHT_RED_HEX_CODE);
  //         setSheetCellBackground(currentCell, LIGHT_RED_HEX_CODE);
  //         setSheetCellBackground(mainSheetCellPhoneHeaderCell, LIGHT_RED_HEX_CODE);
  //         insertCommentToSheetCell(reportSheetCell, INVALID_PHONE_NUMBER_COMMENT);
  //         insertHeaderComment(mainSheetCellPhoneHeaderCell, "Invalid Cell Phone Number/Numbers Found");
  //         setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
  //         }
  //       }
  //     );
  
  //     reportSummaryCommentsBinding.push("Success: ran check for invalid mobile phone numbers");
  
  //   } else {
  //     reportSummaryCommentsBinding.push("Success: ran check for invalid mobile phone numbers, but did not find column");
  //   }
  
  // }
  
  // function checkForInvalidPostalCodes(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding, reportSummaryColumnPositionBinding) { //USERS, NEEDS, PROGRAMS
  //   let headerRow = 1;
  //   let postalCodeColumnRange = getColumnRange('Zip', sheetBinding, columnsHeadersBinding);
  
  //   if (postalCodeColumnRange) {
  //     let postalCodeColumnRangeValues = getValues(postalCodeColumnRange);
  //     let postalCodeColumnRangePosition = postalCodeColumnRange.getColumn();
  //     let mainSheetPostalHeaderCell = getSheetCell(sheetBinding, headerRow, postalCodeColumnRangePosition);
  
  //     postalCodeColumnRangeValues.forEach((code, index) => {
  //       let currentCode = String(code);
  //       let row = index + 2;
  
  //       if (currentCode !== "" && !validatePostalCode(currentCode)) {
  //         let currentCell = getSheetCell(sheetBinding, row, postalCodeColumnRangePosition);
  //         let reportSheetCell = getSheetCell(reportSheetBinding, row, postalCodeColumnRangePosition);
  
  //         //Line below for testing
  
  //         // SpreadsheetApp.getUi().alert(`Invalid Email! ${currentEmail} ${reportSheetCell.getA1Notation()}`);
  //         //testing line ends here
  //         setSheetCellBackground(reportSheetCell, LIGHT_RED_HEX_CODE);
  //         setSheetCellBackground(currentCell, LIGHT_RED_HEX_CODE);
  //         setSheetCellBackground(mainSheetPostalHeaderCell, LIGHT_RED_HEX_CODE);
  //         insertCommentToSheetCell(reportSheetCell, INVALID_POSTAL_CODE_COMMENT);
  //         insertHeaderComment(mainSheetPostalHeaderCell, "Invalid Postal Code/Codes Found");
  //         setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
  //         }
  //       }
  //     );
  
  //     reportSummaryCommentsBinding.push("Success: ran check for invalid postal codes");
  //   } else {
  //     reportSummaryCommentsBinding.push("Success: ran check for invalid postal codes, but did not find column");
  //   }
    
  // }
  
  
  // /*
  
  // END OF USER, NEEDS, AGENCIES CLASS
  
  //  */
  
  
  // /*
  
  // BEGINNING OF USERS CLASS
  
  //  */
  
  
  
  
  
  
  // const DUPLICATE_EMAIL_COMMENT = 'Duplicate email'; //User Import template
  // // const FIRST_NAME_MISSING_COMMENT = 'First name missing'; //user import template DEP
  // // const LAST_NAME_MISSING_COMMENT = 'Last name missing'; // user import template DEP
  // // const EMAIL_MISSING_COMMENT = 'Email missing'; user import teamplate DEP
  
  
  // const GENDER_OPTION_EXPANDED_COMMENT = 'Gender option expanded'; // User import template only
  // // const CAPITALIZATION_FUNCTION_RAN_ON_FN_COMMENT = 'Capitalization function ran on first name column'; NOT IN USE
  // // const CAPITALIZATION_FUNCTION_RAN_ON_LN_COMMENT = 'Capitalization function ran on last name column';
  
  
  
  
  
  // const INVALID_GENDER_OPTION = 'Invalid Gender Option'; // USERS ONLY
  // const GENDER_OPTIONS_ABBREVIATION_TO_FULL = { //USERS
  //   "F":"Female",
  //   "M":"Male",
  //   "N/A":"",
  // }
  // const FULL_GENDER_OPTIONS = ['female', 'male','prefer not to say', 'other']; //USERS
  
  
  
  
  
  
  // function checkForDuplicateEmails(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding, reportSummaryColumnPositionBinding) { //USER TEMP ONLY
  //   let headerRow = 1;
  //   let emailColumnRange = getColumnRange('Email', sheetBinding, columnsHeadersBinding);
  //   let reportSheetEmailColumnRange = getColumnRange('Email', reportSheetBinding, columnsHeadersBinding);
  //   // let emailColumnValues = getValues(emailColumnRange).map(email => email[0].toLowerCase().trim());
  //   let emailColumnPosition = emailColumnRange.getColumn();
  //   let mainSheetHeaderCell = getSheetCell(sheetBinding, headerRow, emailColumnPosition);
  //   let beginningOfEmailColumnRange = emailColumnRange.getA1Notation().slice(0,2); //0,2 represents beginning cell for column range i.e.C1
  //   let endOfEmailColumnRange = emailColumnRange.getA1Notation().slice(3); // slice 3 represents the end of the column range excluding the comma
  //   // SpreadsheetApp.getUi().alert(`${beginningOfEmailColumnRange}: ${endOfEmailColumnRange}`); for tsting
  //   const duplicateEmailsRuleMainSheet = SpreadsheetApp.newConditionalFormatRule()
  //   .whenFormulaSatisfied(`=COUNTIF(${beginningOfEmailColumnRange}:${endOfEmailColumnRange}, ${beginningOfEmailColumnRange})>1`)
  //   .setBackground(LIGHT_RED_HEX_CODE)
  //   // .insertHeaderComment(mainSheetHeaderCell, "Duplicate Emails Found")
  //   .setRanges([emailColumnRange])
  //   .build();
  //   let rules = sheetBinding.getConditionalFormatRules();
  //   rules.push(duplicateEmailsRuleMainSheet);
  //   sheetBinding.setConditionalFormatRules(rules);
  
  //   const duplicateEmailsRuleReportSheet = SpreadsheetApp.newConditionalFormatRule()
  //   .whenFormulaSatisfied(`=COUNTIF(${beginningOfEmailColumnRange}:${endOfEmailColumnRange}, ${beginningOfEmailColumnRange})>1`)
  //   .setBackground(LIGHT_RED_HEX_CODE)
  //   // .setComment(DUPLICATE_EMAIL_COMMENT)
  //   .setRanges([reportSheetEmailColumnRange])
  //   .build();
  
  //   let rules2 = reportSheetBinding.getConditionalFormatRules();
  //   rules2.push(duplicateEmailsRuleReportSheet);
  //   reportSheetBinding.setConditionalFormatRules(rules2);
  
  //   // let duplicates = [];
  
  //   // emailColumnValues.forEach((email, index) => {
  //   //   let currentEmail = String(email);
  //   //   let row = index + 2;
  
  //   //     if (duplicates.indexOf(currentEmail) === -1 && currentEmail.length > 0) {
  //   //      if (emailColumnValues.filter(val => String(val) === currentEmail).length > 1) {
  //   //       let currentCell = getSheetCell(sheetBinding, row, emailColumnPosition);
  //   //       let reportSheetCell = getSheetCell(reportSheetBinding, row, emailColumnPosition);
  //   //       let mainSheetHeaderCell = getSheetCell(sheetBinding, headerRow, emailColumnPosition);
  
  
  //   //       //Line below for testing
  
  //   //       // SpreadsheetApp.getUi().alert(`Duplicate found! ${currentEmail} ${reportSheetCell.getA1Notation()}`);
  //   //       //testing line ends here
  //   //       setSheetCellBackground(reportSheetCell, LIGHT_RED_HEX_CODE);
  //   //       setSheetCellBackground(currentCell, LIGHT_RED_HEX_CODE);
  //   //       setSheetCellBackground(mainSheetHeaderCell, LIGHT_RED_HEX_CODE);
  //   //       insertCommentToSheetCell(reportSheetCell, DUPLICATE_EMAIL_COMMENT);
  //   //       insertHeaderComment(mainSheetHeaderCell, "Duplicate Emails Found");
  //   //       setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
  //   //       duplicates.push(currentEmail);
  //   //       }
  //   //     }
  //   //     }
  //     // );
  
  //   reportSheetBinding.sort(emailColumnPosition);
  //   sheetBinding.sort(emailColumnPosition);
  //   reportSummaryCommentsBinding.push("Success: checked emails column for duplicates");
  
  // }
  
  // function createArrayOfNamesAndEmail(firstNameRangeBinding, lastNameRangeBinding, emailColmnRangeBinding) { //No longer in user
  //   let firstNameRangeValues = getValues(firstNameRangeBinding);
  //   let lastNameRangeValues = getValues(lastNameRangeBinding);
  //   let emailColumnRangeValues = getValues(emailColmnRangeBinding);
  //   let firstNameLastNameEmailValueArrayBinding = [];
  //   for (let i = 0; i < emailColumnRangeValues.length; i += 1) {
  //   firstNameLastNameEmailValueArrayBinding.push([firstNameRangeValues[i], lastNameRangeValues[i], emailColumnRangeValues[i]]);
  //   }
  
  //    return firstNameLastNameEmailValueArrayBinding;
  // }
  
  
  // // function checkForMissingNamesOrEmails(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding, reportSummaryColumnPositionBinding) {
  // //   let headerRow = 1;
  // //   let firstNameRange = getColumnRange('First Name', sheetBinding, columnsHeadersBinding);
  // //   let firstNameRangePosition = firstNameRange.getColumn();
  // //   let lastNameRange = getColumnRange('Last Name', sheetBinding, columnsHeadersBinding);
  // //   let lastNameRangePosition = lastNameRange.getColumn();
  // //   let emailColumnRange = getColumnRange('Email', sheetBinding, columnsHeadersBinding);
  // //   let emailColumnPosition = emailColumnRange.getColumn();
  // //   let firstNameLastNameEmailValueArray = createArrayOfNamesAndEmail(firstNameRange, lastNameRange, emailColumnRange);
  
  // //   for (let i = 0; i < firstNameLastNameEmailValueArray.length; i +=1) {
  
  // //     let row = i + 2;
  // //     let firstName = String(firstNameLastNameEmailValueArray[i][0]);
  // //     let firstNameCurrentCell = getSheetCell(sheetBinding, row, firstNameRangePosition);
  // //     let mainSheetFNHeaderCell = getSheetCell(sheetBinding, headerRow, firstNameRangePosition);
  // //     let reportSheetCurrentFNCell = getSheetCell(reportSheetBinding, row, firstNameRangePosition);
  // //     let lastName = String(firstNameLastNameEmailValueArray[i][1]);
  // //     let lastNameCurrentCell = getSheetCell(sheetBinding, row, lastNameRangePosition);
  // //     let mainSheetLNHeaderCell = getSheetCell(sheetBinding, headerRow, lastNameRangePosition);
  // //     let reportSheetLNCurrentCell = getSheetCell(reportSheetBinding, row, lastNameRangePosition);
  // //     let email = String(firstNameLastNameEmailValueArray[i][2]);
  // //     let emailCurrentCell = getSheetCell(sheetBinding, row, emailColumnPosition);
  // //     let mainSheetEmailHeaderCell = getSheetCell(sheetBinding, headerRow, emailColumnPosition);
  // //     let reportSheetEmailCurrentCell = getSheetCell(reportSheetBinding, row, emailColumnPosition);
  
  // //     if (firstName.length !== 0 || lastName.length !== 0 || email.length !== 0) {
  // //       firstNameLastNameEmailValueArray[i].forEach((val, index) => {
  // //       let currentCellValue = String(val);
  // //       if (currentCellValue === "") {
  // //         switch(index) {
  // //           case 0: 
  // //             setSheetCellBackground(reportSheetCurrentFNCell, LIGHT_RED_HEX_CODE);
  // //             setSheetCellBackground(firstNameCurrentCell, LIGHT_RED_HEX_CODE);
  // //             setSheetCellBackground(mainSheetFNHeaderCell, LIGHT_RED_HEX_CODE);
  // //             insertCommentToSheetCell(reportSheetCurrentFNCell, FIRST_NAME_MISSING_COMMENT);
  // //             setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
  // //             insertHeaderComment(mainSheetFNHeaderCell, "First Name/Names Missing");
  // //             break;
  // //           case 1:
  // //             setSheetCellBackground(reportSheetLNCurrentCell, LIGHT_RED_HEX_CODE);
  // //             setSheetCellBackground(lastNameCurrentCell, LIGHT_RED_HEX_CODE);
  // //             setSheetCellBackground(mainSheetLNHeaderCell, LIGHT_RED_HEX_CODE);
  // //             insertCommentToSheetCell(reportSheetLNCurrentCell, LAST_NAME_MISSING_COMMENT);
  // //             setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
  // //             insertHeaderComment(mainSheetLNHeaderCell, "Last Name/Names Missing");
  // //             break;
  // //           case 2:
  // //             setSheetCellBackground(reportSheetEmailCurrentCell, LIGHT_RED_HEX_CODE);
  // //             setSheetCellBackground(emailCurrentCell, LIGHT_RED_HEX_CODE);
  // //             setSheetCellBackground(mainSheetEmailHeaderCell, LIGHT_RED_HEX_CODE);
  // //             insertCommentToSheetCell(reportSheetEmailCurrentCell, EMAIL_MISSING_COMMENT);
  // //             setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
  // //             insertHeaderComment(mainSheetEmailHeaderCell, "Email/Emails Missing");
  // //             break;
  // //         }
  // //       }
  // //     })
  // //   }
  // // }
  // // reportSummaryCommentsBinding.push("Success: checked for missing names and emails");
  // // }
  
  // function checkFirstThreeColumnsForBlanks(sheetBinding, reportSheetBinding, reportSummaryCommentsBinding, reportSummaryColumnPositionBinding) { //User only and each template has its own variant
  //   let rowStartPosition = 2;
  //   let columnStartPosition = 1;
  //   let middleColumnPosition = 2;
  //   let thirdColumnPositon = 3;
  //   let maxRows = sheetBinding.getMaxRows();
  //   let totalColumsToCheck = 3;
  //   let range = sheetBinding.getRange(rowStartPosition, columnStartPosition, maxRows, totalColumsToCheck);
  //   let values = range.getValues();
  
  //   for (let row = 0; row < values.length; row += 1) {
  //     let headerRowPosition = 1;
  //     let cellRow = row + 2;
  //     let currentRow = values[row];
  //     let item1 = currentRow[0];
  //     let item1CurrentCell = getSheetCell(sheetBinding, cellRow, columnStartPosition);
  //     let item1ReportCurrentCell = getSheetCell(reportSheetBinding, cellRow, columnStartPosition);
  //     let item1HeaderCell = getSheetCell(sheetBinding, headerRowPosition, columnStartPosition);
  //     let item2 = currentRow[1];
  //     let item2CurrentCell = getSheetCell(sheetBinding, cellRow, middleColumnPosition);
  //     let item2ReportCurrentCell = getSheetCell(reportSheetBinding, cellRow, middleColumnPosition);
  //     let item2HeaderCell = getSheetCell(sheetBinding, headerRowPosition, middleColumnPosition);
  //     let item3 = currentRow[2];
  //     let item3CurrentCell = getSheetCell(sheetBinding, cellRow, thirdColumnPositon);
  //     let item3ReportCell = getSheetCell(reportSheetBinding, cellRow, thirdColumnPositon);
  //     let item3HeaderCell = getSheetCell(sheetBinding, headerRowPosition, thirdColumnPositon);
  
  //     if (item1.length !== 0 || item2.length !== 0  || item3.length !== 0) {
  //       currentRow.forEach((val, index) => {
  //         if (val === "") {
  //           switch (index) {
  //             case 0:
  //               setSheetCellBackground(item1CurrentCell, LIGHT_RED_HEX_CODE);
  //               setSheetCellBackground(item1ReportCurrentCell, LIGHT_RED_HEX_CODE);
  //               setSheetCellBackground(item1HeaderCell, LIGHT_RED_HEX_CODE);
  //               insertCommentToSheetCell(item1ReportCurrentCell, VALUES_MISSING_COMMENT);
  //               setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, cellRow);
  //               insertHeaderComment(item1HeaderCell, VALUES_MISSING_COMMENT);
  //               break;
  //             case 1:
  //               setSheetCellBackground(item2CurrentCell, LIGHT_RED_HEX_CODE);
  //               setSheetCellBackground(item2ReportCurrentCell, LIGHT_RED_HEX_CODE);
  //               setSheetCellBackground(item2HeaderCell, LIGHT_RED_HEX_CODE);
  //               insertCommentToSheetCell(item2ReportCurrentCell, VALUES_MISSING_COMMENT);
  //               setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, cellRow);
  //               insertHeaderComment(item2HeaderCell, VALUES_MISSING_COMMENT);
  //               break;
  //             case 2:
  //               setSheetCellBackground(item3CurrentCell, LIGHT_RED_HEX_CODE);
  //               setSheetCellBackground(item3ReportCell, LIGHT_RED_HEX_CODE);
  //               setSheetCellBackground(item3HeaderCell, LIGHT_RED_HEX_CODE);
  //               insertCommentToSheetCell(item3ReportCell, VALUES_MISSING_COMMENT);
  //               setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, cellRow);
  //               insertHeaderComment(item3HeaderCell, VALUES_MISSING_COMMENT);
  //               break;
  //           }
  //         }
  //       });
  //     }
  //   }
  
  // reportSummaryCommentsBinding.push("Success: checked for missing values in first three columns");
  
  // }
  
  
  // //  function checkForMissingNamesOrEmails(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding, reportSummaryColumnPositionBinding) {
  // //     let headerRow = 1;
  // //     let firstNameRange = getColumnRange('First Name', sheetBinding, columnsHeadersBinding);
  // //     let firstNameRangePosition = firstNameRange.getColumn();
  // //     let lastNameRange = getColumnRange('Last Name', sheetBinding, columnsHeadersBinding);
  // //     let lastNameRangePosition = lastNameRange.getColumn();
  // //     let emailColumnRange = getColumnRange('Email', sheetBinding, columnsHeadersBinding);
  // //     let emailColumnPosition = emailColumnRange.getColumn();
  // //     let firstNameLastNameEmailValueArray = createArrayOfNamesAndEmail(firstNameRange, lastNameRange, emailColumnRange);
  
  // //     for (let i = 0; i < firstNameLastNameEmailValueArray.length; i +=1) {
  
  // //       let row = i + 2;
  // //       let firstName = String(firstNameLastNameEmailValueArray[i][0]);
  // //       let firstNameCurrentCell = getSheetCell(sheetBinding, row, firstNameRangePosition);
  // //       let mainSheetFNHeaderCell = getSheetCell(sheetBinding, headerRow, firstNameRangePosition);
  // //       let reportSheetCurrentFNCell = getSheetCell(reportSheetBinding, row, firstNameRangePosition);
  // //       let lastName = String(firstNameLastNameEmailValueArray[i][1]);
  // //       let lastNameCurrentCell = getSheetCell(sheetBinding, row, lastNameRangePosition);
  // //       let mainSheetLNHeaderCell = getSheetCell(sheetBinding, headerRow, lastNameRangePosition);
  // //       let reportSheetLNCurrentCell = getSheetCell(reportSheetBinding, row, lastNameRangePosition);
  // //       let email = String(firstNameLastNameEmailValueArray[i][2]);
  // //       let emailCurrentCell = getSheetCell(sheetBinding, row, emailColumnPosition);
  // //       let mainSheetEmailHeaderCell = getSheetCell(sheetBinding, headerRow, emailColumnPosition);
  // //       let reportSheetEmailCurrentCell = getSheetCell(reportSheetBinding, row, emailColumnPosition);
  
  // //       if (firstName.length !== 0 || lastName.length !== 0 || email.length !== 0) {
  // //         firstNameLastNameEmailValueArray[i].forEach((val, index) => {
  // //         let currentCellValue = String(val);
  // //         if (currentCellValue === "") {
  // //           switch(index) {
  // //             case 0: 
  // //               setSheetCellBackground(reportSheetCurrentFNCell, LIGHT_RED_HEX_CODE);
  // //               setSheetCellBackground(firstNameCurrentCell, LIGHT_RED_HEX_CODE);
  // //               setSheetCellBackground(mainSheetFNHeaderCell, LIGHT_RED_HEX_CODE);
  // //               insertCommentToSheetCell(reportSheetCurrentFNCell, FIRST_NAME_MISSING_COMMENT);
  // //               setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
  // //               insertHeaderComment(mainSheetFNHeaderCell, "First Name/Names Missing");
  // //               break;
  // //             case 1:
  // //               setSheetCellBackground(reportSheetLNCurrentCell, LIGHT_RED_HEX_CODE);
  // //               setSheetCellBackground(lastNameCurrentCell, LIGHT_RED_HEX_CODE);
  // //               setSheetCellBackground(mainSheetLNHeaderCell, LIGHT_RED_HEX_CODE);
  // //               insertCommentToSheetCell(reportSheetLNCurrentCell, LAST_NAME_MISSING_COMMENT);
  // //               setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
  // //               insertHeaderComment(mainSheetLNHeaderCell, "Last Name/Names Missing");
  // //               break;
  // //             case 2:
  // //               setSheetCellBackground(reportSheetEmailCurrentCell, LIGHT_RED_HEX_CODE);
  // //               setSheetCellBackground(emailCurrentCell, LIGHT_RED_HEX_CODE);
  // //               setSheetCellBackground(mainSheetEmailHeaderCell, LIGHT_RED_HEX_CODE);
  // //               insertCommentToSheetCell(reportSheetEmailCurrentCell, EMAIL_MISSING_COMMENT);
  // //               setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
  // //               insertHeaderComment(mainSheetEmailHeaderCell, "Email/Emails Missing");
  // //               break;
  // //           }
  // //         }
  // //       })
  // //     }
  // //   }
  // //   reportSummaryCommentsBinding.push("Success: checked for missing names and emails");
  // // }
  
  
  // /*
  
  
  
  // CATCH UP HERE FOR ORGANIZATION PLANNING
  
  
  //  */
  // function formatUserDateAddedAndBirthdayColumns(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding) { //USER IMPORT TEMPLATE ONLY
  // let row = 1;
  // let birthdayColumn = getColumnRange('Birthday (YYYY-MM-DD)', sheetBinding, columnsHeadersBinding);
  // let userDateAddedColumn = getColumnRange('User Date Added', sheetBinding, columnsHeadersBinding);
  
  // if (birthdayColumn && userDateAddedColumn) {
  //   let birthdayColumnPosition = birthdayColumn.getColumn();
  //   let reportSheetBirthdayColumnHeaderCell = getSheetCell(reportSheetBinding, row, birthdayColumnPosition);
  //   let userDateAddedColumnPosition = userDateAddedColumn.getColumn();
  //   let reportSheetUserDateAddedColumnHeaderCell = getSheetCell(reportSheetBinding, row, userDateAddedColumnPosition);
  
  //   setColumnToYYYYMMDDFormat(birthdayColumn);
  //   setColumnToYYYYMMDDFormat(userDateAddedColumn);
  //   setSheetCellBackground(reportSheetBirthdayColumnHeaderCell, LIGHT_GREEN_HEX_CODE);
  //   insertCommentToSheetCell(reportSheetBirthdayColumnHeaderCell, DATE_FORMATTED_YYYY_MM_DD_COMMENT);
  //   setSheetCellBackground(reportSheetUserDateAddedColumnHeaderCell, LIGHT_GREEN_HEX_CODE);
  //   insertCommentToSheetCell(reportSheetUserDateAddedColumnHeaderCell, DATE_FORMATTED_YYYY_MM_DD_COMMENT);
  //   reportSummaryCommentsBinding.push("Success: formatted birthday column and user date added column");
  // } else if (birthdayColumn) {
  //   let birthdayColumnNumberFormat = birthdayColumn.getNumberFormat();
  //   let birthdayColumnPosition = birthdayColumn.getColumn();
  //   let reportSheetBirthdayColumnHeaderCell = getSheetCell(reportSheetBinding, row, birthdayColumnPosition);
  //   setColumnToYYYYMMDDFormat(birthdayColumn);
  
  //   setSheetCellBackground(reportSheetBirthdayColumnHeaderCell, LIGHT_GREEN_HEX_CODE);
  //   insertCommentToSheetCell(reportSheetBirthdayColumnHeaderCell, DATE_FORMATTED_YYYY_MM_DD_COMMENT); 
  //   reportSummaryCommentsBinding.push("Success: formatted birthday column and did not find user date added column");
  // } else if (userDateAddedColumn) {
  //   let userDateAddedColumnNumberFormat = userDateAddedColumn.getNumberFormat();
  //   let userDateAddedColumnPosition = userDateAddedColumn.getColumn();
  //   let reportSheetUserDateAddedColumnHeaderCell = getSheetCell(reportSheetBinding, row, userDateAddedColumnPosition);
  
  //   setColumnToYYYYMMDDFormat(userDateAddedColumn);
  //   setSheetCellBackground(reportSheetUserDateAddedColumnHeaderCell, LIGHT_GREEN_HEX_CODE);
  //   insertCommentToSheetCell(reportSheetUserDateAddedColumnHeaderCell, DATE_FORMATTED_YYYY_MM_DD_COMMENT);
  //   reportSummaryCommentsBinding.push("Success: formatted user date added column and did not find birthday column")
  // }
  
  // }
  
  // function convertGenderOptionAbbreviationToFullWord(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding) { //No longer needed so DEP, but USERS ONLY
  //   let genderOptionColumnRange = getColumnRange('Gender', sheetBinding, columnsHeadersBinding);
  
  //   if (genderOptionColumnRange) {
  //     let genderOptionColumnRangeValues = getValues(genderOptionColumnRange).map(val => {
  //       let currentGenderOption = String(val);
  //       if (currentGenderOption.length === 1) return currentGenderOption.toUpperCase();
  //     });
  
  //     let genderOptionColumnRangePosition = genderOptionColumnRange.getColumn();
      
  //     genderOptionColumnRangeValues.forEach((val, index) => {
  //       let currentGenderOption = String(val);
  //       // SpreadsheetApp.getUi().alert(`${currentState}`); Testing
  //       // SpreadsheetApp.getUi().alert(`${currentState.length}`); Testing
  //       let row = index + 2;
  //       let currentCell = getSheetCell(sheetBinding, row, genderOptionColumnRangePosition);
  //       let currentReportCell = getSheetCell(reportSheetBinding, row, genderOptionColumnRangePosition);
  //       if (genderOptionsAbbreviations.includes(currentGenderOption)) {
  //         currentCell.setValue(GENDER_OPTIONS_ABBREVIATION_TO_FULL[currentGenderOption]);
  //         setSheetCellBackground(currentReportCell, LIGHT_RED_HEX_CODE);
  //         insertCommentToSheetCell(currentReportCell, GENDER_OPTION_EXPANDED_COMMENT);
  
  //       }
  //   });
  
  //   reportSummaryCommentsBinding.push("Success: ran check to expand gender option abbrevitation to full option (i.e. F => Female)");
  
  
  //   } else {
  //     reportSummaryCommentsBinding.push("Success: ran check to expand gender option abbrevitation to full option (i.e. F => Female), but did not find column");
  
  //   }
    
  
  // }
  
  // function validateGenderOptions(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding, reportSummaryColumnPositionBinding) { //USERS ONLY
  //   let genderOptionColumnRange = getColumnRange('Gender', sheetBinding, columnsHeadersBinding);
  
  //   if (genderOptionColumnRange) {
  //     let genderOptionColumnRangeValues = getValues(genderOptionColumnRange);
  
  //     let genderOptionColumnRangePoition = genderOptionColumnRange.getColumn();
      
  //     genderOptionColumnRangeValues.forEach((val, index) => {
  //       let currentGenderOption = String(val).toLowerCase();
  //       // SpreadsheetApp.getUi().alert(`${currentState}`); Testing
  //       // SpreadsheetApp.getUi().alert(`${currentState.length}`); Testing
        
  //       if (!FULL_GENDER_OPTIONS.includes(currentGenderOption) && currentGenderOption.length !== 0) {
  //         let headerRow = 1;
  //         let row = index + 2;
  //         let currentCell = getSheetCell(sheetBinding, row, genderOptionColumnRangePoition);
  //         let currentReportCell = getSheetCell(reportSheetBinding, row, genderOptionColumnRangePoition);
  //         let mainSheetHeaderCell = getSheetCell(sheetBinding, headerRow, genderOptionColumnRangePoition);
  //         setSheetCellBackground(currentCell, LIGHT_RED_HEX_CODE);
  //         setSheetCellBackground(currentReportCell, LIGHT_RED_HEX_CODE);
  //         insertCommentToSheetCell(currentReportCell, INVALID_GENDER_OPTION);
  //         setSheetCellBackground(mainSheetHeaderCell, LIGHT_RED_HEX_CODE);
  //         insertHeaderComment(mainSheetHeaderCell, "Invalid Gender Option/Options found");
  //         setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
  
  //       }
  //   });
  
  //   reportSummaryCommentsBinding.push("Success: ran check for invalid gender options");
  
  
  //   } else {
  //     reportSummaryCommentsBinding.push("Success: ran check for invalid gender options, but did not find column");
  
  //   }
    
  
  // }
  
  
  
  // function capitalizeFirstLetterOfAName(name) { //not in use
  //   return name.split(" ").filter(name => name.length > 0 && name).map(name => name.trim()[0].toUpperCase() + name.trim().substr(1).toLowerCase()).join(" ").split("-").map(name => name.trim()).filter(name => name !== "-" && name.length > 0).map(name => (name.split(" ").length > 1) ? name : name.trim()[0].toUpperCase() + name.trim().substr(1).toLowerCase()).join("-");
  // }
  
  // //Not using this function for now, but saving for later or personal use
  // // function capitalizeFirstLetterOfWords(sheetBinding, reportSheetBinding, columnsHeadersBinding) {
  // //   let firstNameRange = getColumnRange('First Name', sheetBinding, columnsHeadersBinding);
  // //   let firstNameColumnPosition = firstNameRange.getColumn();
  // //   let lastNameRange = getColumnRange('Last Name', sheetBinding, columnsHeadersBinding);
  // //   let lastNameRangeColumnPosition = lastNameRange.getColumn();
  // //   let firstNameRangeValues = getValues(firstNameRange);
  // //   let lastNameRangeValues = getValues(lastNameRange);
  // //   let reportSheetFirstNameHeaderCell = getSheetCell(reportSheetBinding, 1, firstNameColumnPosition);
  // //   let reportSheetLastNameHeaderCell = getSheetCell(reportSheetBinding, 1, lastNameRangeColumnPosition);
  
  // //   firstNameRangeValues.forEach((name, index) => {
  // //     let currentName = String(name);
  // //     if (currentName.length > 0) {
  // //       let row = index + 2;
  // //       let currentCell = getSheetCell(sheetBinding, row, firstNameColumnPosition);
  // //       currentCell.setValue(capitalizeFirstLetterOfAName(currentName));
      
  // //     }
    
  // //   });
  
  // //   lastNameRangeValues.forEach((name, index) => {
  // //     let currentName = String(name);
  
  // //     if (currentName.length > 0) {
  // //       let row = index + 2;
  // //       let currentCell = getSheetCell(sheetBinding, row, lastNameRangeColumnPosition);
  // //       currentCell.setValue(capitalizeFirstLetterOfAName(currentName));
      
  // //     }
    
  // //   });
  
  // //   setSheetCellBackground(reportSheetFirstNameHeaderCell, LIGHT_GREEN_HEX_CODE);
  // //   insertCommentToSheetCell(reportSheetFirstNameHeaderCell, CAPITALIZATION_FUNCTION_RAN_ON_FN_COMMENT);
  // //   setSheetCellBackground(reportSheetLastNameHeaderCell, LIGHT_GREEN_HEX_CODE);
  // //   insertCommentToSheetCell(reportSheetLastNameHeaderCell, CAPITALIZATION_FUNCTION_RAN_ON_LN_COMMENT);
  
  // // } 
  
  try {
    function checkUserImportTemplate() {
    let userImportTemplate = new UserTemplate();
    
    let reportSheet = userImportTemplate.createReportSheet();
    
  
    userImportTemplate.ss.setActiveSheet(userImportTemplate.getSheet(userImportTemplate.templateName));
   
  
    userImportTemplate.setFrozenRows(userImportTemplate.sheet, 1);
    userImportTemplate.setFrozenRows(reportSheet, 1);
  
    
    try {
      userImportTemplate.removeWhiteSpaceFromCells();
    } catch (err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = userImportTemplate.failedRemovedWhiteSpaceFromCellsComment;
      throw new Error(`White space not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
  
    try {
      userImportTemplate.removeFormattingFromSheetCells();
    } catch (err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = userImportTemplate.failedRemovedFormattingFromCellsComment;
      throw new Error(`Formatting not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
  
    userImportTemplate.clearSheetSummaryColumn(userImportTemplate.sheet);
    userImportTemplate.clearSheetSummaryColumn(reportSheet);
    userImportTemplate.setErrorColumnHeaderInMainSheet();
    
    try {
      userImportTemplate.checkForDuplicateEmails(reportSheet);
    } catch(err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = "Failed: check emails column for duplicates"; //User Only
      throw new Error(`Emails not checked for duplicates. Reason: ${err.name}: ${err.message}. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
  
    try{
      userImportTemplate.checkFirstThreeColumnsForBlanks(reportSheet);
    } catch(err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = "Failed: check first three columns for missing values"; //User Only
      throw new Error(`Check not ran for missing values within first three columns. Reason: ${err.name}: ${err.message}.  Please revert sheet to previous version, ensure the correct values are within the first three colums (i.e. "First Name, Last Name, and Email is running a User Import Check), and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
  
    try {
      userImportTemplate.formatUserDateAddedAndBirthdayColumns(reportSheet);
    } catch(err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = "Failed: did not format user date added and birthday columns"; //User Only
      throw new Error(`Check not ran for formatting of user date added and birthday columns. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
    
    try {
      userImportTemplate.convertStatesToTwoLetterCode(reportSheet);
    } catch (err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = userImportTemplate.failedCheckNotRanForConvertingStatesToTwoLetterCodes; //User Programs Needs Agencies
      throw new Error(`Check not ran for converting states to two-letter code: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
  
    try {
      userImportTemplate.validateStates(reportSheet);
    } catch (err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = userImportTemplate.failedInvalidStatesCheckMessage; //Users Needs Agencies
      throw new Error(`Check not ran for invalid states: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
  
    try {
      userImportTemplate.validateGenderOptions(reportSheet);
    } catch (err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = "Failed: check not ran for invalid gender options"; //Users Only
      throw new Error(`Check not ran for invalid gender options: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
    
    // capitalizeFirstLetterOfWords(sheet, reportSheet);
  
    try {
      userImportTemplate.checkForInvalidEmails(reportSheet);
    } catch(err) {
        Logger.log(err);
        userImportTemplate.reportSummaryComments = userImportTemplate.failedInvalidEmailCheckMessage; //All
        throw new Error(`Check not ran for invalid emails. Reason: ${err.name}: ${err.message} at line. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
      } 
    
    try {
      userImportTemplate.checkForInvalidPostalCodes(reportSheet);
    } catch(err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = "Failed: check not ran for invalid postal codes"; //Usrs Neesd Agencies
      throw new Error(`Check not ran for invalid postal codes. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
    
    try {
      userImportTemplate.checkForInvalidNumbers(reportSheet);
    } catch(err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = "Failed: check not ran for invalid home or mobile phone numbers"; //Users Needs Agencies
      throw new Error(`Check not ran for invalid home or mobile phone numbers. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
    
    try {
      userImportTemplate.setCommentsOnReportCell(reportSheet);
    } catch(err) {
      Logger.log(err);
      throw new Error(`Report sheet cell comment not added for summary of checks. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`); //All
    }
  
     SpreadsheetApp.getUi().alert("User Import Check Complete"); //user
    
  }
  
  } catch(err) {
  Logger.log(err);
  throw new Error(`An error occured the the user import template check did not successfully run. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`); //User
  
  
  }
  
  
  
    
  // Creates "User Template Check" navigation button within Spreadsheet UI
  // try {
    function onOpen() {
    let ui = SpreadsheetApp.getUi()
    ui.createMenu('Import Teplate Checker').addItem('Check User Import Template', 'checkUserImportTemplate').addItem('Check Individual Hours Import Template', 'checkIndvImportTemplate').addItem('Check Responses and Hours Template', 'hoursAndResponsesCheck').addItem('Check Agencies/Programs Import Template', 'programsAndAgenciesTemplateCheck').addToUi();
  }
  
  // } catch (err) {
  //   SpreadsheetApp.getUi().alert(`An error occured and the user import template check button did not load. Please refresh the page: ${err.name}: ${err.message}`);
  //   Logger.log(err);
  // }
  
  
  