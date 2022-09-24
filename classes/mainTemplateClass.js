// @ts-nocheck

////////////////////////////////////////////////////////////
//Superclass of all templates
////////////////////////////////////////////////////////////

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
    this._ranCheckForDuplicateEmailsAndSetComments = "Success: ran check for duplicate emails and set error columns for those found";
    this._valuesMissingComment = 'Values Missing';
    this._invalidEmailComment = 'Invalid Email Format';
    this._invalidEmailCellMessage = "Invalid Email/Emails Found";
    this._formattedDateReportMessage = 'Formatted date column(s) found';
    this._foundCommaSeparatedListColumnAndRemovedWhiteSpaceAndEmptyValuesComment = "Success: removed whitespace and empty values from comma-separated list column(s)";
    this._foundCommaSeparatedListColumnAndRemovedWhiteSpaceAndEmptyValuesOnMainSheetHeaderComment = "Removed whitespace and empty values from comma-separated list(s) on main sheet";
    this._foundCommaSeparatedEmailsColumnAndRanCheckForInvalidEmailsComment = "Success: ran check for invalid emails within comma-separated list";
    this._failedCheckNotRanForInvalidCommaSeparatedEmails = "Failed: check not ran for invalid emails within comma-separated list";
    this._failedFormatDateColumns = 'Failed: check not ran for dates to format';
    this._failedRemovedWhiteSpaceFromCellsComment = "Failed: remove white space from cells";
    this._failedRemovedFormattingFromCellsComment = "Failed: remove formatting from cells";
    this._failedInvalidEmailCheckMessage = "Failed: check not ran for invalid emails";
    this._failedCheckNotRanForDuplicateEmails = "Failed: check emails column for duplicates";
    this._failedErrorColumnsNotSetForDuplicateEmails = "Failed: error columns not set for duplicate emails";
    this._failedCheckFirstThreeColumnsForMissingValues = "Failed: check first three columns for missing values";
    this._failedFormatUserDateAddedAndBirthdayColumns = "Failed: did not format user date added and birthday columns";
    this._failedCheckNotRanForInvalidGenderOptions = "Failed: check not ran for invalid gender options";
    this._failedCheckNotRanForInvalidPostalCodes = "Failed: check not ran for invalid postal codes";
    this._failedRemoveWhiteSpaceAndMissingValuesFromCommaSeparatedListsMessage = "Failed: did not remove whitespace and empty values from comma-separated lists";
    this._dateFormat = 'yyyy-mm-dd';
    this._lightGreenHexCode = '#b6d7a8';
    this._lightRedHexCode = '#f4cccc';
    this._errorFoundMessage = 'ERROR FOUND';
    this._dateFormattedComment = 'Date formatted to YYYY-MM-DD';
    
    this._data = this._sheet.getDataRange();
    this._values = this._data.getValues();
    this._columnHeaders = this._values[0].map(header => header.toLowerCase().trim());
    this._emailColumnHeaderOptions = ['email','email address','user email','emails','email addresses', 'program manager email','program manager email address', 'agency manager email','agency manager email address', 'contact email','user email address'];
    this._dateHeaderOptions = ['birthday (yyyy-mm-dd)', 'birthday', 'birth date', 'dob','d.o.b.', 'date of birth','need date', 'date', 'date served', 'date served  (yyyy-mm-dd)', 'date added', 'date added (yyyy-mm-dd)', 'user date added', 'opportunity date', 'response date (yyyy-mm-dd)', 'response date', 'date of response'];
    this._mutipleEmailsHeaderOptions = ['need contact emails', 'additional contacts (emails)', 'additional contacts', 'user group members (emails)', 'user group members', 'group members', 'leaders', 'user group leaders (emails)' , 'user group leaders', 'group leaders', 'user group leader', 'opportunity contact emails'];
    this._genderOptionsHeaderOptions = ['gender (male, female, other, prefer not to say)', 'gender', 'sex'];
    this._commaSeparatedListCleanupHeaderOptions = ['tags (comma separated)', 'interests (comma seperated)', 'interests (comma separated)','tags', 'interests', 'additional contacts (emails)', 'need contact emails', 'user group members (emails)', 'user group members', 'group members', 'leaders', 'user group leaders (emails)' , 'user group leaders', 'group leaders', 'user group leader', 'allowed email domains', 'domains', 'allowed domains', 'allowed email domain', 'email domains'];
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
  get failedCheckNotRanForDuplicateEmails() {
    return this._failedCheckNotRanForDuplicateEmails;
  }
  get failedErrorColumnsNotSetForDuplicateEmails() {
    return this._failedErrorColumnsNotSetForDuplicateEmails;
  }
  get failedCheckFirstThreeColumnsForMissingValues() {
    return this._failedCheckFirstThreeColumnsForMissingValues;
  }
  get failedFormatUserDateAddedAndBirthdayColumns() {
    return this._failedFormatUserDateAddedAndBirthdayColumns;
  }
  get failedCheckNotRanForInvalidGenderOptions() {
    return this._failedCheckNotRanForInvalidGenderOptions;
  }
  get failedCheckNotRanForInvalidPostalCodes() {
    return this._failedCheckNotRanForInvalidPostalCodes;
  }
  get failedRemoveWhiteSpaceAndMissingValuesFromCommaSeparatedListsMessage() {
    return this._failedRemoveWhiteSpaceAndMissingValuesFromCommaSeparatedListsMessage;
  }
  get failedFormatDateColumns() {
    this._failedFormatDateColumns;
  }
  get formattedDateReportMessage() {
    return this._formattedDateReportMessage;
  }
  get failedCheckNotRanForInvalidCommaSeparatedEmails() {
    return this._failedCheckNotRanForInvalidCommaSeparatedEmails;
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
    const newSheet = this.ss.insertSheet();
    newSheet.setName(newSheetName);

    return newSheet;
  }
  copyValuesToSheet(targetSheet, valuesToCopyOver) {
    const rowCount =  valuesToCopyOver.length;
    const columnCount = valuesToCopyOver[0].length;
    const dataRange = targetSheet.getRange(1, 1, rowCount, columnCount);
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
    //This puts the current comments within an array
    const currentComments = sheetCell.getComment().split(",").map((val) => val.trim());
    //Checks to see if the comment to insert is already on the cell in question
    if (!currentComments.includes(comment)) {
      //Checks to see if the current cell has a comment (so new comment is either appended to the comments that already exist, or a new comment is set)
      if (this.checkCellForComment(sheetCell)) {
      const previousComment = sheetCell.getComment();
      const newComment = `${previousComment}, ${comment}`;
      sheetCell.setComment(newComment);
    } else {
      sheetCell.setComment(comment);
    }
   }
  }
  deleteSheetIfExists(sheetBinding) {
    if (sheetBinding !== null) {
    this.ss.deleteSheet(sheetBinding);
    }
  }
 createReportSheet() {
    const reportSheetName = `${this.templateName} Report`;
    reportSheetBinding = this.createSheetCopy(reportSheetName);
    this.copyValuesToSheet(reportSheetBinding, this._values);
    return reportSheetBinding;
  }
  getColumnRange(columnName, sheetBinding) {
    const row = 2;
    const columnPadding = 1;
    const column =  this._columnHeaders.indexOf(columnName);
    if (column !== -1) {
      return sheetBinding.getRange(row, column + columnPadding,sheetBinding.getMaxRows());
    }
    return null;
  }
  getValues(range) {
    return range.getValues();
  }
  insertHeaderComment(headerCell, commentToInsert) {
    //This puts the current header cell comments within an array
    const currentComments = headerCell.getComment().split(",").map((val) => val.trim());
    //Checks to see if the comment to insert is already on the cell in question and the comment to insert is not equal to a falsey value
    if (!currentComments.includes(commentToInsert) && commentToInsert) {
        this.insertCommentToSheetCell(headerCell, commentToInsert);
    }
  }
//Goes through each header cell and removes its comments
  removeHeaderComments() {
    const indexIncrementor = 1;
    const turnDecreaser = 1;
    const headerRow = 1;
    let headerCommentsToRemove =  this.columnHeaders.length;
    let index = 0;
    while (headerCommentsToRemove) {
      const columnPadding = 1;
      const currentColumnPosition = index + columnPadding;
      const headerCell = this.getSheetCell(this._sheet, headerRow, currentColumnPosition);
      headerCell.setComment("");
      headerCommentsToRemove -= turnDecreaser;
      index += indexIncrementor;
    }
  }

  removeCellComments() {
    const currentCell = this._sheet.getActiveCell();
    currentCell.setComment("");
  }

  clearSheetSummaryColumn(sheetBinding) {
    const row = 1;
    const clearRange = () => {
      sheetBinding.getRange(row,this._reportSummaryColumnPosition, sheetBinding.getMaxRows()).clear();
    }
  clearRange();
  }
  //Checks to see if an error comment already exists within an error record's row on the report sheet, and if not, sets an error comment so users can sort for error records quickly
  setErrorColumns(reportSheetBinding, rowBinding) {
    const reportSummaryCell = this.getSheetCell(reportSheetBinding, rowBinding, this._reportSummaryColumnPosition);
    const reportSummaryCellValue = reportSummaryCell.getValue();
    const mainSheetErrorMessageCell = this.getSheetCell(this._sheet, rowBinding, this._reportSummaryColumnPosition);

    if (!reportSummaryCellValue) {
            reportSummaryCell.setValue(this._errorFoundMessage);
            mainSheetErrorMessageCell.setValue(this._errorFoundMessage);
            }
      
    }
  //Checks to see if an error comment already exists within an error record's row on the main sheet, and if not, sets an error comment so users can sort for error rows quickly
  setErrorColumnHeaderInMainSheet() {
    const row = 1;

    const mainSheetErrorMessageCell = this.getSheetCell(this._sheet, row, this._reportSummaryColumnPosition);
    const mainSheetErrorMessageCellValue = mainSheetErrorMessageCell.getValue();

    if (!mainSheetErrorMessageCellValue) {
            mainSheetErrorMessageCell.setValue(this._errorAlertColumnHeader);
            }

  }
  //Joins the report summary comments into a comma-separated string, then adds the comments to the report summary cell on the report sheet
  setCommentsOnReportCell(reportSheetBinding) {
    const reportSummaryCommentsString = this._reportSummaryComments.join(", ");
    const row = 1;
    
    const reportCell = this.getSheetCell(reportSheetBinding, row, this._reportSummaryColumnPosition);

    reportCell.setValue(this._reportOverviewHeaderMessage);
    this.insertCommentToSheetCell(reportCell, reportSummaryCommentsString);
  }

  createReportSheet() {
    const reportSheetName = `${this.sheet.getName()} Report`;
    const  reportSheetBinding = this.createSheetCopy(this.ss, reportSheetName);
    reportSheetBinding.setName(reportSheetName);
    this.copyValuesToSheet(reportSheetBinding, this._values);

    return reportSheetBinding;
  }
  //Grabs the total rows and columns, gets the range of the whole sheet, then trims the whitespace using Google Sheet's trimWhitespace method
  removeWhiteSpaceFromCells() {
    const row = 1;
    const column = 1;
    const rowCount = this._values.length;
    const columnCount = this._values[0].length;
    const searchRange = this._sheet.getRange(row, column, rowCount, columnCount);
    searchRange.trimWhitespace();

    this._reportSummaryComments.push(this._whiteSpaceRemovedSuccessMessage);
    }
  //Grabs the total rows and columns, gets the range of the whole sheet, then clears the format using Google Sheet's clearFormat method
  removeFormattingFromSheetCells() {
    const row = 1;
    const column = 1;
    const rowCount = this._values.length;
    const columnCount = this._values[0].length;
    const searchRange = this._sheet.getRange(row, column, rowCount, columnCount);
    searchRange.clearFormat();

    this._reportSummaryComments.push(this._removedFormattingSuccessMessage);
  }

  setColumnToYYYYMMDDFormat(columnRangeBinding) {
    columnRangeBinding.setNumberFormat(this._dateFormat);
  }
  validateEmail(email){
    return String(email)
      .toLowerCase()
      .match(
        /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/);
  }

////////////////////////////////////////////////////////////
//Takes each cell of a column that contains single emails and if there is a value in it, checks if the email is valid
////////////////////////////////////////////////////////////
//NOTE: 
//Each validation method will turn the header cell of the column with errors red, 
//set a comment about the ivalid values, set an ERROR FOUND comment within the 
//'Error Alert' column, set a comment on the cell with the invalid value within 
//the report sheet, and update the report summary comments binding with a comment 
//saying that the function was ran
////////////////////////////////////////////////////////////

 checkForInvalidEmails(reportSheetBinding) {
    const headerRow = 1;
    //Checks to see if any headers from the valid header options contain email columns so that if multiple email columns are present, this method will check for the validity of all needed columns
    this._emailColumnHeaderOptions.forEach((headerTitle) => {
      const emailColumnRange = this.getColumnRange(headerTitle, this._sheet);
      if (emailColumnRange) {
        // emailColumnRange.setShowHyperlink(false);
        const emailColumnValues = this.getValues(emailColumnRange);
        const emailColumnPosition = emailColumnRange.getColumn();
        //Goes throguh each email within teh current email columns' values
        emailColumnValues.forEach((email, index) => {
          const currentEmail = String(email);
          const row = index + 2;
            //Checks to see if the email is valid and also checks to makes sure the field is not empty so blank cells are not highlighted by this method
            if (currentEmail !== "" && !this.validateEmail(currentEmail)) {
              const currentCell = this.getSheetCell(this._sheet, row, emailColumnPosition);
              const reportSheetCell = this.getSheetCell(reportSheetBinding, row, emailColumnPosition);
              const mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, emailColumnPosition);
              this.setSheetCellBackground(reportSheetCell, this._lightRedHexCode);
              this.setSheetCellBackground(currentCell, this._lightRedHexCode);
              this.setSheetCellBackground(mainSheetHeaderCell, this._lightRedHexCode);
              this.insertCommentToSheetCell(reportSheetCell, this._invalidEmailComment);
              this.insertHeaderComment(mainSheetHeaderCell, this._invalidEmailCellMessage);
              this.setErrorColumns(reportSheetBinding, row);
              }
            }
          );
         }
        });    
    
      this._reportSummaryComments.push(this._ranEmailValidationCheckSuccessMessage);
  }
  //Formats all dated columns by default
  formatAllDatedColumns(reportSheetBinding) {
  const row = 1;
  //Runs on each cell with a date-related header column
  this._dateHeaderOptions.forEach((headerTitle) => {
    const currentDatedColumn = this.getColumnRange(headerTitle, this._sheet);

    if(currentDatedColumn) {
      const currentDatedColumnPosition = currentDatedColumn.getColumn();
      const reportSheetcurrentDatedColumnHeaderCell = this.getSheetCell(reportSheetBinding, row, currentDatedColumnPosition);

      this.setColumnToYYYYMMDDFormat(currentDatedColumn);
      this.setSheetCellBackground(reportSheetcurrentDatedColumnHeaderCell, this._lightRedHexCode);
      this.insertCommentToSheetCell(reportSheetcurrentDatedColumnHeaderCell, this._dateFormattedComment);;
      this._reportSummaryComments.push(this._formattedDateReportMessage);
    }
  });
 }

 ////////////////////////////////////////////////////////////
 //Takes a comma-separated list and removes unecessary spaces and repeated commas
 ////////////////////////////////////////////////////////////

 formatCommaSeparatedLists(reportSheetBinding) {
    const headerRow = 1;
    //Runs on each cell with a valid comma-separated list header column (since multiple columns such as this may be present)
    this._commaSeparatedListCleanupHeaderOptions.forEach((headerTitle) => {
      const currentColumnRange = this.getColumnRange(headerTitle, this._sheet);

      if (currentColumnRange) {
        const oneItem = 1;
        const currentColumnValues = this.getValues(currentColumnRange);
        const currentColumnPosition = currentColumnRange.getColumn();

        currentColumnValues.forEach((valueString, index) => {
          const zeroRangeAndFirstRow = 2;
          const row = index + zeroRangeAndFirstRow;
          const valueStringArray = String(valueString).split(",");

            //Checks to see if the comma-separated item has more than one item to separate with a column (so it doesn't run on singular items/words)
            if (valueStringArray.length > oneItem) {
              //This cleans up the comma-separated items by ensuring there is no trailing whitespace, double commas, or empty values present
              const valueStringNew = valueStringArray.filter((val) => val.trim().length > 0).map((val) => val.trim()).join(",");
              //Checks to see if the old string is equal to the cleaned string, and if not, replaces it with the new string of clean comma-separated values
              if (valueString !== valueStringNew) {
                const currentCell = this.getSheetCell(this._sheet, row, currentColumnPosition);
                const reportSheetHeaderCell = this.getSheetCell(reportSheetBinding, headerRow, currentColumnPosition);
                const removedWhiteSpaceAndBlanksValue = valueStringNew;
                currentCell.setValue(removedWhiteSpaceAndBlanksValue);

                  this.insertHeaderComment(reportSheetHeaderCell, this._foundCommaSeparatedListColumnAndRemovedWhiteSpaceAndEmptyValuesOnMainSheetHeaderComment);
                  this.setSheetCellBackground(reportSheetHeaderCell, this._lightRedHexCode);
                }
              }
              
          });

      } 
    });
    
     this._reportSummaryComments.push(this._foundCommaSeparatedListColumnAndRemovedWhiteSpaceAndEmptyValuesComment);
  }

  ////////////////////////////////////////////////////////////
  //Takes a list of comma-separated emals and checks to see if any of the lists
  //contain invalid emails
  ////////////////////////////////////////////////////////////

  checkForInvalidCommaSeparatedEmails(reportSheetBinding) {
    const headerRow = 1;
    this._mutipleEmailsHeaderOptions.forEach((headerTitle) => {
      const emailColumnRange = this.getColumnRange(headerTitle, this._sheet);
      if (emailColumnRange) {
        const emailColumnValues = this.getValues(emailColumnRange);
        const emailColumnPosition = emailColumnRange.getColumn();

        emailColumnValues.forEach((emailArray, index) => {
          const zeroRangeAndFirstRow = 2;
          const row = index + zeroRangeAndFirstRow;
          //If the email array is equal to a string, this implies that only one email was present within the cell
          if (typeof emailArray === 'string') {
            const currentEmail = emailArray;
            if (currentEmail !== "" && !this.validateEmail(currentEmail)) {
              const currentCell = this.getSheetCell(this._sheet, row, emailColumnPosition);
              const reportSheetCell = this.getSheetCell(reportSheetBinding, row, emailColumnPosition);
              const mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, emailColumnPosition);

              //Line below for testing
              // SpreadsheetApp.getUi().alert(`Invalid Email! ${currentEmail} ${reportSheetCell.getA1Notation()}`);
              //testing line ends here
              this.setSheetCellBackground(reportSheetCell, this._lightRedHexCode);
              this.setSheetCellBackground(currentCell, this._lightRedHexCode);
              this.setSheetCellBackground(mainSheetHeaderCell, this._lightRedHexCode);
              this.insertCommentToSheetCell(reportSheetCell, this._invalidEmailCellMessage);
              this.insertHeaderComment(mainSheetHeaderCell, this._invalidEmailCellMessage);
              this.setErrorColumns(reportSheetBinding, row);
            }
            //if the emailArray is not equal to a string, then there are multiple commas-separated emails to check and the below methods will be used to do so
          } else {
            //email array holds all of the rows' email values
              emailArray.forEach(emailRowArray => {
                //This creates an array of emails within the current row and each email is validated
                emailRowArray.split(",").forEach((email) => {
                  const currentEmail = email.trim();
                  if (currentEmail !== "" && !this.validateEmail(currentEmail)) {
                    const currentCell = this.getSheetCell(this._sheet, row, emailColumnPosition);
                    const reportSheetCell = this.getSheetCell(reportSheetBinding, row, emailColumnPosition);
                    const mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, emailColumnPosition);

                    this.setSheetCellBackground(reportSheetCell, this._lightRedHexCode);
                    this.setSheetCellBackground(currentCell, this._lightRedHexCode);
                    this.setSheetCellBackground(mainSheetHeaderCell, this._lightRedHexCode);
                    this.insertCommentToSheetCell(reportSheetCell, this._invalidEmailCellMessage);
                    this.insertHeaderComment(mainSheetHeaderCell, this._invalidEmailCellMessage);
                    this.setErrorColumns(reportSheetBinding, row);
                    }
                  });
            });
          }
        });
      }
    });
        
    this._reportSummaryComments.push(this._foundCommaSeparatedEmailsColumnAndRanCheckForInvalidEmailsComment);   
  } 

}

