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
    let currentComments = sheetCell.getComment().split(",").map((val) => val.trim());
    if (!currentComments.includes(comment)) {
      if (this.checkCellForComment(sheetCell)) {
      let previousComment = sheetCell.getComment();
      let newComment = `${previousComment}, ${comment}`;
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
    let currentComments = headerCell.getComment().split(",").map((val) => val.trim());
    if (!currentComments.includes(commentToInsert) && commentToInsert) {
      if (!this.checkCellForComment(headerCell)) {
        this.insertCommentToSheetCell(headerCell, commentToInsert);
      } else {
        this.insertCommentToSheetCell(headerCell, commentToInsert);
      }
    }
  }
  removeHeaderComments() {
    const headerRow = 1;
    this.columnHeaders.forEach((headerColumn, index) => {
      let currentColumnPosition = index + 1;
      let headerCell = this.getSheetCell(this._sheet, headerRow, currentColumnPosition);
      headerCell.setComment("");
    });
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

  createReportSheet() {
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
    let searchRange = this._sheet.getRange(1, 1, rowCount, columnCount);
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

////////////////////////////////////////////////////////////
//Takes each cell of a column that contains single emails and if there is a value in it, checks if the email is valid
////////////////////////////////////////////////////////////
//NOTE: 
//Each validation method will turn the header cell of the column with errors red, 
//sets a comment about the ivalid values, set an ERROR FOUND comment withint the 
//'Error Alert' column, set a comment on the cell with the invalid value within 
//the report sheet, and update the report summary comments binding with a comment 
//saying that the function was ran
////////////////////////////////////////////////////////////

 checkForInvalidEmails(reportSheetBinding) {
    let headerRow = 1;

    this._emailColumnHeaderOptions.forEach((headerTitle) => {
      let emailColumnRange = this.getColumnRange(headerTitle, this._sheet);
      if (emailColumnRange) {
        // emailColumnRange.setShowHyperlink(false);
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
         }
        });    
    
      this._reportSummaryComments.push(this._ranEmailValidationCheckSuccessMessage);
  }
  formatAllDatedColumns(reportSheetBinding) {
  let row = 1;

  this._dateHeaderOptions.forEach((headerTitle) => {
    let currentDatedColumn = this.getColumnRange(headerTitle, this._sheet);

    if(currentDatedColumn) {
      let currentDatedColumnPosition = currentDatedColumn.getColumn();
      let reportSheetcurrentDatedColumnHeaderCell = this.getSheetCell(reportSheetBinding, row, currentDatedColumnPosition);

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
    let headerRow = 1;
    this._commaSeparatedListCleanupHeaderOptions.forEach((headerTitle) => {
      let currentColumnRange = this.getColumnRange(headerTitle, this._sheet);

      if (currentColumnRange) {
      let currentColumnValues = this.getValues(currentColumnRange);
      let currentColumnPosition = currentColumnRange.getColumn();

      currentColumnValues.forEach((valueString, index) => {
        let row = index + 2;
        let valueStringArray = String(valueString).split(",");


          if (valueStringArray.length > 1) {
            let valueStringNew = valueStringArray.filter((val) => val.trim().length > 0).map((val) => val.trim()).join(",");
            if (valueString !== valueStringNew) {
              let currentCell = this.getSheetCell(this._sheet, row, currentColumnPosition);
              let reportSheetHeaderCell = this.getSheetCell(reportSheetBinding, headerRow, currentColumnPosition);
              let removedWhiteSpaceAndBlanksValue = valueStringNew;
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
    let headerRow = 1;
    this._mutipleEmailsHeaderOptions.forEach((headerTitle) => {
      let emailColumnRange = this.getColumnRange(headerTitle, this._sheet);
      if (emailColumnRange) {
        let emailColumnValues = this.getValues(emailColumnRange);
        let emailColumnPosition = emailColumnRange.getColumn();

        emailColumnValues.forEach((emailArray, index) => {
          let row = index + 2;

          if (typeof emailArray === 'string') {
            let currentEmail = emailArray;
            if (currentEmail !== "" && !this.validateEmail(currentEmail)) {
              let currentCell = this.getSheetCell(this._sheet, row, emailColumnPosition);
              let reportSheetCell = this.getSheetCell(reportSheetBinding, row, emailColumnPosition);
              let mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, emailColumnPosition);

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
          } else {
            //email array holds all of the rows' email values
              emailArray.forEach(emailRowArray => {
                //This creates and array of emails within the current row and each email is validated
                emailRowArray.split(",").forEach((email) => {
                  let currentEmail = email.trim();
                  if (currentEmail !== "" && !this.validateEmail(currentEmail)) {
                  let currentCell = this.getSheetCell(this._sheet, row, emailColumnPosition);
                  let reportSheetCell = this.getSheetCell(reportSheetBinding, row, emailColumnPosition);
                  let mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, emailColumnPosition);

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

////////////////////////////////////////////////////////////
//Superclass of Agency/Programs, Needs/Opportunities, and Users Templates
////////////////////////////////////////////////////////////

class UsersNeedsAndAgenciesTemplate extends Template {
  constructor() {
    super();
    this._stateHeaderOptions = ['state','state (ex: nh)', 'state (e.g. tn)','need state', 'opportunity state'];
    this._plainTextNumberFormat = `@STRING@`;
    this._phoneHeaderOptions = ['phone number', 'phone','mobile', 'cell phone', 'cell phone numbers','number','user phone','user phone number','home phone number','mobile phone number', 'mobile phone', 'cell phone number', 'mobile phone numbers', 'user phone numbers', 'user phone'];
    this._zipHeaderOptions = ['zipcode', 'zip code', 'postal','postal code', 'zip', 'zip codes', 'postal codes', 'user zip', 'user zip codes', 'user postal codes', 'user postal code', 'need zip', 'opportunity zip', 'opportunity zip code', 'need zip code', 'need postal code', 'opportunity postal code'];
    this._generalURLColumnOptions = ['website url', 'main site', 'webpage'];
    this._facebookLinkColumnOptions = ['facebook link', 'facebook', 'facebook page', 'fb', 'fb page', 'fb link', 'face book', 'face book page'];
    this._twitterColumnOptions = ['twitter link', 'twitter', 'twitter page', 'twitter url'];
    this._linkedInColumnOptions = ['linkedin link', 'linked in link', 'linked in', 'linkedin'];
    this._instagramColumnOptions = ['instagram link', 'instagram page', 'instagram'];
    this._youTubeColumnoptions = ['agency video (youtube or vimeo url)', 'youtube', 'youtube link', 'youtube video'];
    this._invalidLinkComment = 'Invalid link';
    this._invalidLinkHeaderComment = 'Invalid link/links found';
    this._stateTwoLetterCodeComment = 'State converted to two letter code'; 
    this._invalidPostalCodeComment = 'Invalid postal code'; 
    this._invalidPhoneNumberComment = 'Invalid phone number'; 
    this._invalidStateComment = 'Invalid State'; 
    this._invalidStateFoundHeaderComment = "Invalid State/States found";
    this._invalidPhoneNumbersFoundHeaderComment =  "Invalid Phone Number/Numbers Found";
    this._invalidPostalCodeFoundHeaderComment = "Invalid Postal Code/Codes Found";
    this._stateColumnFoundAndConversionFunctionRanComment = "Success: ran check to convert states to two-digit format";
    this._stateColumnNotFoundAndConversionRanComment = "Success: ran check to convert states to two-digit format, but did not find column";
    this._stateColumnFoundAndValidationCheckRanComment = "Success: ran check for invalid states";
    this._stateColumnNotFoundAndValidationCheckRanComment = "Success: ran check for invalid states, but did not find column";
    this._phoneNumberColumnFoundAndValidationCheckRanComment = "Success: ran check for invalid phone numbers";
    this._postalCodeColumnFoundAndValidationCheckRan = "Success: ran check for invalid postal codes";
    this._postalCodeColumnNotFoundAndValidationCheckRan = "Success: ran check for invalid postal codes, but did not find column";
    this._generalLinkFoundAndValidationCheckRan = "Success: ran check for invalid general Program/Agency URL";
    this._twitterLinkFoundAndValidationCheckRan = "Success: ran check for invalid Twitter Program/Agency URL";
    this._facebookLinkFoundAndValidationCheckRan = "Success: ran check for invalid Facebook Program/Agency URL";
    this._instagramLinkFoundAndValidationCheckRan = "Success: ran check for invalid Instagram Program/Agency URL";
    this._youtubeLinkFoundAndValidationCheckRan = "Success: ran check for invalid YouTube Program/Agency URL";
    this._linkedInLinkFoundAndValidationCheckRan = "Success: ran check for invalid LinkedIn Program/Agency URL";
    this._failedCheckNotRanForGeneralURL = "Failed: check not ran for general URL";
    this._failedCheckNotRanForTwitterLink = "Failed: check not ran for Twitter Link";
    this._failedCheckNotRanForFacebookLink = "Failed: check not ran for Facebook Link";
    this._failedCheckNotRanForInstagramLink = "Failed: check not ran for Instagram Link";
    this._failedCheckNotRanForYouTubeLink = "Failed: check not ran for YouTube Link";
    this._failedCheckNotRanForLinkedInLink = "Failed: check not ran for LinkedIn Link";
    this._failedCheckNotRanForConvertingStatesToTwoLetterCodes = "Failed: check not ran for converting states to two-letter code";
    this._failedInvalidStatesCheckMessage = "Failed: check not ran for invalid states";
    this._failedInvalidPostalCodeCheck = "Failed: check not ran for invalid postal codes";
    this._failedInvalidPhoneNumbersCheck = "Failed: check not ran for invalid phone numbers";
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
  get failedInvalidPhoneNumbersCheckk() {
    return this._failedInvalidPhoneNumbersCheck;
  }
  get failedCheckNotRanForGeneralURL() {
    return this._failedCheckNotRanForGeneralURL;
  }
  get failedCheckNotRanForInstagramLink() {
    return this._failedCheckNotRanForInstagramLink;
  }
  get failedCheckNotRanForLinkedInLink() {
    return this._failedCheckNotRanForLinkedInLink;
  }
  get failedCheckNotRanForTwitterLink() {
    return this._failedCheckNotRanForTwitterLink;
  }
  get failedCheckNotRanForFacebookLink() {
    return this._failedCheckNotRanForFacebookLink;
  }
  get failedCheckNotRanForYouTubeLink() {
    return this._failedCheckNotRanForYouTubeLink;
  }
  convertStatesToTwoLetterCode(reportSheetBinding) {
    this._stateHeaderOptions.forEach((headerTitle) => {
      let stateColumnRange = this.getColumnRange(headerTitle, this._sheet);
      if (stateColumnRange) {
        let stateColumnRangeValues = this.getValues(stateColumnRange).map(val => {
          let currentState = String(val);
          if (currentState.length > 2) {
            return (currentState[0].toUpperCase() + currentState.slice(1).toLowerCase());
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

    }
   }); 
    
  }
  validateStates(reportSheetBinding) {
    this._stateHeaderOptions.forEach((headerTitle) => {
      let stateColumnRange = this.getColumnRange(headerTitle, this._sheet);
      if (stateColumnRange) {
        let stateColumnRangeValues = this.getValues(stateColumnRange);

        let stateColumnRangePosition = stateColumnRange.getColumn();
        
        stateColumnRangeValues.forEach((val, index) => {
          let currentState = String(val).toUpperCase();
          
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
      }
    });
  }
  validatePhoneNumbers(number) { 
    return String(number)
      .toLowerCase()
      .match(
    /^\s*(?:\+?(\d{1,3}))?[-. (]*(\d{3})[-. )]*(\d{3})[-. ]*(\d{4})(?: *x(\d+))?\s*$/
      );
  }

  validatePostalCode(postalCode) { 
    return /(^\d{5}$)|(^\d{5}-\d{4}$)/.test(postalCode);
  }
  validateGeneralURL(url) {
    return String(url)
    .toLowerCase()
    .match(
      /^(?:(?:https?|http):\/\/)?(?:(?!(?:10|127)(?:\.\d{1,3}){3})(?!(?:169\.254|192\.168)(?:\.\d{1,3}){2})(?!172\.(?:1[6-9]|2\d|3[0-1])(?:\.\d{1,3}){2})(?:[1-9]\d?|1\d\d|2[01]\d|22[0-3])(?:\.(?:1?\d{1,2}|2[0-4]\d|25[0-5])){2}(?:\.(?:[1-9]\d?|1\d\d|2[0-4]\d|25[0-4]))|(?:(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)(?:\.(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)*(?:\.(?:[a-z\u00a1-\uffff]{2,})))(?::\d{2,5})?(?:\/\S*)?$/
    );
  }
  validateTwitterLink(twitterLink) {
    return String(twitterLink)
      .toLowerCase()
      .match(
        /https:\/\/[www.]*twitter.com\/.+/
      );
  }
  validateInstagramLink(instagramLink) {
    return String(instagramLink)
      .toLowerCase()
      .match(
        /https:\/\/[www.]*instagram.com\/.+/
      );
  }
  validateYouTubeLink(youTubeLink) {
    return String(youTubeLink)
      .toLowerCase()
      .match(
        /https:\/\/[www.]*youtube.com\/.+/
      );
  }
  validateLinkedInLink(linkedInLink) {
    return String(linkedInLink)
      .toLowerCase()
      .match(
        /https:\/\/[www.]*linkedin.com\/.+/
      );
  }
  validateFaceBookLink(faceBookLink) {
    return String(faceBookLink)
      .toLowerCase()
      .match(
        /https:\/\/[www.]*facebook.com\/.+/
      );
  }
  checkForInvalidURL(reportSheetBinding, urlType) {
    let headerOptions;
    let validationFunction;
    let reportSummaryCommentToUse;

    switch (urlType) {
      case 'general':
        headerOptions = this._generalURLColumnOptions;
        validationFunction = this.validateGeneralURL;
        reportSummaryCommentToUse = this._generalLinkFoundAndValidationCheckRan;
        break;
      case 'youtube':
        headerOptions = this._youTubeColumnoptions;
        validationFunction = this.validateYouTubeLink;
        reportSummaryCommentToUse = this._youtubeLinkFoundAndValidationCheckRan;
        break;
      case 'twitter':
        headerOptions = this._twitterColumnOptions;
        validationFunction = this.validateTwitterLink;
        reportSummaryCommentToUse = this._twitterLinkFoundAndValidationCheckRan;
        break;
      case 'linkedin':
        headerOptions = this._linkedInColumnOptions;
        validationFunction = this.validateLinkedInLink;
        reportSummaryCommentToUse = this._linkedInLinkFoundAndValidationCheckRan;
        break;
      case 'instagram':
        headerOptions = this._instagramColumnOptions;
        validationFunction = this.validateInstagramLink;
        reportSummaryCommentToUse = this._instagramLinkFoundAndValidationCheckRan;
        break;
      case 'facebook':
        headerOptions = this._facebookLinkColumnOptions;
        validationFunction = this.validateFaceBookLink;
        reportSummaryCommentToUse = this._facebookLinkFoundAndValidationCheckRan;
    }

    let headerRow = 1;
    headerOptions.forEach((headerTitle) => {
      let linkColumn = this.getColumnRange(headerTitle, this._sheet);
      if (linkColumn) {
        let linkColumnRangeValues = this.getValues(linkColumn);
        let linkColumnRangePosition = linkColumn.getColumn();
        let mainSheetLinkHeaderCell = this.getSheetCell(this._sheet, headerRow, linkColumnRangePosition);

        linkColumnRangeValues.forEach((link, index) => {
          //Pickup changing bindings here
          let currentLink = String(link);
          let row = index + 2;

          if (currentLink !== "" && !validationFunction(currentLink)) {
            let currentCell = this.getSheetCell(this._sheet, row, linkColumnRangePosition);
            let reportSheetCell = this.getSheetCell(reportSheetBinding, row, linkColumnRangePosition);

            this.setSheetCellBackground(reportSheetCell, this._lightRedHexCode);
            this.setSheetCellBackground(currentCell, this._lightRedHexCode);
            this.setSheetCellBackground(mainSheetLinkHeaderCell, this._lightRedHexCode);
            this.insertCommentToSheetCell(reportSheetCell, this._invalidLinkComment);
            this.insertHeaderComment(mainSheetLinkHeaderCell,this._invalidLinkHeaderComment);
            this.setErrorColumns(reportSheetBinding, row);
            }
          });
        }
      });

    this._reportSummaryComments.push(reportSummaryCommentToUse);
  }
  setColumnToPlainTextNumerFormat(columnRangeBinding) { //ALL BUT PROGRAMS AND AGENCIES
    columnRangeBinding.setNumberFormat(this._plainTextNumberFormat);
  }

  checkForInvalidNumbers(reportSheetBinding) {
    let headerRow = 1;
    this._phoneHeaderOptions.forEach((headerTitle) => {
      let phoneNumberRange = this.getColumnRange(headerTitle, this._sheet);
      if (phoneNumberRange) {
      let phoneNumberRangeValues = this.getValues(phoneNumberRange);
      let phoneNumberRangePosition = phoneNumberRange.getColumn();
      let mainSheetPhoneHeaderCell = this.getSheetCell(this._sheet, headerRow, phoneNumberRangePosition);

      phoneNumberRangeValues.forEach((number, index) => {

        let currentNumber= String(number);
        let row = index + 2;

        if (currentNumber !== "" && !this.validatePhoneNumbers(currentNumber)) {
          let currentCell = this.getSheetCell(this._sheet, row, phoneNumberRangePosition);
          let reportSheetCell = this.getSheetCell(reportSheetBinding, row, phoneNumberRangePosition);

          this.setSheetCellBackground(reportSheetCell, this._lightRedHexCode);
          this.setSheetCellBackground(currentCell, this._lightRedHexCode);
          this.setSheetCellBackground(mainSheetPhoneHeaderCell, this._lightRedHexCode);
          this.insertCommentToSheetCell(reportSheetCell, this._invalidPhoneNumberComment);
          this.insertHeaderComment(mainSheetPhoneHeaderCell,this._invalidPhoneNumbersFoundHeaderComment);
          this.setErrorColumns(reportSheetBinding, row);
          }
        });
      }
    });

    this._reportSummaryComments.push(this._phoneNumberColumnFoundAndValidationCheckRanComment);
  }
  checkForInvalidPostalCodes(reportSheetBinding) {
    let headerRow = 1;
    this._zipHeaderOptions.forEach((headerTitle) => {
      let postalCodeColumnRange = this.getColumnRange(headerTitle, this._sheet);
      if (postalCodeColumnRange) {
        this.setColumnToPlainTextNumerFormat(postalCodeColumnRange)
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
          });
         }
       });
    this._reportSummaryComments.push(this._postalCodeColumnFoundAndValidationCheckRan);
  } 

}

class UserTemplate extends UsersNeedsAndAgenciesTemplate {
  constructor() {
    super();
    this._duplicateEmailComment = 'Duplicate Email';
    this._duplicateEmailsFoundMessage = "Duplicate Email/Emails Found";
    this._genderExpandedComment = 'Gender option expanded';
    this._invalidGenderOptionComment = 'Invalid Gender Option';
    this._genderOptionsAbbreviationToFullObject = {
      "F":"Female",
      "M":"Male",
      "N/A":"",
    }
    this._genderOptionsAbbreviations = Object.keys(this._genderOptionsAbbreviationToFullObject);
    this._fullGenderOptions = ['female', 'male','prefer not to say', 'other', 'f','m'];
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
    this._emailColumnHeaderOptions.forEach((headerTitle) => {
      let emailColumnRange = this.getColumnRange(headerTitle, this._sheet);
      if (emailColumnRange) {
        let reportSheetEmailColumnRange = this.getColumnRange(headerTitle, reportSheetBinding);
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
        reportSheetBinding.sort(emailColumnPosition, true);
        this._sheet.sort(emailColumnPosition, true);
    
      }
    });
    
    this._reportSummaryComments.push(this._foundEmailColumnAndCheckedForDuplicatesComment);

  }
  checkForLightRedHexCode(cell) {
    return cell.getBackground() === this._lightRedHexCode;
  }
  checkForDuplicateEmailsAndSetErrors(reportSheetBinding) {
    let headerRow = 1;
    this._emailColumnHeaderOptions.forEach((headerTitle) => {
      let emailColumnRange = this.getColumnRange(headerTitle, this._sheet);
      if (emailColumnRange) {
        let emailColumnValues = this.getValues(emailColumnRange);
        let emailColumnPosition = emailColumnRange.getColumn();

        emailColumnValues.forEach((email, index) => {
          let currentEmail = String(email).trim();
          let row = index + 2;
          let nextRow = row + 1;
          let currentCell = this.getSheetCell(this._sheet, row, emailColumnPosition);

          if (currentEmail.length > 0 && this.checkForLightRedHexCode(currentCell)) {
            // let nextCell = this.getSheetCell(this._sheet, nextRow, emailColumnPosition);
            let reportSheetCell = this.getSheetCell(reportSheetBinding, row, emailColumnPosition);
            let nextReportSheetCell = this.getSheetCell(reportSheetBinding, nextRow, emailColumnPosition)
            let mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, emailColumnPosition);
            this.setSheetCellBackground(reportSheetCell, this._lightRedHexCode);
            this.setSheetCellBackground(nextReportSheetCell, this._lightRedHexCode);
            this.setSheetCellBackground(mainSheetHeaderCell, this._lightRedHexCode);
            this.insertCommentToSheetCell(reportSheetCell, this._duplicateEmailComment);
            this.insertCommentToSheetCell(nextReportSheetCell, this._duplicateEmailComment);
            this.insertHeaderComment(mainSheetHeaderCell, this._duplicateEmailsFoundMessage);
            this.setErrorColumns(reportSheetBinding, row);
            this.setErrorColumns(reportSheetBinding, nextRow);
            }
          });
        }
      });

      this._reportSummaryComments.push(this._ranCheckForDuplicateEmailsAndSetComments);
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
  validateGenderOptions(reportSheetBinding) {
    this._genderOptionsHeaderOptions.forEach((headerTitle) => {
      let genderOptionColumnRange = this.getColumnRange(headerTitle, this._sheet);
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
       }
      });

    this._reportSummaryComments.push(this._foundGenderOptionsColumnAndRanCheckSuccessfullyComment);
  }
} 

////////////////////////////////////////////////////////////
//Runs the checks for the User Import Template
////////////////////////////////////////////////////////////

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
    userImportTemplate.reportSummaryComments = userImportTemplate.failedCheckNotRanForDuplicateEmails; //User Only
    throw new Error(`Emails not checked for duplicates. Reason: ${err.name}: ${err.message}. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  try {
    userImportTemplate.checkForDuplicateEmailsAndSetErrors(reportSheet);
  } catch(err) {
    Logger.log(err);
    userImportTemplate.reportSummaryComments = userImportTemplate.failedErrorColumnsNotSetForDuplicateEmails; //User Only
    throw new Error(`Error columns not set for duplicate emails. Reason: ${err.name}: ${err.message}. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
  }

  try{
    userImportTemplate.checkFirstThreeColumnsForBlanks(reportSheet);
  } catch(err) {
    Logger.log(err);
    userImportTemplate.reportSummaryComments = userImportTemplate.failedCheckFirstThreeColumnsForMissingValues;
    throw new Error(`Check not ran for missing values within first three columns. Reason: ${err.name}: ${err.message}.  Please revert sheet to previous version, ensure the correct values are within the first three colums (i.e. "First Name, Last Name, and Email is running a User Import Check), and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
  }

  try {
    userImportTemplate.formatAllDatedColumns(reportSheet);
  } catch(err) {
    Logger.log(err);
    userImportTemplate.reportSummaryComments = userImportTemplate.failedFormatDateColumns; //User Only
    throw new Error(`Check not ran for formatting of dated columns. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  try {
      userImportTemplate.formatCommaSeparatedLists(reportSheet);
    } catch (err) {
        Logger.log(err);
        userImportTemplate.reportSummaryComments = userImportTemplate.failedRemoveWhiteSpaceAndMissingValuesFromCommaSeparatedListsMessage;
        throw new Error(`Did not remove whitespace and empty values from comma-separated list(s). Reason: ${err.name}: ${err.message} at line. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
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
    userImportTemplate.reportSummaryComments = userImportTemplate.failedCheckNotRanForInvalidGenderOptions; //Users Only
    throw new Error(`Check not ran for invalid gender options: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }

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
    userImportTemplate.reportSummaryComments = userImportTemplate.failedCheckNotRanForInvalidPostalCodes; //Usrs Neesd Agencies
    throw new Error(`Check not ran for invalid postal codes. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  
  try {
    userImportTemplate.checkForInvalidNumbers(reportSheet);
  } catch(err) {
    Logger.log(err);
    userImportTemplate.reportSummaryComments = userImportTemplate.failedInvalidPhoneNumbersCheck; //Users Needs Agencies
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
throw new Error(`An error occured the the user import template check did not successfully run. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
}

  
// Creates the "Template Check" menu and buttons within Spreadsheet UI
  function onOpen() {
  let ui = SpreadsheetApp.getUi()
  ui.createMenu('Import Teplate Checker').addItem('Remove Header Cell Comments', 'removeHeaderCellComments').addItem('Script Termination Page', 'openURL').addItem('Check User Import Template', 'checkUserImportTemplate').addItem('Check Individual Hours Import Template', 'checkIndvImportTemplate').addItem('Check Responses and Hours Template', 'hoursAndResponsesCheck').addItem('Check Need/Opportunity Responses Import Template', 'needAndOpportunityResponsesCheck').addItem('Check Agencies/Programs Template', 'programsAndAgenciesTemplateCheck').addItem('Check Needs/Opportunities Template', 'needsAndOpportunitiesTemplateCheck').addItem('Check User Groups Template', 'userGroupsTemplateCheck').addToUi();
}


