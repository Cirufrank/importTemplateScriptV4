/*

Goals met:

removed white space
Check for invalid additional contact emails YASSSSS
Checks for invalid emails 
checks for missing required values (first four columns-this is position based and imports)


Things to keep in mind:

Need columns for singular email check to be titles "Email"

This may be okay due to position being what matters in the future */

class ProgramsAndAgenciesTemplate extends UsersNeedsAndAgenciesTemplate {
  constructor() {
    super();
    this._foundFirstFourColumnsAndCheckedThemForMissingValuesComment = "Success: checked for missing values in first four columns";
    this._foundAdditionalContactsColumnAndRemovedWhiteSpaceAndEmptyValuesFromAdditionalContactEmailsOnMainSheetHeaderComment = "Removed whitespace and empty values from Additional Contact Emails on main sheet";
    this._foundAdditionalContactsColumnAndRemovedWhiteSpaceAndEmptyValuesComment = "Success: removed whitespace and empty values from Additional Contact Emails";
    this._didNotFindAdditionalContactColumnButRanTheFunctionToRemoveWhiteSpaceAndEmptyValuesComment = "Success: removed whitespace and empty values from Additional Contact Emails, but did note find column";
    this._invalidEmailsFoundHeaderComment = "Invalid Email/Emails found";
    this._foundAdditionalContactsColumnAndRanCheckForInvalidEmailsComment = "Success: ran check for invalid emails";
    this._didNotFindAdditionalContactsColumnButRanCheckForInvalidEmailsComment = "Success: ran check for invalid emails, but did not find column";
    this._failedFirstFourColumnsMissingValuesCheckMessage = "Failed: check first four columns for missing values";
    this._failedRemoveWhiteSpaceAndMissingValuesFromAdditionalContactEmailsMessage = "Failed: did not remove whitespace and empty values from Additional Contact Emails";
    this._failedCheckNotRanForInvalidAdditonalContactEmails = "Failed: check not ran for invalid additional contact emails";
  }
  get failedFirstFourColumnsMissingValuesCheckMessage() {
    return this._failedFirstFourColumnsMissingValuesCheckMessage;
  }
  get failedRemoveWhiteSpaceAndMissingValuesFromAdditionalContactEmailsMessage() {
    return this._failedRemoveWhiteSpaceAndMissingValuesFromAdditionalContactEmailsMessage;
  }
  get failedCheckNotRanForInvalidAdditonalContactEmails() {
    return this._failedCheckNotRanForInvalidAdditonalContactEmails;
  }
  checkFirstFourColumnsForBlanks(reportSheetBinding) {
    let rowStartPosition = 2;
    let columnStartPosition = 1;
    let secondColumnPosition = 2;
    let thirdColumnPositon = 3;
    let fourthColumnPosition = 4;
    let maxRows = this._sheet.getMaxRows();
    let totalColumsToCheck = 4;
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
      let item2CurrentCell = this.getSheetCell(this._sheet, cellRow, secondColumnPosition);
      let item2ReportCurrentCell = this.getSheetCell(reportSheetBinding, cellRow, secondColumnPosition);
      let item2HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, secondColumnPosition);
      let item3 = currentRow[2];
      let item3CurrentCell = this.getSheetCell(this._sheet, cellRow, thirdColumnPositon);
      let item3ReportCell = this.getSheetCell(reportSheetBinding, cellRow, thirdColumnPositon);
      let item3HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, thirdColumnPositon);
      let item4 = currentRow[3];
      let item4CurrentCell = this.getSheetCell(this._sheet, cellRow, fourthColumnPosition);
      let item4ReportCell = this.getSheetCell(reportSheetBinding, cellRow, fourthColumnPosition);
      let item4HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, fourthColumnPosition);

      if (item1.length !== 0 || item2.length !== 0  || item3.length !== 0 || item4.length !== 0) {
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
              case 3:
                this.setSheetCellBackground(item4CurrentCell, this._lightRedHexCode);
                this.setSheetCellBackground(item4ReportCell, this._lightRedHexCode);
                this.setSheetCellBackground(item4HeaderCell, this._lightRedHexCode);
                this.insertCommentToSheetCell(item4ReportCell, this._valuesMissingComment);
                this.setErrorColumns(reportSheetBinding, cellRow);
                this.insertHeaderComment(item4HeaderCell, this._valuesMissingComment);
                break;
            }
          }
        });
      }
    }

  this._reportSummaryComments.push(this._foundFirstFourColumnsAndCheckedThemForMissingValuesComment);

  }
  formatAdditionalContactEmails(reportSheetBinding) {
    let headerRow = 1;
    let emailColumnRange = this.getColumnRange('Additional Contacts (emails)', this._sheet);

    if (emailColumnRange) {
    let emailColumnValues = this.getValues(emailColumnRange);
    let emailColumnPosition = emailColumnRange.getColumn();

    emailColumnValues.forEach((emailString, index) => {
      let row = index + 2;
      let emailArray = String(emailString).split(",");


        if (emailArray.length > 1) {
          let emailArrayNew = emailArray.filter((val) => val.trim().length > 0).map((val) => val.trim()).join(", ");
          let currentCell = this.getSheetCell(this._sheet, row, emailColumnPosition);
          let reportSheetHeaderCell = this.getSheetCell(reportSheetBinding, headerRow, emailColumnPosition);
          let mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, emailColumnPosition);
          let removedWhiteSpaceAndBlanksValue = emailArrayNew;
          currentCell.setValue(removedWhiteSpaceAndBlanksValue);

            this.insertHeaderComment(reportSheetHeaderCell, this._foundAdditionalContactsColumnAndRemovedWhiteSpaceAndEmptyValuesFromAdditionalContactEmailsOnMainSheetHeaderComment);
            this.setSheetCellBackground(reportSheetHeaderCell, this._lightRedHexCode);
            this.setSheetCellBackground(mainSheetHeaderCell, this._lightRedHexCode);
        }

        });

        this._reportSummaryComments.push(this._foundAdditionalContactsColumnAndRemovedWhiteSpaceAndEmptyValuesComment);

        } else {
      this._reportSummaryComments.push(this._didNotFindAdditionalContactColumnButRanTheFunctionToRemoveWhiteSpaceAndEmptyValuesComment);
    }
    
  }
  checkForInvalidAdditionalContactEmails(reportSheetBinding) {
    let headerRow = 1;
    let emailColumnRange = this.getColumnRange('Additional Contacts (emails)', this._sheet);

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
            this.insertCommentToSheetCell(reportSheetCell, this._invalidEmailComment);
            this.insertHeaderComment(mainSheetHeaderCell, this._invalidEmailsFoundHeaderComment);
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
                this.insertCommentToSheetCell(reportSheetCell, this._invalidEmailComment);
                this.insertHeaderComment(mainSheetHeaderCell, this._invalidEmailsFoundHeaderComment);
                this.setErrorColumns(reportSheetBinding, row);
                }
              });
          });

        }

        
        }
      );

      this._reportSummaryComments.push(this._foundAdditionalContactsColumnAndRanCheckForInvalidEmailsComment);
      
    } else {
      this._reportSummaryComments.push(this._didNotFindAdditionalContactsColumnButRanCheckForInvalidEmailsComment);
    }
    
  }
} 

try {
  function programsAndAgenciesTemplateCheck() {
    const programsAndAgenciesTemplate = new ProgramsAndAgenciesTemplate();
    let reportSheet = programsAndAgenciesTemplate.createReportSheet();

    programsAndAgenciesTemplate.sheet.setName(programsAndAgenciesTemplate.templateName);
    

   programsAndAgenciesTemplate.ss.setActiveSheet(programsAndAgenciesTemplate.getSheet(programsAndAgenciesTemplate.templateName));

    


    programsAndAgenciesTemplate.setFrozenRows(programsAndAgenciesTemplate.sheet, 1);
    programsAndAgenciesTemplate.setFrozenRows(reportSheet, 1);

   try {
          programsAndAgenciesTemplate.removeWhiteSpaceFromCells();
        } catch (err) {
          Logger.log(err);
          programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedRemovedWhiteSpaceFromCellsComment;
          throw new Error(`White space not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
        }

        try {
          programsAndAgenciesTemplate.removeFormattingFromSheetCells();
        } catch (err) {
          Logger.log(err);
          programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedRemovedFormattingFromCellsComment;
          throw new Error(`Formatting not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
        }
    
    programsAndAgenciesTemplate.clearSheetSummaryColumn(programsAndAgenciesTemplate.sheet);
    programsAndAgenciesTemplate.clearSheetSummaryColumn(reportSheet);
    programsAndAgenciesTemplate.setErrorColumnHeaderInMainSheet(programsAndAgenciesTemplate.sheet);

    try{
      programsAndAgenciesTemplate.checkFirstFourColumnsForBlanks(reportSheet);
        } catch(err) {
            Logger.log(err);
            programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedFirstFourColumnsMissingValuesCheckMessage;
            throw new Error(`Check not ran for missing values within first four columns. Reason: ${err.name}: ${err.message}.  Please revert sheet to previous version, ensure the correct values are within the first four colums (i.e. "First Name, Last Name, and Email" is running a User Import Check [this would be for first three columns, but this checker is for four]), and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
        }

    try {
    programsAndAgenciesTemplate.convertStatesToTwoLetterCode(reportSheet);
  } catch (err) {
    Logger.log(err);
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedCheckNotRanForConvertingStatesToTwoLetterCodes ;
    throw new Error(`Check not ran for converting states to two-letter code: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }

  try {
    programsAndAgenciesTemplate.validateStates(reportSheet);
  } catch (err) {
    Logger.log(err);
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedInvalidStatesCheckMessage;
    throw new Error(`Check not ran for invalid states: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }

  try {
    programsAndAgenciesTemplate.checkForInvalidEmails(reportSheet);
  } catch(err) {
      Logger.log(err);
      programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedInvalidEmailCheckMessage;
      throw new Error(`Check not ran for invalid emails. Reason: ${err.name}: ${err.message} at line. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
    } 

    try {
      programsAndAgenciesTemplate.formatAdditionalContactEmails(reportSheet);
    } catch (err) {
        Logger.log(err);
        programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedRemoveWhiteSpaceAndMissingValuesFromAdditionalContactEmailsMessage;
        throw new Error(`Did not remove whitespace and empty values from Additional Contact Emails. Reason: ${err.name}: ${err.message} at line. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }

    try {
    programsAndAgenciesTemplate.checkForInvalidAdditionalContactEmails(reportSheet);
    } catch(err) {
        Logger.log(err);
        programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedCheckNotRanForInvalidAdditonalContactEmails;
        throw new Error(`Check not ran for invalid additional contact emails. Reason: ${err.name}: ${err.message} at line. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
      } 
  
  try {
    programsAndAgenciesTemplate.checkForInvalidPostalCodes(reportSheet);
  } catch(err) {
    Logger.log(err);
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedInvalidPostalCodeCheck;
    throw new Error(`Check not ran for invalid postal codes. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  try {
    programsAndAgenciesTemplate.checkForInvalidNumbers(reportSheet);
  } catch(err) {
    Logger.log(err);
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedInvalidHomeAndMobilePhoneNumbersCheck;
    throw new Error(`Check not ran for invalid home or mobile phone numbers. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  
  try {
    programsAndAgenciesTemplate.setCommentsOnReportCell(reportSheet);
  } catch(err) {
    Logger.log(err);
    throw new Error(`Report sheet cell comment not added for summary of checks. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }

   SpreadsheetApp.getUi().alert("Agencies/Programs Import Check Complete");
  }

} catch (err) {
    Logger.log(err);
    throw new Error(`An error occured the the Agencies/Programs import template check did not successfully run. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }

