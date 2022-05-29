class ProgramsAndAgenciesTemplate extends UsersNeedsAndAgenciesTemplate {
  constructor() {
    super();
    this._foundFirstFourColumnsAndCheckedThemForMissingValuesComment = "Success: checked for missing values in first four columns";
    this._failedFirstFourColumnsMissingValuesCheckMessage = "Failed: check first four columns for missing values";
  }
  get failedFirstFourColumnsMissingValuesCheckMessage() {
    return this._failedFirstFourColumnsMissingValuesCheckMessage;
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
} 

////////////////////////////////////////////////////////////
//Runs the checks for the Programs/Agencies Import Template
////////////////////////////////////////////////////////////

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
      programsAndAgenciesTemplate.formatCommaSeparatedLists(reportSheet);
    } catch (err) {
        Logger.log(err);
        programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedRemoveWhiteSpaceAndMissingValuesFromCommaSeparatedListsMessage;
        throw new Error(`Did not remove whitespace and empty values from comma-separated list(s). Reason: ${err.name}: ${err.message} at line. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }

    try {
    programsAndAgenciesTemplate.checkForInvalidCommaSeparatedEmails(reportSheet);
    } catch(err) {
        Logger.log(err);
        programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedCheckNotRanForInvalidCommaSeparatedEmails;
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
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedInvalidPhoneNumbersCheck;
    throw new Error(`Check not ran for invalid phone numbers. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  try {
    programsAndAgenciesTemplate.checkForInvalidURL(reportSheet, 'general');
  } catch(err) {
    Logger.log(err);
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedCheckNotRanForGeneralURL;
    throw new Error(`Check not ran for invalid general URL. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  try {
    programsAndAgenciesTemplate.checkForInvalidURL(reportSheet, 'youtube');
  } catch(err) {
    Logger.log(err);
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedCheckNotRanForYouTubeLink;
    throw new Error(`Check not ran for invalid YouTube link. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  try {
    programsAndAgenciesTemplate.checkForInvalidURL(reportSheet, 'twitter');
  } catch(err) {
    Logger.log(err);
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedCheckNotRanForTwitterLink;
    throw new Error(`Check not ran for invalid Twitter link. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  try {
    programsAndAgenciesTemplate.checkForInvalidURL(reportSheet, 'linkedin');
  } catch(err) {
    Logger.log(err);
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedCheckNotRanForTwitterLink;
    throw new Error(`Check not ran for invalid LinkedIn link. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  try {
    programsAndAgenciesTemplate.checkForInvalidURL(reportSheet, 'instagram');
  } catch(err) {
    Logger.log(err);
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedCheckNotRanForTwitterLink;
    throw new Error(`Check not ran for invalid Instagram link. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  try {
    programsAndAgenciesTemplate.checkForInvalidURL(reportSheet, 'facebook');
  } catch(err) {
    Logger.log(err);
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedCheckNotRanForFacebookLink;
    throw new Error(`Check not ran for invalid Facebook link. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
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

