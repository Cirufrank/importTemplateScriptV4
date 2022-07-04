class NeedsAndOpportunitiesTemplate extends UsersNeedsAndAgenciesTemplate {
  constructor() {
    super();
    this._foundFirstTwoColumnsAndCheckedThemForMissingValuesComment = "Success: checked for missing values in first two columns";
    this._failedFirstTwoColumnsMissingValuesCheckMessage = "Failed: check first two columns for missing values";
  }
  checkFirstTwoColumnsForBlanks(reportSheetBinding) {
    let rowStartPosition = 2;
    let firstColumnPosition = 1;
    let secondColumnPosition = 2;
    let maxRows = this._sheet.getMaxRows();
    let totalColumsToCheck = 2;
    let range = this._sheet.getRange(rowStartPosition, firstColumnPosition, maxRows, totalColumsToCheck);
    let values = range.getValues();

    for (let row = 0; row < values.length; row += 1) {
      let headerRowPosition = 1;
      let cellRow = row + 2;
      let currentRow = values[row];
      let item1 = currentRow[0];
      let item1CurrentCell = this.getSheetCell(this._sheet, cellRow, firstColumnPosition);
      let item1ReportCurrentCell = this.getSheetCell(reportSheetBinding, cellRow, firstColumnPosition);
      let item1HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, firstColumnPosition);
      let item2 = currentRow[1];
      let item2CurrentCell = this.getSheetCell(this._sheet, cellRow, secondColumnPosition);
      let item2ReportCurrentCell = this.getSheetCell(reportSheetBinding, cellRow, secondColumnPosition);
      let item2HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, secondColumnPosition);

      if (item1.length !== 0 || item2.length !== 0) {
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
            }
          }
        });
      }
    }

  this._reportSummaryComments.push(this._foundFirstTwoColumnsAndCheckedThemForMissingValuesComment);

  }
}

////////////////////////////////////////////////////////////
//Runs the checks for the Needs/Opportunities Import Template
////////////////////////////////////////////////////////////

try {
  function needsAndOpportunitiesTemplateCheck() {
    const needsAndOpportunitiesTemplate = new NeedsAndOpportunitiesTemplate();
    let reportSheet = needsAndOpportunitiesTemplate.createReportSheet();

    needsAndOpportunitiesTemplate.sheet.setName(needsAndOpportunitiesTemplate.templateName);
    

   needsAndOpportunitiesTemplate.ss.setActiveSheet(needsAndOpportunitiesTemplate.getSheet(needsAndOpportunitiesTemplate.templateName));

    


    needsAndOpportunitiesTemplate.setFrozenRows(needsAndOpportunitiesTemplate.sheet, 1);
    needsAndOpportunitiesTemplate.setFrozenRows(reportSheet, 1);

   try {
          needsAndOpportunitiesTemplate.removeWhiteSpaceFromCells();
        } catch (err) {
          Logger.log(err);
          needsAndOpportunitiesTemplate.reportSummaryComments = needsAndOpportunitiesTemplate.failedRemovedWhiteSpaceFromCellsComment;
          throw new Error(`White space not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
        }

        try {
          needsAndOpportunitiesTemplate.removeFormattingFromSheetCells();
        } catch (err) {
          Logger.log(err);
          needsAndOpportunitiesTemplate.reportSummaryComments = needsAndOpportunitiesTemplate.failedRemovedFormattingFromCellsComment;
          throw new Error(`Formatting not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
        }
    
    needsAndOpportunitiesTemplate.clearSheetSummaryColumn(needsAndOpportunitiesTemplate.sheet);
    needsAndOpportunitiesTemplate.clearSheetSummaryColumn(reportSheet);
    needsAndOpportunitiesTemplate.setErrorColumnHeaderInMainSheet(needsAndOpportunitiesTemplate.sheet);

    try{
      needsAndOpportunitiesTemplate.checkFirstTwoColumnsForBlanks(reportSheet);
        } catch(err) {
            Logger.log(err);
            needsAndOpportunitiesTemplate.reportSummaryComments = needsAndOpportunitiesTemplate._failedFirstTwoColumnsMissingValuesCheckMessage;
            throw new Error(`Check not ran for missing values within first two columns. Reason: ${err.name}: ${err.message}.  Please revert sheet to previous version, ensure the correct values are within the first four colums (i.e. "First Name, Last Name, and Email" is running a User Import Check [this would be for first three columns, but this checker is for four]), and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
        }

    try {
    needsAndOpportunitiesTemplate.convertStatesToTwoLetterCode(reportSheet);
  } catch (err) {
    Logger.log(err);
    needsAndOpportunitiesTemplate.reportSummaryComments = needsAndOpportunitiesTemplate.failedCheckNotRanForConvertingStatesToTwoLetterCodes ;
    throw new Error(`Check not ran for converting states to two-letter code: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }

  try {
    needsAndOpportunitiesTemplate.validateStates(reportSheet);
  } catch (err) {
    Logger.log(err);
    needsAndOpportunitiesTemplate.reportSummaryComments = needsAndOpportunitiesTemplate.failedInvalidStatesCheckMessage;
    throw new Error(`Check not ran for invalid states: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  try {
    needsAndOpportunitiesTemplate.formatAllDatedColumns(reportSheet);
  } catch(err) {
    Logger.log(err);
    needsAndOpportunitiesTemplate.reportSummaryComments = needsAndOpportunitiesTemplate.failedFormatDateColumns;
    throw new Error(`Check not ran for formatting of dated columns. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }

  try {
    needsAndOpportunitiesTemplate.checkForInvalidEmails(reportSheet);
  } catch(err) {
      Logger.log(err);
      needsAndOpportunitiesTemplate.reportSummaryComments = needsAndOpportunitiesTemplate.failedInvalidEmailCheckMessage;
      throw new Error(`Check not ran for invalid emails. Reason: ${err.name}: ${err.message} at line. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
    } 

    try {
      needsAndOpportunitiesTemplate.formatCommaSeparatedLists(reportSheet);
    } catch (err) {
        Logger.log(err);
        needsAndOpportunitiesTemplate.reportSummaryComments = needsAndOpportunitiesTemplate.failedRemoveWhiteSpaceAndMissingValuesFromCommaSeparatedListsMessage;
        throw new Error(`Did not remove whitespace and empty values from comma-separated list(s). Reason: ${err.name}: ${err.message} at line. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }

    try {
    needsAndOpportunitiesTemplate.checkForInvalidCommaSeparatedEmails(reportSheet);
    } catch(err) {
        Logger.log(err);
        needsAndOpportunitiesTemplate.reportSummaryComments = needsAndOpportunitiesTemplate.failedCheckNotRanForInvalidCommaSeparatedEmails;
        throw new Error(`Check not ran for invalid need/opportunity contact emails. Reason: ${err.name}: ${err.message} at line. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
      } 
  
  try {
    needsAndOpportunitiesTemplate.checkForInvalidPostalCodes(reportSheet);
  } catch(err) {
    Logger.log(err);
    needsAndOpportunitiesTemplate.reportSummaryComments = needsAndOpportunitiesTemplate.failedInvalidPostalCodeCheck;
    throw new Error(`Check not ran for invalid postal codes. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  try {
    needsAndOpportunitiesTemplate.checkForInvalidNumbers(reportSheet);
  } catch(err) {
    Logger.log(err);
    needsAndOpportunitiesTemplate.reportSummaryComments = needsAndOpportunitiesTemplate.failedInvalidPhoneNumbersCheck;
    throw new Error(`Check not ran for invalid home or mobile phone numbers. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  
  try {
    needsAndOpportunitiesTemplate.setCommentsOnReportCell(reportSheet);
  } catch(err) {
    Logger.log(err);
    throw new Error(`Report sheet cell comment not added for summary of checks. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }

   SpreadsheetApp.getUi().alert("Neeeds/Opportunities Import Check Complete");
  }

} catch (err) {
    Logger.log(err);
    throw new Error(`An error occured the the Needs/Opportunities import template check did not successfully run. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
}
